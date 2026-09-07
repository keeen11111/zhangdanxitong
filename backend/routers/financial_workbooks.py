"""Company-agnostic financial workbook integration endpoints.

This router deliberately has no dependency on the payroll entity pipeline.  It
only joins a user-selected generic financial master with source update files,
then records the latest result for that project.
"""
from __future__ import annotations

import hashlib
import io
import json
import logging
import os
import re
import shutil
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote
from uuid import uuid4

import openpyxl
from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from sqlalchemy.orm import Session

from backend.auth import get_current_user
from backend.database import DATA_DIR, EXPORT_DIR, UPLOAD_DIR, get_db
from backend.models import Project, UploadFile, User
from backend.schemas import (
    FinancialWorkbookIntegrationOut,
    FinancialWorkbookPublishedVersionOut,
    FinancialWorkQueueOut,
)
from core.sheet_mapper import apply_semantic_sheet_updates


router = APIRouter(prefix="/api/financial-workbooks", tags=["financial-workbooks"])
logger = logging.getLogger(__name__)

_META_DIR = Path(DATA_DIR) / "financial-workbook-integrations"
_MASTER_FILE_TYPE = "financial_master"
_SOURCE_FILE_TYPE = "financial_source"
_RESULT_META_NAME = "latest.json"
_PROGRESS_META_NAME = "progress.json"
_SESSION_DIR = Path(DATA_DIR) / "sessions"
_PUBLISHED_DIR_NAME = "published"
_XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_ATOMIC_WRITE_MAX_ATTEMPTS = 5
_ATOMIC_WRITE_RETRY_SECONDS = 0.05
_PROGRESS_STAGES = (
    ("prepare", "准备总表和来源文件"),
    ("match", "核对工作表与业务主键"),
    ("update", "安全写入总表数据"),
    ("export", "生成正式稿和修订稿"),
)


def _load_project_or_404(project_id: str, user: User, db: Session) -> Project:
    project = db.query(Project).filter(Project.id == project_id).first()
    if not project:
        raise HTTPException(status_code=404, detail="项目不存在")
    if project.owner.tenant_id != user.tenant_id:
        raise HTTPException(status_code=403, detail="无权访问")
    return project


def _uploaded_file_path(uploaded_file: UploadFile) -> Path:
    """Resolve a database-backed upload only when it remains under UPLOAD_DIR."""
    upload_root = Path(UPLOAD_DIR).resolve()
    candidate = (upload_root / str(uploaded_file.stored_path)).resolve()
    if candidate != upload_root and upload_root not in candidate.parents:
        raise HTTPException(status_code=500, detail="上传文件路径无效")
    if not candidate.is_file():
        raise HTTPException(status_code=404, detail=f"上传文件已丢失：{uploaded_file.original_name}")
    return candidate


def _result_meta_path(project_id: str) -> Path:
    return _META_DIR / project_id / _RESULT_META_NAME


def _progress_meta_path(project_id: str) -> Path:
    return _META_DIR / project_id / _PROGRESS_META_NAME


def _result_path(project_id: str, filename: str) -> Path:
    output_root = (Path(EXPORT_DIR) / project_id).resolve()
    candidate = (output_root / filename).resolve()
    if output_root not in candidate.parents:
        raise HTTPException(status_code=500, detail="导出文件路径无效")
    return candidate


def _load_matching_policy(project_id: str) -> dict[str, Any] | None:
    path = _SESSION_DIR / f"{project_id}_matching_policy.json"
    if not path.is_file():
        return None
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def _published_meta_dir(project_id: str) -> Path:
    return _META_DIR / project_id / _PUBLISHED_DIR_NAME


def _published_meta_path(project_id: str, version_id: str) -> Path:
    if not re.fullmatch(r"[0-9a-f]{32}", version_id):
        raise HTTPException(status_code=404, detail="正式账单版本不存在")
    return _published_meta_dir(project_id) / f"{version_id}.json"


def _published_result_path(project_id: str, filename: str) -> Path:
    output_root = (Path(EXPORT_DIR) / project_id / _PUBLISHED_DIR_NAME).resolve()
    candidate = (output_root / filename).resolve()
    if output_root not in candidate.parents:
        raise HTTPException(status_code=500, detail="正式账单归档路径无效")
    return candidate


def _input_signature(files: list[UploadFile]) -> str:
    records = [
        {
            "id": str(file.id),
            "file_type": str(file.file_type),
            "stored_path": str(file.stored_path),
            "created_at": file.created_at.isoformat() if file.created_at else "",
        }
        for file in files
    ]
    payload = json.dumps(sorted(records, key=lambda item: item["id"]), ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _select_integration_files(project_id: str, db: Session) -> tuple[UploadFile, list[UploadFile]]:
    files = db.query(UploadFile).filter(
        UploadFile.project_id == project_id,
        UploadFile.file_type.in_([_MASTER_FILE_TYPE, _SOURCE_FILE_TYPE]),
    ).all()
    masters = [file for file in files if file.file_type == _MASTER_FILE_TYPE]
    sources = [file for file in files if file.file_type == _SOURCE_FILE_TYPE]
    if len(masters) != 1:
        raise HTTPException(status_code=400, detail="请先且仅保留一份通用财务总表")
    if not sources:
        raise HTTPException(status_code=400, detail="请至少上传一份来源更新文件")
    return masters[0], sources


def _release_state(result: dict[str, Any]) -> tuple[str, dict[str, Any]]:
    """Classify a generated workbook without ever auto-publishing it."""
    issue_count = len(result.get("issues", []))
    matched_sheet_count = len(result.get("matches", []))
    checks = {
        "can_release": issue_count == 0 and matched_sheet_count > 0,
        "issue_count": issue_count,
        "matched_sheet_count": matched_sheet_count,
        "auto_update_count": int(result.get("auto_update_count", 0)),
    }
    if result.get("personnel_coverage", {}).get("input_person_count"):
        checks["personnel_coverage"] = result["personnel_coverage"]
        checks["can_release"] = checks["can_release"] and result["personnel_coverage"]["unprocessed_person_count"] == 0
    if issue_count:
        return "review_required", checks
    if not matched_sheet_count:
        return "blocked", checks
    if not checks["can_release"]:
        return "blocked", checks
    return "ready_for_release", checks


def _to_response(meta: dict[str, Any]) -> FinancialWorkbookIntegrationOut:
    return FinancialWorkbookIntegrationOut(
        filename=str(meta["filename"]),
        review_filename=str(meta["review_filename"]) if meta.get("review_filename") else None,
        auto_update_count=int(meta["auto_update_count"]),
        source_file_count=int(meta["source_file_count"]),
        issues=list(meta.get("issues", [])),
        matches=list(meta.get("matches", [])),
        updates=list(meta.get("updates", [])),
        status=str(meta["status"]),
        release_checks=dict(meta["release_checks"]),
        completed_at=str(meta["completed_at"]),
        published_at=str(meta["published_at"]) if meta.get("published_at") else None,
        published_version=_to_published_version(meta["published_version"])
        if meta.get("published_version")
        else None,
    )


def _to_published_version(meta: dict[str, Any]) -> FinancialWorkbookPublishedVersionOut:
    return FinancialWorkbookPublishedVersionOut(
        version_id=str(meta["version_id"]),
        filename=str(meta["filename"]),
        original_draft_filename=str(meta["original_draft_filename"]),
        auto_update_count=int(meta["auto_update_count"]),
        source_file_count=int(meta["source_file_count"]),
        release_checks=dict(meta["release_checks"]),
        published_at=str(meta["published_at"]),
        published_by={key: str(value) for key, value in dict(meta["published_by"]).items()},
        sha256=str(meta["sha256"]),
    )


def _write_result_meta(project_id: str, meta: dict[str, Any]) -> None:
    _write_json_atomically(_result_meta_path(project_id), meta)


def _write_json_atomically(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(f".{uuid4().hex}.tmp")
    try:
        temporary_path.write_text(json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8")
        for attempt in range(_ATOMIC_WRITE_MAX_ATTEMPTS):
            try:
                os.replace(temporary_path, path)
                return
            except PermissionError:
                if attempt == _ATOMIC_WRITE_MAX_ATTEMPTS - 1:
                    raise
                time.sleep(_ATOMIC_WRITE_RETRY_SECONDS * (attempt + 1))
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _copy_file_atomically(source_path: Path, destination_path: Path) -> None:
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = destination_path.with_suffix(f".{uuid4().hex}.tmp")
    try:
        shutil.copyfile(source_path, temporary_path)
        os.replace(temporary_path, destination_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _progress_stages(active_index: int | None = None, *, completed: bool = False) -> list[dict[str, str]]:
    stages: list[dict[str, str]] = []
    for index, (key, label) in enumerate(_PROGRESS_STAGES):
        status = "completed" if completed or (active_index is not None and index < active_index) else "pending"
        if active_index == index and not completed:
            status = "running"
        stages.append({"key": key, "label": label, "status": status})
    return stages


def _write_progress(
    project_id: str,
    *,
    status: str,
    stages: list[dict[str, str]],
    detail: str,
) -> dict[str, Any]:
    progress = {
        "status": status,
        "stages": stages,
        "detail": detail,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    try:
        _write_json_atomically(_progress_meta_path(project_id), progress)
    except PermissionError:
        # Windows antivirus, sync clients, or a previous request can briefly lock
        # the progress file. Progress is informational; it must not abort an
        # otherwise valid workbook integration.
        logger.warning("Unable to persist financial workbook progress for project %s", project_id)
    return progress


def _load_progress(project_id: str) -> dict[str, Any]:
    path = _progress_meta_path(project_id)
    if not path.is_file():
        return {
            "status": "idle",
            "stages": _progress_stages(),
            "detail": "等待开始整合",
            "updated_at": "",
        }
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise HTTPException(status_code=500, detail="通用财务整合进度信息损坏") from error


def _create_review_workbook(
    final_path: Path,
    review_path: Path,
    updates: list[dict[str, Any]],
) -> None:
    """Create a value-identical review copy with comments and a change ledger."""
    _copy_file_atomically(final_path, review_path)
    workbook = openpyxl.load_workbook(review_path, data_only=False)
    try:
        audit_title = "修改记录"
        suffix = 2
        while audit_title in workbook.sheetnames:
            audit_title = f"修改记录({suffix})"
            suffix += 1
        audit = workbook.create_sheet(audit_title)
        audit.append(["序号", "工作表", "单元格", "原值", "新值", "来源文件", "来源工作表", "来源行"])
        for index, update in enumerate(updates, start=1):
            sheet_name = str(update.get("target_sheet") or "")
            coordinate = str(update.get("target_cell") or "")
            if sheet_name in workbook.sheetnames and coordinate:
                cell = workbook[sheet_name][coordinate]
                cell.comment = Comment(
                    "本次整合更新\n"
                    f"位置：{sheet_name}!{coordinate}\n"
                    f"原值：{update.get('old_value', '')}\n"
                    f"新值：{update.get('new_value', '')}\n"
                    f"来源：{update.get('source_file', '')} / {update.get('source_sheet', '')}",
                    "外服账单系统",
                )
            audit.append([
                index,
                sheet_name,
                coordinate,
                update.get("old_value", ""),
                update.get("new_value", ""),
                update.get("source_file", ""),
                update.get("source_sheet", ""),
                update.get("source_row", ""),
            ])
        header_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in audit[1]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        audit.freeze_panes = "A2"
        audit.auto_filter.ref = f"A1:H{max(1, audit.max_row)}"
        workbook.save(review_path)
    finally:
        workbook.close()


def _load_published_version_meta(project_id: str, version_id: str) -> dict[str, Any]:
    meta_path = _published_meta_path(project_id, version_id)
    if not meta_path.is_file():
        raise HTTPException(status_code=404, detail="正式账单版本不存在")
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise HTTPException(status_code=500, detail="正式账单版本元数据损坏") from error
    if str(meta.get("version_id")) != version_id:
        raise HTTPException(status_code=500, detail="正式账单版本元数据不一致")
    archive_path = _published_result_path(project_id, str(meta.get("filename", "")))
    if not archive_path.is_file():
        raise HTTPException(status_code=404, detail="正式账单归档文件已丢失")
    expected_sha256 = str(meta.get("sha256", ""))
    if not re.fullmatch(r"[0-9a-f]{64}", expected_sha256):
        raise HTTPException(status_code=500, detail="正式账单版本校验信息损坏")
    if _file_sha256(archive_path) != expected_sha256:
        raise HTTPException(status_code=409, detail="正式账单归档文件校验不一致，已拒绝下载")
    return meta


def _archive_published_version(project_id: str, draft_meta: dict[str, Any], user: User) -> dict[str, Any]:
    """Copy the final workbook before marking the mutable latest result as published."""
    version_id = uuid4().hex
    published_at = datetime.now(timezone.utc).isoformat()
    filename = f"正式账单_{version_id}.xlsx"
    archive_meta = {
        "version_id": version_id,
        "filename": filename,
        "original_draft_filename": str(draft_meta["filename"]),
        "auto_update_count": int(draft_meta["auto_update_count"]),
        "source_file_count": int(draft_meta["source_file_count"]),
        "release_checks": dict(draft_meta["release_checks"]),
        "published_at": published_at,
        "published_by": {"id": str(user.id), "name": str(user.name)},
    }
    archive_path = _published_result_path(project_id, filename)
    _copy_file_atomically(_result_path(project_id, str(draft_meta["filename"])), archive_path)
    archive_meta["sha256"] = _file_sha256(archive_path)
    _write_json_atomically(_published_meta_path(project_id, version_id), archive_meta)
    return archive_meta


def _load_latest_result(project_id: str, db: Session) -> dict[str, Any]:
    meta_path = _result_meta_path(project_id)
    if not meta_path.is_file():
        raise HTTPException(status_code=404, detail="尚未生成通用财务整合结果")
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise HTTPException(status_code=500, detail="通用财务整合结果元数据损坏") from error

    master, sources = _select_integration_files(project_id, db)
    if meta.get("input_signature") != _input_signature([master, *sources]):
        raise HTTPException(status_code=409, detail="总表或来源文件已变化，请重新执行通用财务整合")
    output_path = _result_path(project_id, str(meta.get("filename", "")))
    if not output_path.is_file():
        raise HTTPException(status_code=404, detail="通用财务整合结果文件已丢失")
    if "status" not in meta or "release_checks" not in meta:
        status, checks = _release_state(meta)
        meta["status"] = status
        meta["release_checks"] = checks
    return meta


def _build_financial_work_queue(
    projects: list[Project],
    files_by_project: dict[str, list[UploadFile]] | None = None,
) -> list[dict[str, Any]]:
    """Expose only batches where a user has a concrete next action."""
    action_by_status = {
        "review_required": lambda issue_count: f"查看 {issue_count} 项问题并修正来源",
        "ready_for_release": lambda _issue_count: "确认发布正式账单",
        "blocked": lambda _issue_count: "检查文件并重新整合",
    }
    priority_by_status = {"review_required": 0, "blocked": 1, "ready_for_release": 2}
    items: list[dict[str, Any]] = []
    for project in projects:
        meta_path = _result_meta_path(str(project.id))
        if not meta_path.is_file():
            continue
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            status = "blocked"
            issue_count = 0
            completed_at = ""
            action_label = "检查损坏的整合结果"
        else:
            status = str(meta.get("status", "blocked"))
            if status not in action_by_status:
                continue
            issue_count = len(meta.get("issues", []))
            completed_at = str(meta.get("completed_at", ""))
            action_label = action_by_status[status](issue_count)
            project_files = (files_by_project or {}).get(str(project.id))
            if (
                project_files is not None
                and meta.get("input_signature")
                and meta["input_signature"] != _input_signature(project_files)
            ):
                status = "blocked"
                issue_count = 0
                action_label = "文件已变化，请重新整合"
        items.append({
            "project_id": str(project.id),
            "project_name": str(project.name),
            "salary_month": str(project.salary_month),
            "status": status,
            "issue_count": issue_count,
            "action_label": action_label,
            "completed_at": completed_at,
        })
    # A user should see the newest item first within each urgency bucket.
    # Python's stable sort lets us express that without parsing partial legacy timestamps.
    items.sort(key=lambda item: str(item["completed_at"]), reverse=True)
    items.sort(key=lambda item: priority_by_status.get(str(item["status"]), 3))
    return items


def _content_disposition(filename: str) -> str:
    return f"attachment; filename*=UTF-8''{quote(filename)}"


def _build_preview(path: Path, max_rows: int = 20, max_columns: int = 16) -> dict[str, Any]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        rows: list[list[Any]] = []
        for values in sheet.iter_rows(max_row=max_rows, max_col=max_columns, values_only=True):
            row = [value.isoformat() if hasattr(value, "isoformat") else value for value in values]
            while row and row[-1] in (None, ""):
                row.pop()
            if any(value not in (None, "") for value in row):
                rows.append(row)
        return {"sheet_name": sheet.title, "rows": rows}
    finally:
        workbook.close()


@router.get("/work-queue", response_model=FinancialWorkQueueOut)
def list_financial_work_queue(
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=100),
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """List the current user's batches that need a human action, in priority order."""
    projects = (
        db.query(Project)
        .join(User, Project.owner_id == User.id)
        .filter(User.tenant_id == user.tenant_id)
        .order_by(Project.updated_at.desc())
        .all()
    )
    project_ids = [project.id for project in projects]
    files_by_project: dict[str, list[UploadFile]] = {str(project_id): [] for project_id in project_ids}
    if project_ids:
        files = db.query(UploadFile).filter(
            UploadFile.project_id.in_(project_ids),
            UploadFile.file_type.in_([_MASTER_FILE_TYPE, _SOURCE_FILE_TYPE]),
        ).all()
        for file in files:
            files_by_project.setdefault(str(file.project_id), []).append(file)
    items = _build_financial_work_queue(projects, files_by_project)
    start = (page - 1) * page_size
    return {
        "items": items[start:start + page_size],
        "total": len(items),
        "page": page,
        "page_size": page_size,
    }


@router.post("/{project_id}/integrations", response_model=FinancialWorkbookIntegrationOut, status_code=201)
def create_financial_workbook_integration(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
    sheet_overrides: dict[tuple[str, str], str] | None = None,
):
    """Create the latest safe, content-driven integration result for one project."""
    project = _load_project_or_404(project_id, user, db)
    current_stages = _progress_stages(0)
    _write_progress(
        project.id,
        status="processing",
        stages=current_stages,
        detail="正在准备总表和来源文件",
    )
    try:
        master, sources = _select_integration_files(project.id, db)
        master_path = _uploaded_file_path(master)
        source_paths = [_uploaded_file_path(source) for source in sources]

        current_stages = _progress_stages(1)
        _write_progress(
            project.id,
            status="processing",
            stages=current_stages,
            detail=f"正在核对 {len(sources)} 份来源文件的工作表与业务主键",
        )
        result_id = uuid4().hex[:12]
        filename = f"财务整合正式稿_{result_id}.xlsx"
        review_filename = f"财务整合修订稿_{result_id}.xlsx"
        output_path = _result_path(project.id, filename)
        review_path = _result_path(project.id, review_filename)
        matching_policy = _load_matching_policy(project.id) or {
            "auto_match_threshold": 0.85,
            "field_weights": {"姓名": 0.50, "公司": 0.25, "部门": 0.15, "岗位": 0.10},
            "source_rules": [],
        }
        integration_kwargs = {
            "sheet_overrides": sheet_overrides,
            "matching_policy": matching_policy,
            "salary_month": getattr(project, "salary_month", None),
        }
        if sheet_overrides:
            # Overrides are keyed by the accountant-visible original names;
            # preserve the historical stored-basename response when no
            # override is involved.
            integration_kwargs["source_names"] = {
                str(path): str(source.original_name) for path, source in zip(source_paths, sources)
            }
        result = apply_semantic_sheet_updates(master_path, source_paths, output_path, **integration_kwargs)
        if result.get('structured_hire') and not result['issues']:
            from backend.workbook_calculation import recalculate_personnel_workbook
            result['calculation'] = recalculate_personnel_workbook(master_path, output_path, result.get('row_insertions', []))

        current_stages = _progress_stages(3)
        _write_progress(
            project.id,
            status="processing",
            stages=current_stages,
            detail="总表数据已更新，正在生成正式稿和修订稿",
        )
        _create_review_workbook(output_path, review_path, result["updates"])
        if result.get('calculation'):
            from backend.workbook_calculation import preserve_review_formula_caches
            preserve_review_formula_caches(output_path, review_path)
        status, release_checks = _release_state(result)
        meta = {
            "filename": filename,
            "review_filename": review_filename,
            "input_signature": _input_signature([master, *sources]),
            "auto_update_count": result["auto_update_count"],
            "source_file_count": len(sources),
            "issues": result["issues"],
            "matches": result["matches"],
            "updates": result["updates"],
            "status": status,
            "release_checks": release_checks,
            "completed_at": datetime.now(timezone.utc).isoformat(),
            "published_at": None,
            "sheet_mappings": [
                {"source_file": source_file, "source_sheet": source_sheet, "target_sheet": target_sheet}
                for (source_file, source_sheet), target_sheet in (sheet_overrides or {}).items()
            ],
        }
        _write_result_meta(project.id, meta)
        completed_status = "completed_with_issues" if result["issues"] else "completed"
        _write_progress(
            project.id,
            status=completed_status,
            stages=_progress_stages(completed=True),
            detail="整合完毕，可以查看结果",
        )
        return _to_response(meta)
    except Exception as error:
        _write_progress(
            project.id,
            status="failed",
            stages=current_stages,
            detail=str(error) if isinstance(error, HTTPException) else "整合处理失败，请检查文件后重试",
        )
        raise


@router.get("/{project_id}/integrations/progress")
def get_financial_workbook_integration_progress(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    return _load_progress(project.id)


@router.get("/{project_id}/integrations/latest", response_model=FinancialWorkbookIntegrationOut)
def get_latest_financial_workbook_integration(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    return _to_response(_load_latest_result(project.id, db))


@router.post("/{project_id}/integrations/latest/release", response_model=FinancialWorkbookIntegrationOut)
def release_latest_financial_workbook_integration(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """Publish a draft only after the caller explicitly approves a clean result."""
    project = _load_project_or_404(project_id, user, db)
    meta = _load_latest_result(project.id, db)
    if meta["status"] == "published":
        if not meta.get("published_version"):
            raise HTTPException(status_code=409, detail="历史正式账单未归档，请重新生成草稿后发布")
        return _to_response(meta)
    if meta["status"] != "ready_for_release":
        raise HTTPException(status_code=409, detail="当前草稿未通过发布校验，不能发布正式账单")
    published_version = _archive_published_version(project.id, meta, user)
    meta["status"] = "published"
    meta["published_at"] = published_version["published_at"]
    meta["published_version"] = published_version
    _write_result_meta(project.id, meta)
    return _to_response(meta)


@router.get(
    "/{project_id}/integrations/published",
    response_model=list[FinancialWorkbookPublishedVersionOut],
)
def list_published_financial_workbook_versions(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """List archived formal versions without consulting mutable upload inputs."""
    project = _load_project_or_404(project_id, user, db)
    archive_dir = _published_meta_dir(project.id)
    if not archive_dir.is_dir():
        return []
    versions = [_load_published_version_meta(project.id, path.stem) for path in archive_dir.glob("*.json")]
    versions.sort(key=lambda version: str(version["published_at"]), reverse=True)
    return [_to_published_version(version) for version in versions]


@router.get("/{project_id}/integrations/latest/preview")
def preview_latest_financial_workbook_integration(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    meta = _load_latest_result(project.id, db)
    return {"filename": meta["filename"], **_build_preview(_result_path(project.id, meta["filename"]))}


def _download_workbook(project_id: str, meta: dict[str, Any], filename: str) -> StreamingResponse:
    data = _result_path(project_id, meta["filename"]).read_bytes()
    return StreamingResponse(
        io.BytesIO(data),
        media_type=_XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": _content_disposition(filename)},
    )


def _download_published_workbook(project_id: str, version: dict[str, Any]) -> StreamingResponse:
    data = _published_result_path(project_id, str(version["filename"])).read_bytes()
    return StreamingResponse(
        io.BytesIO(data),
        media_type=_XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": _content_disposition(str(version["filename"]))},
    )


@router.get("/{project_id}/integrations/latest/draft/download")
def download_latest_financial_workbook_draft(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    meta = _load_latest_result(project.id, db)
    return _download_workbook(project.id, meta, f"草稿_{meta['filename']}")


@router.get("/{project_id}/integrations/latest/download")
def download_latest_financial_workbook_integration(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    meta = _load_latest_result(project.id, db)
    published_version = meta.get("published_version")
    if meta.get("status") == "published" and published_version:
        version = _load_published_version_meta(project.id, str(published_version["version_id"]))
        return _download_published_workbook(project.id, version)
    return _download_workbook(project.id, meta, f"正式稿_{meta['filename']}")


@router.get("/{project_id}/integrations/latest/review/download")
def download_latest_financial_workbook_review(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    meta = _load_latest_result(project.id, db)
    review_filename = str(meta.get("review_filename") or "")
    review_path = _result_path(project.id, review_filename)
    if not review_filename or not review_path.is_file():
        raise HTTPException(status_code=404, detail="当前历史结果没有修订稿，请重新整合后生成")
    data = review_path.read_bytes()
    return StreamingResponse(
        io.BytesIO(data),
        media_type=_XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": _content_disposition(f"修订稿_{review_filename}")},
    )


@router.get("/{project_id}/integrations/published/{version_id}/download")
def download_published_financial_workbook_version(
    project_id: str,
    version_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    project = _load_project_or_404(project_id, user, db)
    return _download_published_workbook(
        project.id,
        _load_published_version_meta(project.id, version_id),
    )
