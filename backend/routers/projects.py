"""项目空间路由：CRUD + 文件上传/预览。"""
import json
import os
import subprocess
import tempfile
import time
import uuid
from datetime import datetime
from io import BytesIO
from zipfile import BadZipFile, ZipFile

import openpyxl
import pandas as pd
try:
    import xlrd
except ImportError:  # pragma: no cover - optional fallback for minimal installs
    xlrd = None
from msoffcrypto import OfficeFile
from msoffcrypto.exceptions import DecryptionError, FileFormatError, InvalidKeyError, ParseError
from fastapi import APIRouter, Depends, HTTPException, UploadFile as FastUploadFile, File, Form, Query
from sqlalchemy.orm import Session

from backend.auth import get_current_user, require_tenant_match
from backend.database import DATA_DIR, get_db, UPLOAD_DIR
from backend.models import User, Project, UploadFile
from backend.schemas import ProjectIn, ProjectOut, FileOut, FilePreview, FileTypeUpdateIn

router = APIRouter(prefix="/api/projects", tags=["projects"])
SESSION_DIR = os.path.join(DATA_DIR, "sessions")

# 允许的文件类型
_ALLOWED_TYPES = {
    "last_month",
    "current",
    "social",
    "template",
    "financial_master",
    "financial_source",
    "source",
}


def _project_result_summary(project_id: str) -> dict[str, object]:
    """Return the latest usable result state for either supported flow.

    The project row's legacy ``status`` only represents the import wizard. A
    completed integration is stored separately so historic projects can open
    their existing result instead of restarting the import flow.
    """
    empty = {
        "has_result": False,
        "result_completed_at": None,
        "pending_issue_count": 0,
    }
    meta_path = os.path.join(SESSION_DIR, f"{project_id}_export_meta.json")
    try:
        with open(meta_path, "r", encoding="utf-8") as fp:
            meta = json.load(fp)
    except (OSError, ValueError, TypeError):
        meta = None

    if isinstance(meta, dict) and meta.get("status") not in {"processing", "blocked", "stale"} and meta.get("filename"):
        completed_at = meta.get("completed_at")
        try:
            pending_issue_count = max(0, int(meta.get("issue_count", 0) or 0))
        except (TypeError, ValueError):
            pending_issue_count = 0
        return {
            "has_result": True,
            "result_completed_at": completed_at if isinstance(completed_at, str) and completed_at else None,
            "pending_issue_count": pending_issue_count,
        }

    financial_meta_path = os.path.join(
        DATA_DIR,
        "financial-workbook-integrations",
        project_id,
        "latest.json",
    )
    try:
        with open(financial_meta_path, "r", encoding="utf-8") as fp:
            financial_meta = json.load(fp)
    except (OSError, ValueError, TypeError):
        return empty

    if not isinstance(financial_meta, dict):
        return empty
    if financial_meta.get("status") not in {"review_required", "ready_for_release", "published"}:
        return empty
    if not financial_meta.get("filename"):
        return empty

    completed_at = financial_meta.get("completed_at")
    issues = financial_meta.get("issues", [])
    return {
        "has_result": True,
        "result_completed_at": completed_at if isinstance(completed_at, str) and completed_at else None,
        "pending_issue_count": len(issues) if isinstance(issues, list) else 0,
    }


def _project_out(project: Project) -> ProjectOut:
    """Build the project list payload with the latest integration result state."""
    summary = _project_result_summary(project.id)
    output = ProjectOut.from_orm(project)
    output.file_count = len(project.files)
    output.has_result = bool(summary["has_result"])
    completed_at = summary["result_completed_at"]
    if isinstance(completed_at, str):
        try:
            output.result_completed_at = datetime.fromisoformat(completed_at.replace("Z", "+00:00"))
        except ValueError:
            output.result_completed_at = None
    output.pending_issue_count = int(summary["pending_issue_count"])
    return output
_SINGLE_MASTER_FILE_TYPES = {"template", "financial_master"}

_OLE_WORKBOOK_SIGNATURE = bytes.fromhex("D0 CF 11 E0 A1 B1 1A E1")
_OOXML_WORKBOOK_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
_ENCRYPTED_OOXML_STREAM_NAMES = (
    "EncryptionInfo".encode("utf-16le"),
    "EncryptedPackage".encode("utf-16le"),
)
_MAX_UPLOAD_BYTES = 100 * 1024 * 1024
_MAX_UNCOMPRESSED_WORKBOOK_BYTES = 1024 * 1024 * 1024
_MAX_WORKBOOK_ARCHIVE_ENTRIES = 20_000
_DESKTOP_CONVERSION_TIMEOUT_SECONDS = 60

_MASTER_SUPPORT_SHEETS = {
    "工资汇总表",
    "工资条",
    "台账",
    "个税申报",
    "个税导出核对",
    "人员异动",
    "规则备忘",
    "薪酬福利政策",
}


def _read_upload_content(upload: FastUploadFile) -> bytes:
    """Validate the declared upload shape and bound memory usage before parsing."""
    filename = upload.filename or "未命名文件.xlsx"
    if not filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail=f"{filename} 不是支持的 Excel 文件")
    content = upload.file.read(_MAX_UPLOAD_BYTES + 1)
    if len(content) > _MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="单个 Excel 文件不能超过 100 MB")
    return content


def _validate_ooxml_archive(content: bytes) -> None:
    """Reject malformed or excessively expanded OOXML archives before openpyxl."""
    try:
        with ZipFile(BytesIO(content)) as archive:
            members = archive.infolist()
            if len(members) > _MAX_WORKBOOK_ARCHIVE_ENTRIES:
                raise HTTPException(status_code=400, detail="Excel 压缩包内文件数量异常")
            if sum(member.file_size for member in members) > _MAX_UNCOMPRESSED_WORKBOOK_BYTES:
                raise HTTPException(status_code=413, detail="Excel 解压后内容过大，无法安全处理")
    except BadZipFile as exc:
        raise HTTPException(status_code=400, detail="Excel 文件结构已损坏") from exc


def _workbook_binary_kind(content: bytes) -> str:
    """Identify Excel containers by their bytes instead of trusting the suffix."""
    if content.startswith(_OLE_WORKBOOK_SIGNATURE):
        # Password-protected OOXML workbooks are wrapped in an OLE container.
        # They must be opened as .xlsx by the desktop converter; treating them
        # as BIFF .xls makes Excel/WPS reject the input before conversion.
        if all(marker in content for marker in _ENCRYPTED_OOXML_STREAM_NAMES):
            return "xlsx-encrypted"
        return "xls"
    if content.startswith(_OOXML_WORKBOOK_SIGNATURES):
        return "xlsx"
    return "unknown"


def _convert_legacy_excel_with_desktop_excel(
    content: bytes,
    source_extension: str = ".xls",
) -> bytes:
    """Convert an OLE Excel container to OOXML while preserving workbook layout.

    OLE bytes can represent either BIFF ``.xls`` or an encrypted OOXML package.
    The Windows desktop spreadsheet application can safely normalize both while
    preserving the uploaded master workbook's layout.
    """
    if source_extension not in {".xls", ".xlsx"}:
        raise ValueError(f"不支持的 Excel 源扩展名：{source_extension}")
    source_format_label = "加密的 .xlsx" if source_extension == ".xlsx" else "旧版 .xls"

    # A desktop spreadsheet process may briefly keep the source file open even
    # after a timeout.  Cleanup must never mask the actionable HTTP error.
    with tempfile.TemporaryDirectory(
        prefix="payroll_excel_",
        ignore_cleanup_errors=True,
    ) as temp_dir:
        source_path = os.path.join(temp_dir, f"source{source_extension}")
        output_path = os.path.join(temp_dir, "converted.xlsx")
        with open(source_path, "wb") as source:
            source.write(content)

        def ps_literal(value: str) -> str:
            return "'" + value.replace("'", "''") + "'"

        script = f"""
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {{
  foreach ($programId in @('Excel.Application', 'KET.Application')) {{
    try {{
      $excel = New-Object -ComObject $programId
      break
    }} catch {{}}
  }}
  if ($null -eq $excel) {{ throw 'No compatible desktop spreadsheet application is registered' }}
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  try {{ $excel.AutomationSecurity = 3 }} catch {{}}
  try {{ $excel.AskToUpdateLinks = $false }} catch {{}}
  # Pass an explicit empty password and disable recommendations so an encrypted
  # legacy workbook fails instead of opening an invisible password dialog.
  $workbook = $excel.Workbooks.Open({ps_literal(source_path)}, 0, $false, 5, '', '', $true)
  # OLE input may be a true legacy workbook or a password-encrypted OOXML
  # package.  The application has already opened it in the current desktop
  # session; clear file-level encryption so the backend can inspect and update
  # the local working copy.
  $workbook.Password = ''
  $workbook.WritePassword = ''
  $workbook.SaveAs({ps_literal(output_path)}, 51)
}} finally {{
  if ($null -ne $workbook) {{ $workbook.Close($false) }}
  if ($null -ne $excel) {{ $excel.Quit() }}
}}
"""
        process = None
        try:
            process = subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            stdout, stderr = process.communicate(timeout=_DESKTOP_CONVERSION_TIMEOUT_SECONDS)
            completed = subprocess.CompletedProcess(
                process.args,
                process.returncode,
                stdout=stdout,
                stderr=stderr,
            )
        except subprocess.TimeoutExpired as exc:
            if process is not None:
                try:
                    process.kill()
                except OSError:
                    pass
                try:
                    process.communicate(timeout=5)
                except (OSError, subprocess.TimeoutExpired):
                    pass
            raise HTTPException(
                status_code=408,
                detail=(
                    "旧版 .xls 总表需要 Excel/WPS 自动转换，但 60 秒内未完成；"
                    "请关闭占用该文件的窗口，或另存为 .xlsx 后重试"
                ),
            ) from exc
        except OSError as exc:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"检测到{source_format_label}内容，但无法调用本机 Excel/WPS 转换；"
                    "请在表格软件中另存为未加密的 .xlsx 后重试"
                ),
            ) from exc
        if completed.returncode != 0 or not os.path.exists(output_path):
            raise HTTPException(
                status_code=400,
                detail=(
                    f"检测到{source_format_label}内容，但自动转换失败；"
                    "请关闭占用文件的 Excel/WPS 窗口，并另存为未加密的 .xlsx 后重试"
                ),
            )
        with open(output_path, "rb") as converted:
            converted_content = converted.read()
        if not converted_content.startswith(_OOXML_WORKBOOK_SIGNATURES):
            raise HTTPException(
                status_code=400,
                detail="Excel/WPS 未生成有效的 .xlsx 文件；请在表格软件中取消打开密码并另存为 .xlsx 后重试",
            )
        return converted_content


def _convert_legacy_xls_with_xlrd(content: bytes) -> bytes:
    """Normalize a readable BIFF workbook when no desktop converter is available.

    This fallback is intentionally value-only. Legacy tax attachments used by
    the payroll workflow are source evidence, so preserving their values and
    sheet names is safer than blocking the whole batch on a missing Excel COM
    installation. Desktop conversion remains the first choice because it keeps
    richer legacy formatting when available.
    """
    if xlrd is None:
        raise HTTPException(
            status_code=400,
            detail="检测到旧版 .xls 文件，但当前环境没有可用的 .xls 解析器；请安装 xlrd 或另存为 .xlsx 后重试",
        )
    try:
        legacy = xlrd.open_workbook(file_contents=content, on_demand=False)
    except Exception as exc:
        raise HTTPException(status_code=400, detail="旧版 .xls 文件无法解析，请确认文件未损坏") from exc

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    try:
        for sheet_index in range(legacy.nsheets):
            source_sheet = legacy.sheet_by_index(sheet_index)
            title = str(source_sheet.name or f"Sheet{sheet_index + 1}")[:31] or f"Sheet{sheet_index + 1}"
            # Excel sheet names must be unique after the 31-character limit.
            base_title = title
            suffix = 1
            while title in workbook.sheetnames:
                suffix += 1
                title = f"{base_title[:31 - len(str(suffix)) - 1]}_{suffix}"
            target_sheet = workbook.create_sheet(title)
            for row_index in range(source_sheet.nrows):
                for column_index in range(source_sheet.ncols):
                    cell = source_sheet.cell(row_index, column_index)
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            value = xlrd.xldate_as_datetime(value, legacy.datemode)
                        except (TypeError, ValueError, OverflowError):
                            pass
                    elif cell.ctype == xlrd.XL_CELL_ERROR:
                        # Preserve the visible error marker as text; it must
                        # never become an executable formula or an exception.
                        value = f"#ERR:{cell.value}"
                    target_sheet.cell(row=row_index + 1, column=column_index + 1, value=value)
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()
    finally:
        workbook.close()


def _decrypt_encrypted_ooxml(content: bytes, password: str) -> bytes:
    """Decrypt an OOXML workbook in memory without persisting its password."""
    if not password:
        raise HTTPException(status_code=400, detail="检测到加密的 Excel 文件，请在上传区域输入打开密码后重试")
    try:
        encrypted = OfficeFile(BytesIO(content))
        encrypted.load_key(password=password, verify_password=True)
        decrypted = BytesIO()
        encrypted.decrypt(decrypted, verify_integrity=True)
    except InvalidKeyError as exc:
        raise HTTPException(status_code=400, detail="Excel 打开密码不正确，请重新输入") from exc
    except (DecryptionError, FileFormatError, ParseError) as exc:
        raise HTTPException(status_code=400, detail="Excel 文件无法使用该密码解密，请确认文件和打开密码") from exc
    return decrypted.getvalue()


def _normalize_excel_content(content: bytes, password: str | None = None) -> tuple[bytes, str]:
    """Return an OOXML workbook and the safe storage suffix."""
    kind = _workbook_binary_kind(content)
    if kind == "xlsx":
        _validate_ooxml_archive(content)
        return content, ".xlsx"
    if kind == "xlsx-encrypted":
        decrypted = _decrypt_encrypted_ooxml(content, password or "")
        _validate_ooxml_archive(decrypted)
        return decrypted, ".xlsx"
    if kind == "xls":
        try:
            converted = _convert_legacy_excel_with_desktop_excel(content, ".xls")
        except HTTPException as desktop_error:
            # CI, portable demos, and server containers commonly lack Excel or
            # WPS. Use the read-only BIFF fallback for source workbooks.
            if desktop_error.status_code not in {400, 408}:
                raise
            converted = _convert_legacy_xls_with_xlrd(content)
        _validate_ooxml_archive(converted)
        return converted, ".xlsx"
    raise HTTPException(status_code=400, detail="文件内容不是有效的 Excel 工作簿")


def _inspect_ooxml_workbook(content: bytes) -> dict[str, object]:
    """Validate an OOXML workbook and read only structure needed at upload time."""
    try:
        workbook = openpyxl.load_workbook(BytesIO(content), read_only=True, data_only=False)
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Excel 文件无法解析，可能已损坏或仍有打开密码") from exc
    try:
        sheet_names = list(workbook.sheetnames)
        populated_dimensions: list[tuple[int, int]] = []
        for sheet in workbook.worksheets:
            if sheet.max_row is None or sheet.max_column is None:
                # Some valid OOXML producers omit the worksheet dimension.
                # Read-only openpyxl then exposes unknown bounds until it has
                # scanned the sheet once.
                try:
                    sheet.calculate_dimension(force=True)
                except (TypeError, ValueError):
                    pass
            row_count = max(0, int(sheet.max_row or 0))
            col_count = max(0, int(sheet.max_column or 0))
            first_value = sheet.cell(1, 1).value
            if row_count > 1 or col_count > 1 or first_value not in (None, ""):
                populated_dimensions.append((max(1, row_count), max(1, col_count)))
    finally:
        workbook.close()

    if populated_dimensions:
        return {
            "sheet_names": sheet_names,
            "row_count": max(rows for rows, _ in populated_dimensions),
            "col_count": max(columns for _, columns in populated_dimensions),
        }

    try:
        with ZipFile(BytesIO(content)) as archive:
            has_drawings = any(
                name.startswith(("xl/media/", "xl/drawings/", "xl/charts/"))
                for name in archive.namelist()
            )
    except BadZipFile as exc:
        raise HTTPException(status_code=400, detail="Excel 文件结构已损坏") from exc
    if not has_drawings:
        raise HTTPException(status_code=400, detail="文件为空（所有 Sheet 均无单元格数据、图片或图表）")
    return {"sheet_names": sheet_names, "row_count": 0, "col_count": 0}


def _sheet_category(sheet_name: str) -> str:
    normalized = str(sheet_name).replace(" ", "")
    rules = (
        (("工资核算",), "工资核算主表"),
        (("工资汇总",), "工资汇总"),
        (("工资条",), "工资条"),
        (("人员异动", "入离职", "转岗", "调岗", "转正"), "人员异动"),
        (("考勤", "迟到", "早退", "旷工", "值班"), "考勤与出勤"),
        (("社保", "公积金", "台账"), "社保公积金"),
        (("个税",), "个税"),
        (("绩效", "奖金", "提成", "年终奖", "调差", "福利", "通讯费", "防暑", "降温费"), "薪酬调整"),
        (("规则", "政策"), "规则政策"),
        (("审批", "请款"), "审批与请款"),
    )
    return next((category for keywords, category in rules if any(keyword in normalized for keyword in keywords)), "其他")


def analyze_workbook_structure(file_path: str) -> dict[str, object]:
    """Classify every master sheet and summarize its observable content types."""
    workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
    sheets: list[dict[str, object]] = []
    total_formulas = 0
    try:
        for sheet in workbook.worksheets:
            non_empty = 0
            formula_count = 0
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value not in (None, ""):
                        non_empty += 1
                    if cell.data_type == "f":
                        formula_count += 1
            total_formulas += formula_count
            content_types = []
            if non_empty:
                content_types.append("数据")
            if formula_count:
                content_types.append("公式")
            if not content_types:
                content_types.append("版式/附件")
            sheets.append({
                "name": sheet.title,
                "category": _sheet_category(sheet.title),
                "row_count": sheet.max_row if non_empty else 0,
                "col_count": sheet.max_column if non_empty else 0,
                "non_empty_cells": non_empty,
                "formula_count": formula_count,
                "content_types": content_types,
            })
    finally:
        workbook.close()
    return {
        "sheet_count": len(sheets),
        "formula_count": total_formulas,
        "categories": sorted({str(sheet["category"]) for sheet in sheets}),
        "sheets": sheets,
    }


def _classify_uploaded_workbook(filename: str, all_sheets: dict[str, pd.DataFrame]) -> str:
    """按工作簿结构确定模板或来源表，不使用模型推断。"""
    normalized_sheets = {str(name).replace(" ", "").lower() for name in all_sheets}
    normalized_name = os.path.splitext(filename)[0].replace(" ", "").lower()

    if any("工资核算" in name for name in normalized_sheets):
        return "template"
    if any(keyword in normalized_name for keyword in ("主模板", "导出模板", "工资核算模板")):
        return "template"
    return "source"


def _looks_like_complete_master_workbook(sheet_names: list[str]) -> bool:
    normalized = {str(name).replace(" ", "") for name in sheet_names}
    supporting_count = len(normalized & _MASTER_SUPPORT_SHEETS)
    return "工资核算" in normalized and supporting_count >= 2


def _validate_declared_workbook_role(file_type: str, sheet_names: list[str]) -> None:
    """Enforce explicit upload roles without guessing from filenames."""
    normalized = {str(name).replace(" ", "") for name in sheet_names}
    # A financial master is intentionally structure-agnostic.  Unlike the
    # legacy payroll template, it must not be identified from sheet names.
    if file_type == "financial_master":
        return
    if file_type == "template" and "工资核算" not in normalized:
        raise HTTPException(
            status_code=400,
            detail="所选总表缺少“工资核算”Sheet，请在变更文件区域上传此文件",
        )
    if file_type == "source" and _looks_like_complete_master_workbook(sheet_names):
        raise HTTPException(
            status_code=400,
            detail="检测到这是一份完整总表，请改在“上传总表”区域上传",
        )


def _workbook_dimensions(
    all_sheets: dict[str, pd.DataFrame],
    content: bytes | None = None,
) -> tuple[int, int]:
    """返回最大工作表尺寸，并接受只有图片或图表的有效工作簿。"""
    populated = [frame for frame in all_sheets.values() if not frame.empty]
    if populated:
        largest = max(populated, key=len)
        return len(largest), len(largest.columns)

    if content:
        try:
            with ZipFile(BytesIO(content)) as archive:
                names = archive.namelist()
                if any(
                    name.startswith(("xl/media/", "xl/drawings/", "xl/charts/"))
                    for name in names
                ):
                    return 0, 0
        except BadZipFile:
            pass

    raise HTTPException(status_code=400, detail="文件为空（所有 Sheet 均无单元格数据、图片或图表）")


def _load_project_or_404(project_id: str, user: User, db: Session) -> Project:
    p = db.query(Project).filter(Project.id == project_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="项目不存在")
    if p.owner_id != user.id:
        require_tenant_match(p.owner.tenant_id, user)
        # 同租户可访问，但仅 owner 自己能改？这里允许同租户访问
    return p


# ---------- 项目 CRUD ----------
@router.get("", response_model=list[ProjectOut])
def list_projects(user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    qs = db.query(Project).filter(Project.owner_id == user.id).order_by(Project.updated_at.desc())
    return [_project_out(project) for project in qs]


@router.post("", response_model=ProjectOut, status_code=201)
def create_project(payload: ProjectIn, user: User = Depends(get_current_user),
                   db: Session = Depends(get_db)):
    p = Project(owner_id=user.id, name=payload.name, salary_month=payload.salary_month)
    db.add(p)
    db.commit()
    db.refresh(p)
    return _project_out(p)


@router.get("/{project_id}", response_model=ProjectOut)
def get_project(project_id: str, user: User = Depends(get_current_user),
                db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    return _project_out(p)


@router.delete("/{project_id}", status_code=204)
def delete_project(project_id: str, user: User = Depends(get_current_user),
                   db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    # 仅 owner 可删
    if p.owner_id != user.id:
        raise HTTPException(status_code=403, detail="无权删除")
    # 删除物理文件
    for f in p.files:
        fp = os.path.join(UPLOAD_DIR, f.stored_path)
        if os.path.exists(fp):
            try:
                os.remove(fp)
            except OSError:
                pass
    db.delete(p)
    db.commit()


@router.patch("/{project_id}/status", response_model=ProjectOut)
def update_status(project_id: str, status: str = Query(...),
                  user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    p.status = status
    p.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(p)
    return _project_out(p)


# ---------- 文件上传 ----------
@router.post("/{project_id}/files", response_model=FileOut, status_code=201)
def upload_file(project_id: str,
                file_type: str = Form(...),
                sheet_name: str = Form(None),
                password: str | None = Form(default=None, max_length=256),
                file: FastUploadFile = File(...),
                user: User = Depends(get_current_user),
                db: Session = Depends(get_db)):
    if file_type not in _ALLOWED_TYPES:
        raise HTTPException(status_code=400, detail=f"file_type 必须为 {list(_ALLOWED_TYPES)}")

    started_at = time.perf_counter()
    p = _load_project_or_404(project_id, user, db)

    if file_type in _SINGLE_MASTER_FILE_TYPES:
        existing_master = db.query(UploadFile).filter(
            UploadFile.project_id == p.id,
            UploadFile.file_type == file_type,
        ).first()
        if existing_master:
            master_label = "通用财务总表" if file_type == "financial_master" else "总表"
            raise HTTPException(
                status_code=409,
                detail=f"当前项目已有{master_label}，请先移除原文件后再上传新文件",
            )

    original_content = _read_upload_content(file)
    original_kind = _workbook_binary_kind(original_content)
    content, storage_extension = _normalize_excel_content(original_content, password)
    workbook_info = _inspect_ooxml_workbook(content)
    sheet_names = [str(name) for name in workbook_info["sheet_names"]]
    _validate_declared_workbook_role(file_type, sheet_names)
    max_rows = int(workbook_info["row_count"])
    max_cols = int(workbook_info["col_count"])

    # 存物理文件
    stored_name = f"{uuid.uuid4().hex}{storage_extension}"
    stored_path = os.path.join(p.id, stored_name)
    abs_path = os.path.join(UPLOAD_DIR, stored_path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    with open(abs_path, "wb") as f:
        f.write(content)

    rec = UploadFile(
        project_id=p.id,
        file_type=file_type,
        original_name=file.filename or stored_name,
        stored_path=stored_path,
        sheet_name=sheet_name,
        row_count=max_rows,
        col_count=max_cols,
    )
    try:
        db.add(rec)
        db.commit()
    except Exception:
        db.rollback()
        if os.path.exists(abs_path):
            os.remove(abs_path)
        raise
    db.refresh(rec)
    result = FileOut.from_orm(rec)
    result.processing_seconds = round(time.perf_counter() - started_at, 2)
    if original_kind == "xlsx-encrypted":
        result.normalization_note = "已使用本次输入的打开密码解密后导入"
    elif original_kind == "xls":
        result.normalization_note = "已通过本机 Excel/WPS 自动转换为可处理的 .xlsx"
    return result


@router.post("/{project_id}/files/batch-auto", response_model=list[FileOut], status_code=201)
def upload_files_batch_auto(
    project_id: str,
    files: list[FastUploadFile] = File(...),
    password: str | None = Form(default=None, max_length=256),
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """一次接收全部 Excel，并按工作簿结构自动识别主模板和来源表。"""
    p = _load_project_or_404(project_id, user, db)
    if not files:
        raise HTTPException(status_code=400, detail="请至少选择一个 Excel 文件")

    prepared: list[tuple[str, bytes, str, int, int, str]] = []
    for upload in files:
        filename = upload.filename or "未命名文件.xlsx"
        content, storage_extension = _normalize_excel_content(_read_upload_content(upload), password)
        workbook_info = _inspect_ooxml_workbook(content)
        row_count = int(workbook_info["row_count"])
        col_count = int(workbook_info["col_count"])
        empty_sheets = {
            str(name): pd.DataFrame()
            for name in workbook_info["sheet_names"]
        }
        prepared.append((
            filename,
            content,
            _classify_uploaded_workbook(filename, empty_sheets),
            row_count,
            col_count,
            storage_extension,
        ))

    stored_files: list[str] = []
    records: list[UploadFile] = []
    try:
        for filename, content, file_type, row_count, col_count, storage_extension in prepared:
            stored_name = f"{uuid.uuid4().hex}{storage_extension}"
            stored_path = os.path.join(p.id, stored_name)
            abs_path = os.path.join(UPLOAD_DIR, stored_path)
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            with open(abs_path, "wb") as target:
                target.write(content)
            stored_files.append(abs_path)

            record = UploadFile(
                project_id=p.id,
                file_type=file_type,
                original_name=filename,
                stored_path=stored_path,
                row_count=row_count,
                col_count=col_count,
            )
            db.add(record)
            records.append(record)
        db.commit()
    except Exception:
        db.rollback()
        for stored_file in stored_files:
            if os.path.exists(stored_file):
                os.remove(stored_file)
        raise

    for record in records:
        db.refresh(record)
    return [FileOut.from_orm(record) for record in records]


@router.get("/{project_id}/files", response_model=list[FileOut])
def list_files(project_id: str, user: User = Depends(get_current_user),
               db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    return [FileOut.from_orm(f) for f in p.files]


@router.get("/{project_id}/files/{file_id}/analysis")
def analyze_file(project_id: str, file_id: str,
                 user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    """Return sheet-level classification for an uploaded workbook."""
    p = _load_project_or_404(project_id, user, db)
    file = db.query(UploadFile).filter(
        UploadFile.id == file_id,
        UploadFile.project_id == p.id,
    ).first()
    if not file:
        raise HTTPException(status_code=404, detail="文件不存在")
    file_path = os.path.join(UPLOAD_DIR, file.stored_path)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="物理文件已丢失")
    try:
        analysis = analyze_workbook_structure(file_path)
    except Exception as exc:
        raise HTTPException(status_code=400, detail="工作簿分析失败，请确认文件未损坏") from exc
    return {"file_id": file.id, "filename": file.original_name, **analysis}


@router.delete("/{project_id}/files/{file_id}", status_code=204)
def delete_file(project_id: str, file_id: str, user: User = Depends(get_current_user),
                db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    f = db.query(UploadFile).filter(UploadFile.id == file_id,
                                    UploadFile.project_id == p.id).first()
    if not f:
        raise HTTPException(status_code=404, detail="文件不存在")
    fp = os.path.join(UPLOAD_DIR, f.stored_path)
    if os.path.exists(fp):
        try:
            os.remove(fp)
        except OSError:
            pass
    db.delete(f)
    db.commit()


@router.patch("/{project_id}/files/{file_id}", response_model=FileOut)
def update_file_type(project_id: str, file_id: str, payload: FileTypeUpdateIn,
                     user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    """Reclassify an uploaded file without requiring the user to upload it again."""
    if payload.file_type not in _ALLOWED_TYPES:
        raise HTTPException(status_code=400, detail=f"file_type 必须为 {list(_ALLOWED_TYPES)}")

    p = _load_project_or_404(project_id, user, db)
    f = db.query(UploadFile).filter(UploadFile.id == file_id,
                                    UploadFile.project_id == p.id).first()
    if not f:
        raise HTTPException(status_code=404, detail="文件不存在")

    file_path = os.path.join(UPLOAD_DIR, f.stored_path)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="物理文件已丢失")
    workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
    try:
        _validate_declared_workbook_role(payload.file_type, list(workbook.sheetnames))
    finally:
        workbook.close()
    if payload.file_type in _SINGLE_MASTER_FILE_TYPES:
        existing_master = db.query(UploadFile).filter(
            UploadFile.project_id == p.id,
            UploadFile.file_type == payload.file_type,
            UploadFile.id != f.id,
        ).first()
        if existing_master:
            master_label = "通用财务总表" if payload.file_type == "financial_master" else "总表"
            raise HTTPException(
                status_code=409,
                detail=f"当前项目已有{master_label}，请先移除原文件后再变更",
            )

    f.file_type = payload.file_type
    db.commit()
    db.refresh(f)
    return FileOut.from_orm(f)


# ---------- 文件预览 ----------
@router.get("/{project_id}/files/{file_id}/preview", response_model=FilePreview)
def preview_file(project_id: str, file_id: str,
                 user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    f = db.query(UploadFile).filter(UploadFile.id == file_id,
                                    UploadFile.project_id == p.id).first()
    if not f:
        raise HTTPException(status_code=404, detail="文件不存在")
    fp = os.path.join(UPLOAD_DIR, f.stored_path)
    if not os.path.exists(fp):
        raise HTTPException(status_code=404, detail="物理文件已丢失")

    df = pd.read_excel(fp)
    cols = [str(c) for c in df.columns]
    # 前 20 行，转 dict（NaN -> None）
    head = df.head(20).where(df.notna(), None)
    rows = head.to_dict(orient="records")
    # dict 的 key 转字符串
    rows = [{str(k): v for k, v in r.items()} for r in rows]
    return FilePreview(
        file_id=f.id, columns=cols, rows=rows,
        row_count=len(df), col_count=len(df.columns),
    )


# ---------- 文件数据画像（列质量诊断 + 分布统计） ----------
@router.get("/{project_id}/files/{file_id}/profile")
def profile_file(project_id: str, file_id: str,
                 user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    """返回文件完整数据画像：总览 + 每列质量 + 关键分布。"""
    p = _load_project_or_404(project_id, user, db)
    f = db.query(UploadFile).filter(UploadFile.id == file_id,
                                    UploadFile.project_id == p.id).first()
    if not f:
        raise HTTPException(status_code=404, detail="文件不存在")
    fp = os.path.join(UPLOAD_DIR, f.stored_path)
    if not os.path.exists(fp):
        raise HTTPException(status_code=404, detail="物理文件已丢失")

    df = pd.read_excel(fp)
    total_rows = len(df)
    total_cols = len(df.columns)

    # 总览
    total_cells = total_rows * total_cols
    missing_cells = int(df.isna().sum().sum())
    completeness = round((1 - missing_cells / total_cells) * 100, 1) if total_cells else 100.0

    # 每列质量诊断
    columns_profile = []
    for col in df.columns:
        s = df[col]
        col_str = str(col)
        missing = int(s.isna().sum())
        unique = int(s.nunique(dropna=True))
        # 推断数据类型
        non_null = s.dropna()
        if len(non_null) == 0:
            dtype = "empty"
        elif pd.api.types.is_numeric_dtype(non_null):
            dtype = "number"
        elif pd.api.types.is_datetime64_any_dtype(non_null):
            dtype = "datetime"
        else:
            # 尝试判断是否为日期字符串
            dtype = "text"

        col_info = {
            "name": col_str,
            "dtype": dtype,
            "missing_count": missing,
            "missing_rate": round(missing / total_rows * 100, 1) if total_rows else 0.0,
            "unique_count": unique,
            "sample_values": [str(v) for v in non_null.head(3).tolist()],
        }
        # 数值列附加统计
        if dtype == "number":
            col_info["min"] = float(non_null.min()) if len(non_null) else None
            col_info["max"] = float(non_null.max()) if len(non_null) else None
            col_info["mean"] = round(float(non_null.mean()), 2) if len(non_null) else None
        columns_profile.append(col_info)

    # 关键分布：分类列（唯一值 ≤ 15）的值分布
    distributions = []
    for col in df.columns:
        s = df[col].dropna()
        if len(s) == 0:
            continue
        unique_n = s.nunique()
        if 1 < unique_n <= 15:
            vc = s.astype(str).value_counts().head(15)
            distributions.append({
                "column": str(col),
                "type": "category",
                "items": [{"label": str(k), "count": int(v)} for k, v in vc.items()],
            })
        elif pd.api.types.is_numeric_dtype(s) and unique_n > 5:
            # 数值列做分箱直方图（6 个箱）
            try:
                bins = pd.cut(s, bins=6, include_lowest=True)
                vc = bins.value_counts().sort_index()
                distributions.append({
                    "column": str(col),
                    "type": "histogram",
                    "items": [
                        {"label": str(interval).replace("(", "").replace("]", "")
                                   .replace(", ", "~"),
                         "count": int(v)}
                        for interval, v in vc.items()
                    ],
                })
            except Exception:
                pass

    return {
        "file_id": f.id,
        "filename": f.original_name,
        "file_type": f.file_type,
        "overview": {
            "total_rows": total_rows,
            "total_cols": total_cols,
            "missing_cells": missing_cells,
            "completeness_pct": completeness,
        },
        "columns_profile": columns_profile,
        "distributions": distributions,
    }
