"""薪资业务路由：异动计算 / 诊断 / 在线补齐 / 导出。

把 core/ 模块包装成 REST API。中间结果存 session 级缓存（backend/data/sessions/），
便于多步骤流转，不污染数据库。
"""
import io
import json
import os
import uuid
from datetime import datetime

import openpyxl
import pandas as pd
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from backend.auth import get_current_user
from backend.database import DATA_DIR, EXPORT_DIR, UPLOAD_DIR, get_db
from backend.models import User, Project, UploadFile
from backend.schemas import (
    DiffConfigIn, DiffResult, AuditResult, AuditSaveIn, ExportResult,
)

# 把项目根加入 sys.path 以便 import core
import sys
_proj_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if _proj_root not in sys.path:
    sys.path.insert(0, _proj_root)

from core.db_engine import PayrollDBEngine
from core.excel_engine import ExcelFormulaEngine
from core.normalizer import DataNormalizer
from core.validator import DataValidator

router = APIRouter(prefix="/api/payroll", tags=["payroll"])

# 会话缓存目录：存放异动/诊断中间结果 JSON
SESSION_DIR = os.path.join(DATA_DIR, "sessions")
os.makedirs(SESSION_DIR, exist_ok=True)

# 主键列别名（与前端一致）
_ID_ALIASES = [
    "工号", "员工工号", "员工编号", "工 号", "編號", "员工ID", "ID",
    "Employee ID", "Emp ID", "工牌号",
]


def _detect_id_column(df: pd.DataFrame):
    if df is None or df.empty:
        return None
    cols = [str(c).strip() for c in df.columns]
    for alias in _ID_ALIASES:
        for c in cols:
            if c == alias:
                return c
    for c in cols:
        if any(k in c for k in ("工号", "编号", "ID", "id")):
            return c
    return cols[0] if cols else None


def _load_project_or_404(project_id: str, user: User, db: Session) -> Project:
    p = db.query(Project).filter(Project.id == project_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="项目不存在")
    if p.owner.tenant_id != user.tenant_id:
        raise HTTPException(status_code=403, detail="无权访问")
    return p


def _read_file_df(db: Session, project_id: str, file_id: str) -> pd.DataFrame:
    f = db.query(UploadFile).filter(UploadFile.id == file_id,
                                    UploadFile.project_id == project_id).first()
    if not f:
        raise HTTPException(status_code=404, detail=f"文件不存在: {file_id}")
    fp = os.path.join(UPLOAD_DIR, f.stored_path)
    if not os.path.exists(fp):
        raise HTTPException(status_code=404, detail="物理文件已丢失")
    return pd.read_excel(fp)


def _session_path(project_id: str, key: str) -> str:
    return os.path.join(SESSION_DIR, f"{project_id}_{key}.json")


def _save_session(project_id: str, key: str, data: dict):
    with open(_session_path(project_id, key), "w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False, default=str)


def _load_session(project_id: str, key: str) -> dict | None:
    p = _session_path(project_id, key)
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8") as fp:
        return json.load(fp)


def _df_to_records(df: pd.DataFrame) -> list[dict]:
    """DataFrame -> list[dict]，NaN -> None。"""
    out = df.where(df.notna(), None)
    return [{str(k): v for k, v in r.items()} for r in out.to_dict(orient="records")]


# ============================================================
# 异动计算
# ============================================================
@router.post("/{project_id}/diff", response_model=DiffResult)
def compute_diff(project_id: str, payload: DiffConfigIn,
                 user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)

    lm_df = _read_file_df(db, p.id, payload.last_month_file_id)
    cs_df = _read_file_df(db, p.id, payload.current_file_id)

    engine = PayrollDBEngine()
    try:
        diff = engine.get_personnel_changes(
            lm_df, cs_df,
            id_col=payload.id_col_lm,
            compare_cols=payload.compare_cols or None,
        )
    finally:
        engine.close()

    # 拼接主表（上月 + 新增入职）用于后续诊断
    master = lm_df.copy()
    hires = diff["NEW_HIRES"].copy()
    for c in master.columns:
        if c not in hires.columns:
            hires[c] = None
    hires = hires[master.columns]
    master = pd.concat([master, hires], ignore_index=True)

    # 跨表关联补身份证号
    master_key = payload.id_col_lm or _detect_id_column(master)
    if master_key and "身份证号" not in master.columns:
        master["身份证号"] = ""
    for sfid in payload.social_file_ids:
        try:
            sdf = _read_file_df(db, p.id, sfid)
        except HTTPException:
            continue
        s_key = _detect_id_column(sdf)
        if s_key and "身份证号" in sdf.columns:
            id_map = dict(zip(
                sdf[s_key].apply(DataNormalizer.clean_text),
                sdf["身份证号"].apply(DataNormalizer.clean_text),
            ))
            master["身份证号"] = master.apply(
                lambda r: id_map.get(
                    DataNormalizer.clean_text(r.get(master_key)),
                    r.get("身份证号", "")), axis=1)

    # 存会话
    _save_session(project_id, "diff", {
        "new_hires": _df_to_records(diff["NEW_HIRES"]),
        "departed": _df_to_records(diff["DEPARTED"]),
        "transfers": _df_to_records(diff["TRANSFERS"]),
        "master": _df_to_records(master),
        "master_columns": list(master.columns),
        "id_col_lm": payload.id_col_lm,
        "id_col_cs": payload.id_col_cs,
    })
    return DiffResult(
        new_hires=_df_to_records(diff["NEW_HIRES"]),
        departed=_df_to_records(diff["DEPARTED"]),
        transfers=_df_to_records(diff["TRANSFERS"]),
    )


# ============================================================
# 诊断
# ============================================================
@router.get("/{project_id}/audit", response_model=AuditResult)
def audit(project_id: str, user: User = Depends(get_current_user),
          db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    sess = _load_session(project_id, "diff")
    if not sess:
        # 单表诊断模式：尝试用上月文件直接诊断
        raise HTTPException(status_code=400, detail="请先执行异动计算或上传文件")

    master = pd.DataFrame(sess["master"], columns=sess["master_columns"])
    validator = DataValidator()
    audited = validator.audit_employee_profiles(master)
    metrics = validator.get_completeness_metrics(audited)

    # 存诊断结果供补齐回传比对
    _save_session(project_id, "audit", {
        "rows": _df_to_records(audited),
        "columns": list(audited.columns),
    })
    return AuditResult(
        columns=list(audited.columns),
        rows=_df_to_records(audited),
        status_col=audited["诊断状态"].astype(str).tolist(),
        missing_col=audited["缺失字段清单"].tolist(),
        metrics=metrics,
    )


@router.post("/{project_id}/audit/save", response_model=AuditResult)
def save_audit(project_id: str, payload: AuditSaveIn,
               user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    # 用户回传的整表
    edited = pd.DataFrame(payload.rows)
    if edited.empty:
        raise HTTPException(status_code=400, detail="数据为空")
    # 去掉旧的两列再重新诊断
    biz_cols = [c for c in edited.columns if c not in ("诊断状态", "缺失字段清单")]
    validator = DataValidator()
    re_audited = validator.audit_employee_profiles(edited[biz_cols])
    metrics = validator.get_completeness_metrics(re_audited)
    _save_session(project_id, "audit", {
        "rows": _df_to_records(re_audited),
        "columns": list(re_audited.columns),
    })
    return AuditResult(
        columns=list(re_audited.columns),
        rows=_df_to_records(re_audited),
        status_col=re_audited["诊断状态"].astype(str).tolist(),
        missing_col=re_audited["缺失字段清单"].tolist(),
        metrics=metrics,
    )


# ============================================================
# 导出
# ============================================================
@router.post("/{project_id}/export", response_model=ExportResult)
def export(project_id: str, user: User = Depends(get_current_user),
           db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)

    audit_sess = _load_session(project_id, "audit")
    if not audit_sess:
        raise HTTPException(status_code=400, detail="请先完成诊断")
    audited_df = pd.DataFrame(audit_sess["rows"], columns=audit_sess["columns"])

    validator = DataValidator()
    metrics = validator.get_completeness_metrics(audited_df)
    if metrics["high_risk_count"] > 0:
        raise HTTPException(status_code=400,
                            detail=f"仍有 {metrics['high_risk_count']} 人核心项缺失，无法导出")

    diff_sess = _load_session(project_id, "diff")
    log: list[str] = []

    # 加载上月主模板（保留原始公式样式）
    last_month_file = db.query(UploadFile).filter(
        UploadFile.project_id == p.id, UploadFile.file_type == "last_month"
    ).first()
    if last_month_file:
        wb = openpyxl.load_workbook(os.path.join(UPLOAD_DIR, last_month_file.stored_path))
        log.append(f"加载主模板：{last_month_file.original_name}")
    else:
        # 无模板则用主表生成新工作簿
        wb = openpyxl.Workbook()
        ws_new = wb.active
        master_sess = diff_sess or {}
        master_df = pd.DataFrame(master_sess.get("master", []),
                                 columns=master_sess.get("master_columns", []))
        ws_new.append(list(master_df.columns))
        for _, r in master_df.iterrows():
            ws_new.append([r[c] if pd.notna(r[c]) else None for c in master_df.columns])
        log.append("无原始模板，按主表生成新工作簿")

    ws = wb.active
    log.append(f"工作表：{ws.title}，{ws.max_row} 行 × {ws.max_column} 列")

    # 定位工号列
    id_col_idx = 1
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h and str(h).strip() in _ID_ALIASES:
            id_col_idx = c
            break
    log.append(f"工号列定位：第 {id_col_idx} 列")

    header_map = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h:
            header_map[str(h).strip()] = c

    # 删除离职行
    if diff_sess and diff_sess.get("departed"):
        departed_df = pd.DataFrame(diff_sess["departed"])
        key_col_name = None
        for alias in _ID_ALIASES:
            if alias in departed_df.columns:
                key_col_name = alias
                break
        if key_col_name:
            departed_keys = departed_df[key_col_name].apply(DataNormalizer.clean_text).tolist()
            deleted = ExcelFormulaEngine.delete_departed_rows(
                ws, key_col_idx=id_col_idx, departed_keys=departed_keys)
            log.append(f"删除离职行：{deleted} 行")
        else:
            log.append("离职名单无工号列，跳过删行")

    # 插入新增入职行
    if diff_sess and diff_sess.get("new_hires"):
        hires_df = pd.DataFrame(diff_sess["new_hires"])
        template_row = ws.max_row
        inserted = 0
        for _, row in hires_df.iterrows():
            emp = {}
            for col_name, col_idx in header_map.items():
                if col_name in row.index and pd.notna(row[col_name]):
                    emp[col_idx] = row[col_name]
            ExcelFormulaEngine.insert_employee_row(
                ws, template_row_idx=template_row, employee_data=emp)
            template_row += 1
            inserted += 1
        log.append(f"插入新增入职行：{inserted} 行")

    # 写回补齐的身份证号
    if "身份证号" in header_map and "身份证号" in audited_df.columns:
        idc = header_map["身份证号"]
        key_col_name = None
        for alias in _ID_ALIASES:
            if alias in audited_df.columns:
                key_col_name = alias
                break
        if key_col_name:
            id_map = dict(zip(
                audited_df[key_col_name].apply(DataNormalizer.clean_text),
                audited_df["身份证号"].apply(DataNormalizer.clean_text)))
            written = 0
            for r in range(2, ws.max_row + 1):
                eid = DataNormalizer.clean_text(ws.cell(row=r, column=id_col_idx).value)
                if eid in id_map and id_map[eid]:
                    ws.cell(row=r, column=idc).value = id_map[eid]
                    written += 1
            log.append(f"写回身份证号：{written} 条")

    # 保存导出文件
    os.makedirs(os.path.join(EXPORT_DIR, p.id), exist_ok=True)
    out_name = f"薪资汇总表_{p.salary_month}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    out_path = os.path.join(EXPORT_DIR, p.id, out_name)
    wb.save(out_path)
    log.append(f"导出完成：{out_name}")

    # 记录到 session 供下载
    _save_session(project_id, "export", {
        "path": os.path.join(p.id, out_name),
        "filename": out_name,
    })

    return ExportResult(file_id=uuid.uuid4().hex, filename=out_name, log=log)


@router.get("/{project_id}/export/download")
def download_export(project_id: str, user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    sess = _load_session(project_id, "export")
    if not sess:
        raise HTTPException(status_code=400, detail="尚未生成导出文件")
    fp = os.path.join(EXPORT_DIR, sess["path"])
    if not os.path.exists(fp):
        raise HTTPException(status_code=404, detail="导出文件已丢失")
    with open(fp, "rb") as f:
        data = f.read()
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{sess["filename"]}"'},
    )
