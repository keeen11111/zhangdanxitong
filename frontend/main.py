from __future__ import annotations

import re
import zipfile
from decimal import Decimal, InvalidOperation
from io import BytesIO
from typing import Any, Iterable
from urllib.parse import quote

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from starlette.requests import Request


app = FastAPI(title="PayrollForge Python API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://127.0.0.1:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/pyapi/health")
async def health_api():
    return {"ok": True}


@app.exception_handler(Exception)
async def unhandled_exception_api(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"detail": f"Internal Server Error: {type(exc).__name__}: {str(exc)}"},
    )


def _read_excel(data: bytes, sheet_name: str | None) -> pd.DataFrame:
    try:
        df = pd.read_excel(
            BytesIO(data),
            sheet_name=(sheet_name if sheet_name else 0),
            dtype=str,
            na_filter=False,
            engine="openpyxl",
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"无法读取 Excel：{e}")
    if isinstance(df, dict):
        df = next(iter(df.values()))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _list_sheet_names(data: bytes) -> list[str]:
    try:
        xls = pd.ExcelFile(BytesIO(data), engine="openpyxl")
        return [str(s) for s in xls.sheet_names]
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"无法读取工作簿：{e}")


def _clean_header_cell(v: Any) -> str:
    s = "" if v is None else str(v)
    s = s.replace("\u00A0", " ").strip()
    if not s:
        return ""
    if s.startswith("Unnamed:"):
        return ""
    return s


def _read_header_cells(data: bytes, sheet_name: str, header_row_index: int) -> list[str]:
    try:
        row_df = pd.read_excel(
            BytesIO(data),
            sheet_name=sheet_name,
            header=None,
            skiprows=header_row_index,
            nrows=1,
            dtype=str,
            na_filter=False,
            engine="openpyxl",
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"无法读取表头：{e}")
    if isinstance(row_df, dict):
        row_df = next(iter(row_df.values()))
    values = row_df.iloc[0].tolist() if not row_df.empty else []
    headers = [_clean_header_cell(v) for v in values]
    while headers and headers[-1] == "":
        headers.pop()
    return [h for h in headers if h != ""]


def _detect_header_row(data: bytes, sheet_name: str | None, max_rows: int = 30) -> int:
    try:
        tmp = pd.read_excel(
            BytesIO(data),
            sheet_name=(sheet_name if sheet_name else 0),
            header=None,
            nrows=max_rows,
            dtype=str,
            na_filter=False,
            engine="openpyxl",
        )
    except Exception:
        return 0
    if isinstance(tmp, dict):
        tmp = next(iter(tmp.values()))
    best_row = 0
    best_score = -1
    for i in range(min(max_rows, len(tmp))):
        row = tmp.iloc[i].astype(str).map(lambda x: str(x).strip())
        score = int(row.ne("").sum())
        if score > best_score:
            best_score = score
            best_row = i
    return best_row


def _read_excel_auto_header(data: bytes, sheet_name: str | None) -> pd.DataFrame:
    header_row = _detect_header_row(data, sheet_name)
    try:
        df = pd.read_excel(
            BytesIO(data),
            sheet_name=(sheet_name if sheet_name else 0),
            header=header_row,
            dtype=str,
            na_filter=False,
            engine="openpyxl",
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"无法读取 Excel：{e}")
    if isinstance(df, dict):
        df = next(iter(df.values()))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _require_columns(df: pd.DataFrame, cols: list[str], label: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise HTTPException(status_code=400, detail=f"{label}缺少列：{', '.join(missing)}")


def _idcard_checksum_ok(s: str) -> bool:
    s = s.strip()
    if not re.fullmatch(r"\d{17}[\dXx]", s):
        return False
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    mapping = ["1", "0", "X", "9", "8", "7", "6", "5", "4", "3", "2"]
    total = sum(int(s[i]) * weights[i] for i in range(17))
    expected = mapping[total % 11]
    return s[-1].upper() == expected


def _id15_to_18(s: str) -> str:
    s = s.strip()
    if not re.fullmatch(r"\d{15}", s):
        return s
    base = s[:6] + "19" + s[6:]
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    mapping = ["1", "0", "X", "9", "8", "7", "6", "5", "4", "3", "2"]
    total = sum(int(base[i]) * weights[i] for i in range(17))
    return base + mapping[total % 11]


def _normalize_numeric_text(s: str) -> str:
    s = str(s).strip()
    if not s:
        return ""
    s = s.replace("\u00A0", " ").strip()
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".", 1)[0]
    if re.fullmatch(r"\d+", s):
        return s
    if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
        try:
            d = Decimal(s)
        except InvalidOperation:
            return s
        if d == d.to_integral_value():
            return str(int(d))
    return s


def _normalize_key_cell(v: Any) -> str:
    s = "" if v is None else str(v)
    s = s.replace("\u00A0", " ").strip()
    if not s:
        return ""
    if s in {"合计", "总计", "小计", "合计:", "合计：", "总计:", "总计："}:
        return ""
    return _normalize_numeric_text(s)


def _normalize_idcard_value(v: Any) -> str:
    s = _normalize_key_cell(v)
    if not s:
        return ""
    s = s.replace(" ", "").upper()
    if re.fullmatch(r"\d{15}", s):
        return _id15_to_18(s)
    if re.fullmatch(r"\d{17}[\dX]", s):
        return s
    return ""


def _normalize_value_for_compare(v: Any) -> str:
    s = "" if v is None else str(v)
    s = s.replace("\u00A0", " ").strip()
    if not s:
        return ""
    return _normalize_numeric_text(s)


def _is_ignored_compare_column(col: str) -> bool:
    return str(col).strip() == "序号"


def _build_key_parts(df: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame({c: df[c].map(_normalize_key_cell) for c in key_columns}, index=df.index)


def _join_key(parts: pd.DataFrame, key_columns: list[str]) -> pd.Series:
    if len(key_columns) == 1:
        return parts[key_columns[0]].astype(str)
    return parts[key_columns].astype(str).agg("||".join, axis=1)


def _row_to_dict(row: pd.Series, cols: list[str]) -> dict[str, Any]:
    return {c: (None if pd.isna(row.get(c)) else row.get(c)) for c in cols}


def _compare(
    old_df: pd.DataFrame,
    new_df: pd.DataFrame,
    key_columns: list[str],
    validate_idcard: bool,
    include_rows: bool = False,
) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []
    if key_columns:
        _require_columns(old_df, key_columns, "旧表")
        _require_columns(new_df, key_columns, "新表")

    if key_columns:
        old_parts = _build_key_parts(old_df, key_columns)
        new_parts = _build_key_parts(new_df, key_columns)

        if validate_idcard:
            id_col = key_columns[0]
            old_norm = old_df[id_col].map(_normalize_idcard_value)
            new_norm = new_df[id_col].map(_normalize_idcard_value)
            non_id_old = old_df.loc[(old_norm == "") & old_df[id_col].astype(str).str.strip().ne(""), id_col].astype(str)
            non_id_new = new_df.loc[(new_norm == "") & new_df[id_col].astype(str).str.strip().ne(""), id_col].astype(str)
            if not non_id_old.empty:
                issues.append(
                    {
                        "severity": "warning",
                        "scope": "old",
                        "code": "IDCARD_NON_ID",
                        "message": f"旧表存在非身份证值（将忽略这些行的匹配键）：{id_col}",
                        "count": int(non_id_old.shape[0]),
                        "samples": non_id_old.head(5).tolist(),
                    }
                )
            if not non_id_new.empty:
                issues.append(
                    {
                        "severity": "warning",
                        "scope": "new",
                        "code": "IDCARD_NON_ID",
                        "message": f"新表存在非身份证值（将忽略这些行的匹配键）：{id_col}",
                        "count": int(non_id_new.shape[0]),
                        "samples": non_id_new.head(5).tolist(),
                    }
                )

            checksum_fail_old = old_norm[(old_norm != "") & (~old_norm.map(_idcard_checksum_ok))]
            checksum_fail_new = new_norm[(new_norm != "") & (~new_norm.map(_idcard_checksum_ok))]
            if not checksum_fail_old.empty:
                issues.append(
                    {
                        "severity": "warning",
                        "scope": "old",
                        "code": "IDCARD_CHECKSUM_FAIL",
                        "message": f"旧表身份证校验不通过（仍将参与对比）：{id_col}",
                        "count": int(checksum_fail_old.shape[0]),
                        "samples": checksum_fail_old.head(5).tolist(),
                    }
                )
            if not checksum_fail_new.empty:
                issues.append(
                    {
                        "severity": "warning",
                        "scope": "new",
                        "code": "IDCARD_CHECKSUM_FAIL",
                        "message": f"新表身份证校验不通过（仍将参与对比）：{id_col}",
                        "count": int(checksum_fail_new.shape[0]),
                        "samples": checksum_fail_new.head(5).tolist(),
                    }
                )

            old_parts[id_col] = old_norm
            new_parts[id_col] = new_norm

        old_keys = _join_key(old_parts, key_columns)
        new_keys = _join_key(new_parts, key_columns)

        old_invalid = old_parts.eq("").any(axis=1)
        new_invalid = new_parts.eq("").any(axis=1)
        if bool(old_invalid.any()):
            issues.append(
                {
                    "severity": "warning",
                    "scope": "old",
                    "code": "KEY_EMPTY",
                    "message": "旧表匹配键存在空值/总计行，将忽略这些行",
                    "count": int(old_invalid.sum()),
                    "samples": old_keys[old_invalid].head(5).tolist(),
                }
            )
        if bool(new_invalid.any()):
            issues.append(
                {
                    "severity": "warning",
                    "scope": "new",
                    "code": "KEY_EMPTY",
                    "message": "新表匹配键存在空值/总计行，将忽略这些行",
                    "count": int(new_invalid.sum()),
                    "samples": new_keys[new_invalid].head(5).tolist(),
                }
            )

        old_df = old_df.loc[~old_invalid].copy()
        new_df = new_df.loc[~new_invalid].copy()
        old_keys = old_keys.loc[~old_invalid]
        new_keys = new_keys.loc[~new_invalid]

        old_dups = old_keys[old_keys.duplicated(keep=False)]
        new_dups = new_keys[new_keys.duplicated(keep=False)]
        if not old_dups.empty:
            issues.append(
                {
                    "severity": "error",
                    "scope": "old",
                    "code": "DUPLICATE_KEY",
                    "message": "旧表匹配键不唯一，已忽略重复键对应的所有行",
                    "count": int(old_dups.shape[0]),
                    "samples": old_dups.head(5).tolist(),
                }
            )
            keep = ~old_keys.duplicated(keep=False)
            old_df = old_df.loc[keep].copy()
            old_keys = old_keys.loc[keep]

        if not new_dups.empty:
            issues.append(
                {
                    "severity": "error",
                    "scope": "new",
                    "code": "DUPLICATE_KEY",
                    "message": "新表匹配键不唯一，已忽略重复键对应的所有行",
                    "count": int(new_dups.shape[0]),
                    "samples": new_dups.head(5).tolist(),
                }
            )
            keep = ~new_keys.duplicated(keep=False)
            new_df = new_df.loc[keep].copy()
            new_keys = new_keys.loc[keep]
    else:
        old_keys = pd.Series((i + 1 for i in range(len(old_df))), index=old_df.index, dtype=str)
        new_keys = pd.Series((i + 1 for i in range(len(new_df))), index=new_df.index, dtype=str)

    old_df = old_df.copy()
    new_df = new_df.copy()
    old_df["__key"] = old_keys.astype(str)
    new_df["__key"] = new_keys.astype(str)

    old_cols = [c for c in old_df.columns if c != "__key"]
    new_cols = [c for c in new_df.columns if c != "__key"]
    all_cols = [*old_cols, *[c for c in new_cols if c not in old_cols]]
    non_key_cols = [c for c in all_cols if c not in key_columns]
    compare_cols = [c for c in non_key_cols if not _is_ignored_compare_column(c)]

    for c in all_cols:
        if c not in old_df.columns:
            old_df[c] = ""
        if c not in new_df.columns:
            new_df[c] = ""

    old_i = old_df.set_index("__key", drop=True)
    new_i = new_df.set_index("__key", drop=True)

    old_key_set = set(old_i.index.astype(str))
    new_key_set = set(new_i.index.astype(str))
    added_keys = sorted(new_key_set - old_key_set)
    removed_keys = sorted(old_key_set - new_key_set)
    common_keys = sorted(old_key_set & new_key_set)

    diffs: list[dict[str, Any]] = []

    for k in added_keys:
        row = new_i.loc[k]
        diffs.append({"kind": "added", "key": k, "row": _row_to_dict(row, all_cols)})

    for k in removed_keys:
        row = old_i.loc[k]
        diffs.append({"kind": "removed", "key": k, "row": _row_to_dict(row, all_cols)})

    if common_keys:
        old_common = old_i.loc[common_keys, compare_cols]
        new_common = new_i.loc[common_keys, compare_cols]
        old_hash = pd.util.hash_pandas_object(old_common, index=False)
        new_hash = pd.util.hash_pandas_object(new_common, index=False)
        modified_mask = old_hash.ne(new_hash)
        modified_keys = [common_keys[i] for i, m in enumerate(modified_mask.tolist()) if m]

        for k in modified_keys:
            old_row = old_i.loc[k]
            new_row = new_i.loc[k]
            changes: list[dict[str, Any]] = []
            for c in compare_cols:
                ov = old_row.get(c, "")
                nv = new_row.get(c, "")
                if _normalize_value_for_compare(ov) != _normalize_value_for_compare(nv):
                    changes.append(
                        {
                            "column": c,
                            "oldValue": None if ov == "" else str(ov),
                            "newValue": None if nv == "" else str(nv),
                        }
                    )
            if changes:
                item: dict[str, Any] = {"kind": "modified", "key": k, "changes": changes}
                if include_rows:
                    item["oldRow"] = _row_to_dict(old_row, all_cols)
                    item["newRow"] = _row_to_dict(new_row, all_cols)
                diffs.append(item)

    summary = {
        "added": len(added_keys),
        "removed": len(removed_keys),
        "modified": sum(1 for d in diffs if d["kind"] == "modified"),
    }

    return {"summary": summary, "diffs": diffs, "meta": {"columns": all_cols, "issues": issues}}


def _excel_report_bytes(result: dict[str, Any]) -> bytes:
    diffs: list[dict[str, Any]] = result["diffs"]
    cols: list[str] = result["meta"]["columns"]
    issues: list[dict[str, Any]] = result.get("meta", {}).get("issues", [])
    added_rows = [d["row"] for d in diffs if d["kind"] == "added"]
    removed_rows = [d["row"] for d in diffs if d["kind"] == "removed"]

    modified_two_rows: list[dict[str, Any]] = []
    for d in diffs:
        if d["kind"] != "modified":
            continue
        changed_cols = [ch["column"] for ch in d.get("changes", [])]
        modified_two_rows.append(
            {
                "key": d["key"],
                "version": "old",
                "changedColumns": ",".join(changed_cols),
                "__changedCols": changed_cols,
                "__row": d.get("oldRow", {}),
            }
        )
        modified_two_rows.append(
            {
                "key": d["key"],
                "version": "new",
                "changedColumns": "",
                "__changedCols": [],
                "__row": d.get("newRow", {}),
            }
        )
        modified_two_rows.append({"__sep": True})

    summary_df = pd.DataFrame([result["summary"]])
    added_df = pd.DataFrame(added_rows, columns=cols)
    removed_df = pd.DataFrame(removed_rows, columns=cols)
    issues_df = pd.DataFrame(issues, columns=["severity", "scope", "code", "message", "count", "samples"])

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="summary")
        added_df.to_excel(writer, index=False, sheet_name="added")
        removed_df.to_excel(writer, index=False, sheet_name="removed")
        book = writer.book
        ws = book.add_worksheet("modified")
        writer.sheets["modified"] = ws

        header_fmt = book.add_format({"bold": True, "bg_color": "#F1F5F9", "border": 1})
        cell_fmt = book.add_format({"border": 1})
        yellow_fmt = book.add_format({"bg_color": "#FEF3C7", "border": 1})
        key_fmt = book.add_format({"border": 1, "font_color": "#334155"})
        meta_fmt = book.add_format({"border": 1, "font_color": "#64748B"})
        sep_fmt = book.add_format({"bg_color": "#DBEAFE", "border": 1})

        out_cols = ["key", "version", "changedColumns", *cols]
        for j, h in enumerate(out_cols):
            ws.write(0, j, h, header_fmt)

        ws.set_column(0, 0, 26)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 2, 40)
        if cols:
            ws.set_column(3, 3 + len(cols) - 1, 18)

        r = 1
        for item in modified_two_rows:
            if item.get("__sep"):
                ws.set_row(r, 10)
                for j in range(len(out_cols)):
                    ws.write_blank(r, j, None, sep_fmt)
                r += 1
                continue
            key = item["key"]
            ver = item["version"]
            changed = set(item.get("__changedCols", []))
            row = item.get("__row", {})

            ws.write(r, 0, key, key_fmt)
            ws.write(r, 1, ver, meta_fmt)
            ws.write(r, 2, item.get("changedColumns", ""), meta_fmt)

            for i, c in enumerate(cols):
                v = row.get(c, "")
                fmt = yellow_fmt if (ver == "old" and c in changed) else cell_fmt
                ws.write(r, 3 + i, "" if v is None else str(v), fmt)
            r += 1

        if not issues_df.empty:
            issues_df.to_excel(writer, index=False, sheet_name="issues")
    return buf.getvalue()


def _safe_filename(name: str) -> str:
    s = str(name).strip()
    if not s:
        s = "EMPTY"
    s = re.sub(r"[\\/:*?\"<>|]", "_", s)
    s = re.sub(r"\s+", " ", s)
    if len(s) > 80:
        s = s[:80].rstrip()
    return s


def _content_disposition_attachment(filename: str) -> str:
    raw = str(filename).strip() or "output.xlsx"
    ascii_fallback = re.sub(r"[^A-Za-z0-9._ -]", "_", raw)
    ascii_fallback = re.sub(r"\s+", " ", ascii_fallback).strip() or "output.xlsx"
    encoded = quote(raw.encode("utf-8"))
    return f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded}"


def _zip_bytes(files: Iterable[tuple[str, bytes]]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, content in files:
            z.writestr(name, content)
    return buf.getvalue()


def _merge_validate(data: bytes, sheet_names: list[str]) -> dict[str, Any]:
    if not sheet_names:
        raise HTTPException(status_code=400, detail="请至少选择一个 Sheet")
    all_sheets = _list_sheet_names(data)
    for s in sheet_names:
        if s not in all_sheets:
            raise HTTPException(status_code=400, detail=f"工作簿不存在 Sheet：{s}")

    base_sheet = sheet_names[0]
    base_header_row = _detect_header_row(data, base_sheet)
    base_header = _read_header_cells(data, base_sheet, base_header_row)
    if not base_header:
        raise HTTPException(status_code=400, detail=f"无法识别表头：{base_sheet}")

    issues: list[dict[str, Any]] = []
    base_index = {c: i for i, c in enumerate(base_header)}
    base_set = set(base_header)

    for s in sheet_names[1:]:
        header_row = _detect_header_row(data, s)
        hdr = _read_header_cells(data, s, header_row)
        hdr_set = set(hdr)

        for c in sorted(base_set - hdr_set):
            issues.append({"sheetName": s, "type": "missing_column", "column": c})
        for c in sorted(hdr_set - base_set):
            issues.append({"sheetName": s, "type": "extra_column", "column": c})

        common = [c for c in base_header if c in hdr_set]
        hdr_index = {c: i for i, c in enumerate(hdr)}
        for c in common:
            bi = base_index.get(c)
            hi = hdr_index.get(c)
            if bi is None or hi is None:
                continue
            if bi != hi:
                issues.append(
                    {
                        "sheetName": s,
                        "type": "order_mismatch",
                        "column": c,
                        "expectedIndex": int(bi),
                        "actualIndex": int(hi),
                    }
                )

    ok = len(issues) == 0
    return {"baseHeader": base_header, "selectedSheets": sheet_names, "ok": ok, "issues": issues}


@app.post("/pyapi/merge/validate")
async def merge_validate_api(
    file: UploadFile = File(...),
    sheetNames: list[str] = Form([]),
):
    data = await file.read()
    return JSONResponse(content=_merge_validate(data, sheetNames))


@app.post("/pyapi/merge/export")
async def merge_export_api(
    file: UploadFile = File(...),
    sheetNames: list[str] = Form([]),
):
    data = await file.read()
    v = _merge_validate(data, sheetNames)
    if not v.get("ok", False):
        raise HTTPException(status_code=400, detail={"message": "所选 Sheet 表头不一致", "issues": v.get("issues", [])})
    base_header = v.get("baseHeader", [])

    dfs: list[pd.DataFrame] = []
    for s in sheetNames:
        header_row = _detect_header_row(data, s)
        df = pd.read_excel(
            BytesIO(data),
            sheet_name=s,
            header=header_row,
            dtype=str,
            na_filter=False,
            engine="openpyxl",
        )
        if isinstance(df, dict):
            df = next(iter(df.values()))
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, [c for c in df.columns if _clean_header_cell(c) != ""]]
        df = df.reindex(columns=base_header)
        dfs.append(df)

    merged = pd.concat(dfs, ignore_index=True)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        merged.to_excel(writer, index=False, sheet_name="merged")
    out = buf.getvalue()
    out_name = f"{_safe_filename(file.filename or 'merged')}_merged.xlsx"
    return StreamingResponse(
        BytesIO(out),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


@app.post("/pyapi/summary/export")
async def summary_export_api(
    file: UploadFile = File(...),
    sheetName: str = Form(...),
    groupByColumns: list[str] = Form([]),
    valueColumns: list[str] = Form([]),
    metrics: list[str] = Form([]),
):
    data = await file.read()
    grouped = _summary_build(data, sheetName, groupByColumns, valueColumns, metrics)
    try:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            grouped.to_excel(writer, index=False, sheet_name="summary")
        out = buf.getvalue()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"写出 Excel 失败：{str(e)}")

    out_name = f"{_safe_filename(file.filename or 'summary')}_summary.xlsx"
    return StreamingResponse(
        BytesIO(out),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _content_disposition_attachment(out_name)},
    )


@app.post("/pyapi/summary/preview")
async def summary_preview_api(
    file: UploadFile = File(...),
    sheetName: str = Form(...),
    groupByColumns: list[str] = Form([]),
    valueColumns: list[str] = Form([]),
    metrics: list[str] = Form([]),
    limit: int = Form(50),
):
    data = await file.read()
    grouped = _summary_build(data, sheetName, groupByColumns, valueColumns, metrics)
    limit = max(1, min(int(limit or 50), 200))
    cols = [str(c) for c in grouped.columns]
    rows = grouped.head(limit).to_dict(orient="records")
    return {"columns": cols, "rows": rows, "totalGroups": int(len(grouped))}


def _summary_build(
    data: bytes,
    sheet_name: str,
    group_by_columns: list[str],
    value_columns: list[str],
    metrics: list[str],
) -> pd.DataFrame:
    sheet_name = str(sheet_name).strip()
    if not sheet_name:
        raise HTTPException(status_code=400, detail="请指定 Sheet")
    if not group_by_columns:
        raise HTTPException(status_code=400, detail="请至少选择一个分类列")
    metrics = [str(m).strip() for m in metrics if str(m).strip()]
    if not metrics:
        metrics = ["count"]
    if "sum" in metrics and not value_columns:
        raise HTTPException(status_code=400, detail="选择 sum 指标时必须选择至少一个汇总列")

    df = _read_excel(data, sheet_name)
    df.columns = [_clean_header_cell(c) for c in df.columns]
    needed = [*_safe_str_list(group_by_columns), *_safe_str_list(value_columns)]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        df = _read_excel_auto_header(data, sheet_name)
        df.columns = [_clean_header_cell(c) for c in df.columns]
        missing2 = [c for c in needed if c not in df.columns]
        if missing2:
            sample = ", ".join([str(c) for c in df.columns[:20]])
            raise HTTPException(status_code=400, detail=f"指定列不存在：{', '.join(missing2)}（示例列：{sample}）")

    group_by_cols = _safe_str_list(group_by_columns)
    value_cols = [c for c in _safe_str_list(value_columns) if c not in set(group_by_cols)]

    work = df.copy()
    for c in group_by_cols:
        work[c] = work[c].astype(str).map(lambda x: str(x).replace("\u00A0", " ").strip())

    def _to_num(s: pd.Series) -> pd.Series:
        t = s.astype(str).str.replace(",", "", regex=False)
        t = t.str.replace("\u00A0", " ").str.strip()
        return pd.to_numeric(t, errors="coerce").fillna(0)

    try:
        gb = work.groupby(group_by_cols, dropna=False)
        result_df: pd.DataFrame | None = None

        if "count" in metrics:
            result_df = gb.size().reset_index(name="count")

        if "sum" in metrics and value_cols:
            numeric = {c: _to_num(work[c]) for c in value_cols}
            sum_work = pd.concat([work[group_by_cols], pd.DataFrame(numeric)], axis=1)
            sum_df = sum_work.groupby(group_by_cols, dropna=False).sum(numeric_only=True).reset_index()
            sum_df = sum_df.rename(columns={c: f"{c}_sum" for c in value_cols if c in sum_df.columns})
            if result_df is None:
                result_df = sum_df
            else:
                result_df = result_df.merge(sum_df, on=group_by_cols, how="left")

        if result_df is None:
            result_df = gb.size().reset_index(name="count")
        return result_df
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"分类汇总失败：{type(e).__name__}: {str(e)}")


def _safe_str_list(xs: Iterable[Any]) -> list[str]:
    out: list[str] = []
    for x in xs:
        s = str(x).replace("\u00A0", " ").strip()
        if not s:
            continue
        if s in out:
            continue
        out.append(s)
    return out


@app.post("/pyapi/compare")
async def compare_api(
    oldFile: UploadFile = File(...),
    newFile: UploadFile = File(...),
    sheetName: str | None = Form(None),
    oldSheetName: str | None = Form(None),
    newSheetName: str | None = Form(None),
    keyColumns: list[str] = Form([]),
    validateIdCard: bool = Form(False),
):
    old_bytes = await oldFile.read()
    new_bytes = await newFile.read()
    old_df = _read_excel(old_bytes, oldSheetName or sheetName)
    new_df = _read_excel(new_bytes, newSheetName or sheetName)
    result = _compare(old_df, new_df, keyColumns, validateIdCard, include_rows=False)
    return JSONResponse(content=result)


@app.post("/pyapi/compare/report")
async def compare_report_api(
    oldFile: UploadFile = File(...),
    newFile: UploadFile = File(...),
    sheetName: str | None = Form(None),
    oldSheetName: str | None = Form(None),
    newSheetName: str | None = Form(None),
    keyColumns: list[str] = Form([]),
    validateIdCard: bool = Form(False),
):
    old_bytes = await oldFile.read()
    new_bytes = await newFile.read()
    old_df = _read_excel(old_bytes, oldSheetName or sheetName)
    new_df = _read_excel(new_bytes, newSheetName or sheetName)
    result = _compare(old_df, new_df, keyColumns, validateIdCard, include_rows=True)
    report = _excel_report_bytes(result)
    filename = "diff_report.xlsx"
    return StreamingResponse(
        BytesIO(report),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=\"{filename}\""},
    )


@app.post("/pyapi/split")
async def split_api(
    file: UploadFile = File(...),
    sheetName: str | None = Form(None),
    splitColumnName: str = Form(...),
):
    data = await file.read()
    df = _read_excel(data, sheetName)
    splitColumnName = str(splitColumnName).strip()
    if splitColumnName not in df.columns:
        df = _read_excel_auto_header(data, sheetName)
        if splitColumnName not in df.columns:
            sample = ", ".join([str(c) for c in df.columns[:20]])
            raise HTTPException(status_code=400, detail=f"指定的列不存在：{splitColumnName}（示例列：{sample}）")
    col = df[splitColumnName].astype(str).str.strip()
    df = df.copy()
    df[splitColumnName] = col.replace({"": "EMPTY"})

    used: dict[str, int] = {}
    items: list[tuple[str, bytes]] = []
    for value, g in df.groupby(splitColumnName, sort=True):
        base = _safe_filename(value)
        used[base] = used.get(base, 0) + 1
        name = base if used[base] == 1 else f"{base}_{used[base]}"
        out = BytesIO()
        g.to_excel(out, index=False)
        items.append((f"{name}.xlsx", out.getvalue()))

    zip_content = _zip_bytes(items)
    filename = "output.zip"
    return StreamingResponse(
        BytesIO(zip_content),
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=\"{filename}\""},
    )
