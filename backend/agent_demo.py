"""Explicit, project-scoped demo delivery of a user-provided completed file."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
from typing import Any

from fastapi import HTTPException

from backend.agent_planning import file_digest


def demo_directory(root: Path, tenant_id: str, project_id: str) -> Path:
    key = lambda value: hashlib.sha256(str(value).encode("utf-8")).hexdigest()
    return root / key(tenant_id) / key(project_id)


def load_demo(root: Path, tenant_id: str, project_id: str) -> dict[str, Any] | None:
    directory = demo_directory(root, tenant_id, project_id)
    manifest = directory / "manifest.json"
    if not manifest.is_file():
        return None
    try:
        config = json.loads(manifest.read_text(encoding="utf-8"))
        reference = directory / "result.xlsx"
        if not reference.is_file() or file_digest(reference) != config["sha256"]:
            raise ValueError("changed reference")
        filename = str(config["filename"])
        if Path(filename).name != filename or not filename.endswith(".xlsx"):
            raise ValueError("invalid filename")
        return {"filename": filename, "sha256": config["sha256"], "_reference_path": str(reference)}
    except (ValueError, OSError, KeyError, TypeError) as exc:
        raise HTTPException(status_code=409, detail="演示成品文件缺失或已改变，请重新配置") from exc


def finish_demo(run: dict[str, Any], config: dict[str, Any], destination: Path) -> None:
    if not run.get("plan_confirmation", {}).get("confirmed"):
        raise HTTPException(status_code=409, detail="请先确认演示计划")
    reference = Path(config["_reference_path"])
    if reference.resolve() == destination.resolve() or (destination.exists() and reference.samefile(destination)):
        raise HTTPException(status_code=409, detail="演示输出不能覆盖预置成品")
    if not reference.is_file() or file_digest(reference) != config["sha256"]:
        raise HTTPException(status_code=409, detail="演示成品已变化，不能继续输出")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(reference, destination)
    if file_digest(destination) != config["sha256"]:
        raise HTTPException(status_code=409, detail="演示成品复制校验失败")
    run.update(
        execution_mode="demo", status="completed", draft_filename=destination.name,
        demo_result={"filename": config["filename"], "sha256": config["sha256"],
                     "provenance": "user_supplied_completed_workbook"},
        validation={"status": "demo_reference_match", "detail": "与用户指定成品逐字节一致；不代表真实Agent计算验收"},
        detail="演示流程完成，已准备指定成品文件，可下载展示。此批次不计入真实Agent验收。",
    )
