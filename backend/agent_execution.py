"""Run-level model controller. Never invokes the legacy integration pipeline."""
from __future__ import annotations

import json
import os
import re
import threading
from pathlib import Path
from typing import Any, Callable

import openpyxl

from backend.database import DATA_DIR
from backend.keyuan_workflow import detect_keyuan_batch, execute_keyuan_batch, KeyuanWorkflowError
from backend.natural_language_rules import apply_supported_workbook_rules, parse_supported_workbook_rule
from backend.project_agent_skills import render_project_skill_context, select_project_skills
from core.document_agent.contracts import ToolCall
from core.document_agent.model import ModelConfig, OpenAICompatibleProvider
from core.document_agent.orchestrator import ModelOrchestrator, ToolExecutionError, ToolRegistry
from core.document_agent.workbook_session import (
    WorkbookSession,
    formula_divisor_schema,
    month_roll_forward_schema,
    row_copy_schema,
    row_delete_schema,
    source_write_schema,
)


def _deduplicate_tool_schemas(
    schemas: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Return one provider tool definition per function name.

    OpenAI-compatible APIs reject a request before model execution when the
    ``tools`` array contains duplicate names.  Keep the first definition so
    the caller's explicit contract remains authoritative, but fail loudly if
    two definitions for the same name actually differ.  A silent overwrite
    would make the model see a contract different from the registry it calls.
    """
    unique: list[dict[str, Any]] = []
    fingerprints: dict[str, str] = {}
    for schema in schemas:
        if not isinstance(schema, dict):
            raise TypeError("工具 schema 必须是对象")
        function = schema.get("function")
        name = function.get("name") if isinstance(function, dict) else None
        if not isinstance(name, str) or not name.strip():
            raise ValueError("工具 schema 缺少有效的 function.name")
        fingerprint = json.dumps(schema, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
        previous = fingerprints.get(name)
        if previous is None:
            fingerprints[name] = fingerprint
            unique.append(schema)
        elif previous != fingerprint:
            raise ValueError(f"重复工具名称且定义冲突: {name}")
    return unique

# ---------------------------------------------------------------------------
# 协作式停止注册表：HTTP 停止端点写入标志，后台 worker 在安全的检查点
# （每轮模型调用前、每个处理分组之间）读取并退出。写入本进程内存即可，
# 因为 worker 与端点运行在同一后端进程；跨进程重启场景由既有的
# stale-processing 恢复逻辑兜底。
# ---------------------------------------------------------------------------
_STOP_REQUESTED: set[str] = set()
_STOP_REQUESTED_LOCK = threading.Lock()


def request_run_stop(run_id: str) -> None:
    with _STOP_REQUESTED_LOCK:
        _STOP_REQUESTED.add(run_id)


def run_stop_requested(run_id: str) -> bool:
    with _STOP_REQUESTED_LOCK:
        return run_id in _STOP_REQUESTED


def clear_run_stop(run_id: str) -> None:
    with _STOP_REQUESTED_LOCK:
        _STOP_REQUESTED.discard(run_id)


_BASIC_SOURCE_PREFIXES = ("奖金", "考勤", "值班", "补发补扣", "调差")

# 用户明确限制处理范围（如“只处理派遣”“其他sheet不要动”“司机、外包不处理”）的信号。
# 模式是通用的：不依赖任何客户、总表或Sheet的特定名称；命中时固定全量流程
# （科园批次、基础薪资处理器）不得运行，改由模型按含范围约束的计划执行，
# 保证范围外的 sheet 一个格子都不被改动。
_SCOPE_RESTRICTION_RE = re.compile(
    r"(只|仅)(?:需要|要)?(?:处理|更新|写入|整理|核对|修改|改|动|把|将)"
    r"|[^。；;\n]{0,16}(?:不要|别|不能|不用|不)(?:动|处理|改|写|碰|更新)"
    r"|(?:只要|仅需)[^。；;\n]{0,16}(?:人员|数据|表|sheet|页|部分|内容)"
)


def user_scope_restriction(run: dict[str, Any]) -> str | None:
    """Return the user's explicit scope-restriction text, if any.

    Checks both the formal instruction and the user turns of the alignment
    conversation, so a restriction agreed in chat binds the fixed workflows
    exactly like one written in the instruction.
    """
    texts = [str(run.get("instruction") or "")]
    conversation = run.get("conversation") or []
    if isinstance(conversation, list):
        for message in conversation[-20:]:
            if isinstance(message, dict) and message.get("role") == "user":
                texts.append(str(message.get("content") or ""))
    for text in texts:
        if text and _SCOPE_RESTRICTION_RE.search(text):
            return text
    return None


def find_basic_salary_source(source_paths: dict[str, str] | None) -> tuple[str, Path] | None:
    """Return the first allow-listed workbook that contains basic payroll sheets."""
    candidates = list((source_paths or {}).items())
    # Uploaded legacy tax attachments retain an original name containing
    # ``工资薪金`` while their normalized physical path is also ``.xlsx``.
    # That used to make them win the name-only ordering and sent the model into
    # a repeated "missing bonus/attendance" investigation.  A workbook whose
    # name explicitly identifies the monthly salary source must win first;
    # generic workbooks are then ranked by their actual sheet contents.
    def priority(item: tuple[str, str]) -> tuple[int, str]:
        name = Path(str(item[0])).name
        if "薪资数据" in name or name.startswith("薪资"):
            return (0, name)
        if "工资薪金" in name or "税款计算" in name:
            return (2, name)
        if "工资" in name:
            return (1, name)
        return (3, name)

    candidates.sort(key=priority)
    for name, raw_path in candidates:
        path = Path(str(raw_path))
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            continue
        name_hint = "薪资" in str(name) or ("工资" in str(name) and "税款" not in str(name))
        try:
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            matched = any(
                any(sheet_name == prefix or sheet_name.startswith(prefix) for prefix in _BASIC_SOURCE_PREFIXES)
                for sheet_name in workbook.sheetnames
            )
            workbook.close()
        except (OSError, ValueError, KeyError):
            continue
        if matched or name_hint:
            return str(name), path
    return None


def _execution_max_turns() -> int:
    """单段模型轮数上限，来自 AGENT_MAX_TURNS 环境变量，默认 24。"""
    try:
        return max(1, int(os.getenv("AGENT_MAX_TURNS", "24")))
    except ValueError:
        return 24


def _machine_execution_state(run: dict[str, Any]) -> dict[str, Any]:
    """Build the durable, bounded state passed between execution batches."""
    updates = [item for item in (run.get("workbook_updates") or []) if isinstance(item, dict)]
    unresolved = [
        {key: item.get(key) for key in ("id", "person_key", "person_name", "issue_type", "target_sheet", "target_cell", "status")}
        for item in (run.get("items") or []) if isinstance(item, dict)
        and item.get("status") in {"pending", "needs_review", "failed"}
    ]
    mappings = [
        {key: mapping.get(key) for key in ("source_file", "source_sheet", "target_sheet")}
        for mapping in (run.get("sheet_mappings") or []) if isinstance(mapping, dict)
    ]
    validation = run.get("validation") if isinstance(run.get("validation"), dict) else {}
    checkpoint = run.get("execution_checkpoint") if isinstance(run.get("execution_checkpoint"), dict) else {}
    business_updates = [
        update for update in updates
        if str(update.get("rule") or "") not in {
            "roll_forward_month", "rollback_work_item", "retry_rollback",
        }
    ]
    return {
        "current_stage": "writing" if updates else "reading",
        "completed_steps": list(checkpoint.get("completed_steps") or []),
        "completed_writes": [
            {key: update.get(key) for key in ("target_sheet", "target_cell", "source_file", "source_sheet", "source_cells", "new_value")}
            for update in updates[-500:]
        ],
        "business_write_count": len(business_updates),
        "pending_writes": list(checkpoint.get("pending_writes") or []),
        "unresolved_items": unresolved[:100],
        "source_target_mappings": mappings[:100],
        "validation_failures": list(checkpoint.get("validation_failures") or validation.get("failures") or []),
        "last_read_checkpoint": checkpoint.get("last_read_checkpoint"),
        "last_write_checkpoint": checkpoint.get("last_write_checkpoint"),
        "required_validation_tools": list(checkpoint.get("required_validation_tools") or []),
        "skill_state": dict(checkpoint.get("skill_state") or {}),
    }


def _execution_context(run: dict[str, Any], materials: list[dict[str, Any]], rules: dict[str, Any]) -> dict[str, Any]:
    """Stage-specific input; excludes raw conversation and historical reads."""
    plan = run.get("model_plan") if isinstance(run.get("model_plan"), dict) else {}
    validation = run.get("validation") if isinstance(run.get("validation"), dict) else {}
    if str((run.get("workflow") or {}).get("stage") or "") == "validating" or validation.get("failures"):
        stage = "validating"
    elif run.get("workbook_updates"):
        stage = "writing"
    else:
        stage = "reading"
    current_batch: dict[str, Any] = {"stage": stage}
    checkpoint = run.get("execution_checkpoint") if isinstance(run.get("execution_checkpoint"), dict) else {}
    if checkpoint.get("skill_state"):
        current_batch["skill_state"] = {
            key: checkpoint["skill_state"].get(key)
            for key in ("selected_routes", "runtime_routes", "reference_only_routes", "required_validation_tools")
            if key in checkpoint["skill_state"]
        }
    persisted_checkpoints = [
        {
            key: value for key, value in item.items()
            if key != "archive_content"
        }
        for item in (checkpoint.get("checkpoints") or [])
        if isinstance(item, dict)
    ][-8:]
    if stage == "reading":
        current_batch.update({
            "files": [
                {key: item.get(key) for key in ("filename", "role", "sheets")}
                for item in (run.get("file_manifest") or [])[:50] if isinstance(item, dict)
            ],
            "materials": materials[:20],
            "rule_packages": rules,
            "project_agent_skills": render_project_skill_context(),
        })
    elif stage == "writing":
        current_batch.update({
            "source_target_mappings": _machine_execution_state(run)["source_target_mappings"],
            "pending_writes": _machine_execution_state(run)["pending_writes"],
            "draft_status": {"filename": run.get("draft_filename"), "write_count": len(run.get("workbook_updates") or [])},
        })
    else:
        current_batch.update({
            "completed_writes": _machine_execution_state(run)["completed_writes"],
            "validation_failures": list(validation.get("failures") or []),
            "draft_status": {"filename": run.get("draft_filename")},
        })
    # Keep the execution message stage-local. Replaying the complete plan,
    # materials and conversation here was the main source of context growth.
    # The user instruction is needed while reading to establish scope, but
    # writing/validation receive only machine-readable evidence and state.
    if stage == "reading":
        current_batch["instruction"] = str(run.get("instruction") or "")[:6_000]
        # Keep the compact skill policy at the reading boundary for existing
        # clients; writing and validation intentionally omit it.
        return {
            "current_batch": current_batch,
            "state": _machine_execution_state(run),
            "checkpoints": persisted_checkpoints,
            "project_agent_skills": render_project_skill_context(),
        }
    return {
        "current_batch": current_batch,
        "state": _machine_execution_state(run),
    }


def execute_model_plan(
    run: dict[str, Any], *, draft: Path, registry: ToolRegistry,
    read_schemas: list[dict[str, Any]], materials: list[dict[str, Any]],
    rules: dict[str, Any], save: Callable[[dict[str, Any]], None],
    emit: Callable[..., None],
) -> None:
    config = ModelConfig.from_env()
    fallback_config = ModelConfig.fallback_from_env()
    session = WorkbookSession(Path(run["_master_path"]), draft,
                              {name: Path(path) for name, path in run["_source_paths"].items()})

    def require_month_roll_forward() -> None:
        if run.get("requires_month_roll_forward") and not run.get("month_roll_forward"):
            raise ToolExecutionError("跨月 Run 必须先执行月份滚动，冻结上月数据后才能写入本月数据")

    def month_label_matches(value: Any, period: str) -> bool:
        target = re.fullmatch(r"(20\d{2})[.-](0[1-9]|1[0-2])", str(period or "").strip())
        if target is None:
            return False
        text = re.sub(r"\s+", "", str(value or ""))
        explicit = re.search(r"(20\d{2})\D*(0?[1-9]|1[0-2])月?", text)
        if explicit:
            return int(explicit.group(1)) == int(target.group(1)) and int(explicit.group(2)) == int(target.group(2))
        short = re.fullmatch(r"(0?[1-9]|1[0-2])月?", text)
        return bool(short and int(short.group(1)) == int(target.group(2)))

    def reject_protected_history_targets(updates: list[dict[str, Any]]) -> None:
        rolled = run.get("month_roll_forward")
        if not isinstance(rolled, dict):
            return
        protected_sheet = str(rolled.get("sheet") or "")
        protected_row = int(rolled.get("history_row") or 0)
        for update in updates:
            sheet_name = str(update.get("target_sheet") or update.get("sheet") or "")
            coordinate = str(update.get("target_cell") or update.get("cell") or "")
            match = re.fullmatch(r"[A-Z]{1,3}([1-9][0-9]{0,6})", coordinate)
            if sheet_name == protected_sheet and match and int(match.group(1)) >= protected_row:
                raise ToolExecutionError("历史月份行已冻结，禁止再写入")

    def prepare_workbook_copy() -> dict[str, Any]:
        result = session.prepare_copy()
        run["draft_filename"] = draft.name
        save(run)
        emit(run, "progress", {"stage": "copy", "label": "Agent 已请求创建独立草稿；原件不变"})
        return result

    def run_basic_payroll_processor() -> dict[str, Any]:
        """通用基础数据更新脚本：听 Agent 调度，把来源表的标准数据写入总表。

        是否调度由模型根据对齐后的需求判断；本工具不做客户或项目判断，
        模板不匹配时由脚本诚实回报 failed，模型改用读写工具处理。
        """
        if run.get("requires_month_roll_forward"):
            return {"status": "skipped", "reason": "跨月 Run 必须保留已冻结历史，请使用受控滚月和粒度写入工具"}
        previous = run.get("basic_processor")
        if isinstance(previous, dict):
            return previous
        if run.get("workbook_updates"):
            result = {"status": "skipped", "reason": "当前草稿已有写入，避免基础处理器重置已有修改"}
            run["basic_processor"] = result
            save(run)
            return result
        if user_scope_restriction(run) is not None:
            # 数据安全兜底：全量写入与用户范围约束冲突时拒绝执行。
            # 正常情况下模型按提示词就不会在这种需求下调本工具；
            # 若模型误调度，这里保护范围外的 sheet 不被改写。
            result = {"status": "skipped", "reason": "用户限制了处理范围，全量基础处理器与范围约束冲突，请改用读写工具按计划处理"}
            run["basic_processor"] = result
            save(run)
            emit(run, "progress", {
                "stage": "basic_processor",
                "label": "用户限制了处理范围，已拒绝全量基础处理，改为仅按计划处理指定范围",
            })
            return result
        salary = find_basic_salary_source(run.get("_source_paths"))
        if salary is None:
            result = {"status": "skipped", "reason": "当前运行没有匹配的基础薪资来源工作簿（需含奖金、考勤、值班等标准Sheet）"}
            run["basic_processor"] = result
            save(run)
            return result
        try:
            from scripts.keyuan_basic_processor import process_keyuan_basics

            processed = process_keyuan_basics(Path(run["_master_path"]), salary[1], draft)
        except (KeyError, OSError, ValueError, ImportError) as exc:
            result = {"status": "failed", "reason": f"总表模板与基础处理器能力不匹配（{type(exc).__name__}），未写入任何数据，请改用读写工具按计划处理"}
            run["basic_processor"] = result
            save(run)
            emit(run, "progress", {"stage": "basic_processor", "label": "总表模板与基础处理器能力不匹配，未写入数据"})
            return result
        updates = processed.get("changes") if isinstance(processed, dict) else []
        if not isinstance(updates, list):
            updates = []
        run.setdefault("workbook_updates", []).extend(updates)
        result = {
            "status": str(processed.get("status") or "passed"),
            "change_count": len(updates),
            "issue_count": len(processed.get("issues") or []) if isinstance(processed, dict) else 0,
            "issues": processed.get("issues") if isinstance(processed, dict) else [],
        }
        run["basic_processor"] = result
        save(run)
        emit(run, "progress", {
            "stage": "basic_processor",
            "label": f"基础处理器已完成，写入 {result['change_count']} 项；未决 {result['issue_count']} 项",
        })
        return result

    def run_keyuan_workflow() -> dict[str, Any]:
        """通用完整批次更新脚本：听 Agent 调度，一次性完成标准处理并校验。

        是否调度由模型根据对齐后的需求判断；本工具不做客户或项目判断，
        资料不齐或模板不匹配时诚实回报 skipped/failed 及原因。
        """
        if run.get("requires_month_roll_forward"):
            return {"status": "skipped", "reason": "跨月 Run 必须保留已冻结历史，请使用受控滚月和粒度写入工具"}
        batch = detect_keyuan_batch(str(run.get("master_file") or ""), run.get("_source_paths"))
        if batch is None:
            return {"status": "skipped", "reason": "当前批次资料不满足完整处理条件（总表模板或来源文件不匹配）"}
        if user_scope_restriction(run) is not None:
            # 数据安全兜底：批次是全量原子处理，无法只写部分 sheet；
            # 用户限制了范围时拒绝执行，保护范围外的 sheet 不被改写。
            emit(run, "progress", {
                "stage": "keyuan_workflow",
                "label": "用户限制了处理范围，已拒绝全量批次流程，改为仅按计划处理指定范围",
            })
            return {"status": "skipped", "reason": "用户限制了处理范围，全量批次流程与范围约束冲突，请改用读写工具按计划处理"}
        if not batch.matches_project_month(str(run.get("salary_month") or "")):
            # The workbook filename and complete source set are authoritative
            # for this verified batch.  A stale project-month label must not
            # divert the run into an unbounded model read loop; keep the
            # mismatch visible in the audit trail and continue deterministically.
            emit(run, "progress", {
                "stage": "keyuan_workflow",
                "label": f"检测到项目月份 {run.get('salary_month')} 与文件所属月 {batch.payroll_period} 不一致，按文件批次继续处理",
            })
        previous = run.get("keyuan_workflow")
        if isinstance(previous, dict) and previous.get("status") in {"passed", "needs_review"}:
            return previous
        try:
            result = execute_keyuan_batch(
                batch,
                master_path=Path(run["_master_path"]),
                source_paths=run.get("_source_paths") or {},
                output_path=draft,
                work_root=Path(DATA_DIR) / "keyuan-workflows",
            )
        except KeyuanWorkflowError as exc:
            result = {"status": "failed", "reason": str(exc), "issues": []}
        run["keyuan_workflow"] = result
        updates = result.get("changes") if isinstance(result, dict) else []
        if isinstance(updates, list):
            run.setdefault("workbook_updates", []).extend(updates)
        run["draft_filename"] = draft.name if draft.is_file() else run.get("draft_filename")
        save(run)
        emit(run, "progress", {
            "stage": "keyuan_workflow",
            "label": (
                f"科园完整流程已完成，写入 {result.get('change_count', 0)} 项；"
                f"未决 {len(result.get('issues') or [])} 项"
                if result.get("status") in {"passed", "needs_review"}
                else str(result.get("reason") or "科园流程未完成")
            ),
        })
        return result

    def guard_unique_targets(updates: list[dict[str, Any]], *, operation: str) -> None:
        """Reject cross-batch writes to a target already changed in this run."""
        reject_protected_history_targets(updates)
        prior_targets = {
            (str(update.get("target_sheet") or update.get("sheet") or ""),
             str(update.get("target_cell") or update.get("cell") or ""))
            for update in (run.get("workbook_updates") or [])
            if isinstance(update, dict)
        }
        incoming_targets: set[tuple[str, str]] = set()
        for update in updates:
            target = (
                str(update.get("target_sheet") or update.get("sheet") or ""),
                str(update.get("target_cell") or update.get("cell") or ""),
            )
            if target in incoming_targets or target in prior_targets:
                raise ToolExecutionError(f"{operation}包含已写入的重复目标，已拒绝本批")
            incoming_targets.add(target)

    def apply_source_cells(changes: list[dict[str, Any]]) -> dict[str, Any]:
        require_month_roll_forward()
        guard_unique_targets(list(changes), operation="来源写入")
        updates = session.apply_source_cells(changes)
        guard_unique_targets(updates, operation="来源写入")
        run.setdefault("workbook_updates", []).extend(updates)
        save(run)
        emit(run, "progress", {"stage": "source_write", "label": f"Agent 已核对并写入 {len(updates)} 个来源单元格"})
        return {"updates": updates}

    def apply_formula_divisors(changes: list[dict[str, Any]]) -> dict[str, Any]:
        require_month_roll_forward()
        guard_unique_targets(list(changes), operation="公式写入")
        updates = session.apply_formula_divisors(changes)
        guard_unique_targets(updates, operation="公式写入")
        run.setdefault("workbook_updates", []).extend(updates)
        save(run)
        emit(run, "progress", {"stage": "formula_write", "label": f"Agent 已核对并更新 {len(updates)} 个公式参数"})
        return {"updates": updates}

    def insert_and_copy_row(change: dict[str, Any]) -> dict[str, Any]:
        require_month_roll_forward()
        rolled = run.get("month_roll_forward") if isinstance(run.get("month_roll_forward"), dict) else {}
        if rolled and str(change.get("sheet") or "") == str(rolled.get("sheet") or ""):
            raise ToolExecutionError("历史月份及其汇总表结构已冻结，月份滚动后禁止再插行")
        fingerprint = json.dumps(change, ensure_ascii=False, sort_keys=True, default=str)
        prior_fingerprints = set(run.get("write_fingerprints") or [])
        if fingerprint in prior_fingerprints:
            raise ToolExecutionError("相同插行请求已在本次运行执行，禁止重复插入")
        update = session.insert_and_copy_row(change)
        run.setdefault("workbook_updates", []).append(update)
        run.setdefault("write_fingerprints", []).append(fingerprint)
        save(run)
        emit(run, "progress", {"stage": "row_insert", "label": f"Agent 已在 {update['sheet']} 插入并复制第 {update['inserted_row']} 行"})
        return update

    def delete_rows(change: dict[str, Any]) -> dict[str, Any]:
        require_month_roll_forward()
        rolled = run.get("month_roll_forward") if isinstance(run.get("month_roll_forward"), dict) else {}
        if rolled and str(change.get("sheet") or "") == str(rolled.get("sheet") or ""):
            raise ToolExecutionError("历史月份及其汇总表结构已冻结，月份滚动后禁止删除")
        fingerprint = json.dumps(change, ensure_ascii=False, sort_keys=True, default=str)
        prior_fingerprints = set(run.get("write_fingerprints") or [])
        if fingerprint in prior_fingerprints:
            raise ToolExecutionError("相同行删除请求已在本次运行执行，禁止重复删除")
        update = session.delete_rows(change)
        run.setdefault("workbook_updates", []).append(update)
        run.setdefault("write_fingerprints", []).append(fingerprint)
        save(run)
        emit(run, "progress", {"stage": "row_delete", "label": f"Agent 已在 {update['sheet']} 删除 {len(update['deleted_rows'])} 行"})
        return update

    def roll_forward_month(change: dict[str, Any]) -> dict[str, Any]:
        if not run.get("requires_month_roll_forward"):
            raise ToolExecutionError("当前 Run 不是从上一已验收月份滚动，不得新增月份行")
        if run.get("month_roll_forward"):
            raise ToolExecutionError("本次 Run 已完成月份滚动，禁止重复新增")
        if run.get("workbook_updates"):
            raise ToolExecutionError("月份滚动必须是跨月 Run 的第一个数据写入")
        target_month = str(run.get("target_salary_month") or run.get("salary_month") or "")
        baseline_month = str((run.get("baseline") or {}).get("salary_month") or "")
        if not month_label_matches(change.get("new_month"), target_month):
            raise ToolExecutionError("新月份标签与本次 target_salary_month 不一致")
        if not month_label_matches(change.get("expected_current_month"), baseline_month):
            raise ToolExecutionError("当前月份行与上一已验收月份不一致")
        if not draft.is_file():
            prepare_workbook_copy()
        update = session.roll_forward_summary_row(change)
        run["month_roll_forward"] = {
            **update, "target_salary_month": target_month,
            "baseline_sha256": str((run.get("baseline") or {}).get("sha256") or ""),
        }
        run.setdefault("workbook_updates", []).append({
            "rule": "roll_forward_month", "sheet": update["sheet"],
            "inserted_row": update["current_row"], "history_row": update["history_row"],
            "new_month": update["new_month"], "history_values": update["history_values"],
            "history_snapshot": update["history_snapshot"],
        })
        save(run)
        emit(run, "progress", {
            "stage": "month_roll_forward",
            "label": f"已新增 {update['new_month']} 并冻结上月历史行",
        })
        # Natural-language rules are deliberately deferred until the new
        # month has been created.  Apply the deferred deterministic rule now,
        # in a separate atomic batch, so a cross-month roster instruction does
        # not silently end after inserting only the month row.
        if run.get("deferred_instruction_rule"):
            deferred_result = apply_instruction_rules()
            if any(
                result.get("status") == "needs_review"
                for result in deferred_result.get("results", [])
                if isinstance(result, dict)
            ):
                run["deferred_instruction_rule_result"] = deferred_result
                save(run)
            elif deferred_result.get("results"):
                run["deferred_instruction_rule_result"] = deferred_result
                save(run)
            if deferred_result.get("results"):
                update = {**update, "deferred_rule": deferred_result}
        return update

    def apply_instruction_rules() -> dict[str, Any]:
        """Run explicit, supported natural-language rules after the base pass."""
        if parse_supported_workbook_rule(str(run.get("instruction") or "")) is None:
            return {"results": [], "updates": []}
        if run.get("requires_month_roll_forward") and not run.get("month_roll_forward"):
            deferred = {
                "reason": "跨月 Run 必须先完成月份滚动",
                "requires": "roll_forward_month",
            }
            run["deferred_instruction_rule"] = deferred
            save(run)
            emit(run, "progress", {
                "stage": "instruction_rule_deferred",
                "label": "自然语言写入规则已延后，先新增本月并冻结上月数据",
            })
            return {"results": [], "updates": [], "deferred": deferred}
        run.pop("deferred_instruction_rule", None)
        if not draft.is_file():
            prepare_workbook_copy()
        outcome = apply_supported_workbook_rules(
            instruction=str(run.get("instruction") or ""),
            draft_path=draft,
            source_paths=run.get("_source_paths") or {},
            previous_results=run.get("natural_language_rule_results") or [],
            target_period=str(run.get("salary_month") or ""),
        )
        results = list(outcome.get("results") or [])
        updates = list(outcome.get("updates") or [])
        history = list(run.get("natural_language_rule_results") or [])
        prior_ids = {str(item.get("rule_id") or "") for item in history if isinstance(item, dict)}
        history.extend(item for item in results if item.get("rule_id") not in prior_ids)
        run["natural_language_rule_results"] = history
        if updates:
            run.setdefault("workbook_updates", []).extend(updates)
        save(run)
        for result in results:
            if result.get("status") == "applied":
                if result.get("kind") == "source_sheet_roster_sync":
                    run["roster_sync"] = result
                emit(run, "progress", {"stage": "natural_language_rule", "label": str(result.get("detail") or "已执行自然语言规则")})
            elif result.get("status") == "needs_review":
                emit(run, "needs_user_input", {"code": "NATURAL_LANGUAGE_RULE_NEEDS_REVIEW", "detail": str(result.get("detail") or "自然语言规则无法安全执行")})
        return {"results": results, "updates": updates}

    registry.register("prepare_workbook_copy", prepare_workbook_copy)
    registry.register("run_basic_payroll_processor", run_basic_payroll_processor)
    registry.register("run_keyuan_workflow", run_keyuan_workflow)
    registry.register("apply_source_cells", apply_source_cells)
    registry.register("apply_formula_divisors", apply_formula_divisors)
    registry.register("insert_and_copy_row", insert_and_copy_row)
    registry.register("delete_rows", delete_rows)
    registry.register("roll_forward_month", roll_forward_month)
    # Apply the execution-phase allow-list only after all handlers have been
    # registered.  Keeping registration and authorization separate avoids a
    # fragile "restrict then re-register" ordering dependency while ensuring
    # the model can call exactly the read and write tools advertised below.
    execution_tool_names = {
        schema["function"]["name"] for schema in read_schemas
    } | {
        "prepare_workbook_copy",
        "run_basic_payroll_processor",
        "run_keyuan_workflow",
        "apply_source_cells",
        "apply_formula_divisors",
        "insert_and_copy_row",
        "delete_rows",
        "roll_forward_month",
        "validate_workbook",
        "validate_with_officecli",
    }
    registry.restrict_to(execution_tool_names)
    # 固定处理脚本（基础处理器、科园完整批次）不再自动执行。它们已注册
    # 为 Agent 工具，由模型在理解对齐后的需求与计划基础上自主决定是否
    # 调用；通用任务全程使用读写工具，不会被固定流程劫持。
    skill_state = select_project_skills(run)
    checkpoint = run.setdefault("execution_checkpoint", {})
    checkpoint["skill_state"] = skill_state
    emit(run, "progress", {
        "stage": "skill_selection",
        "selected_routes": skill_state.get("selected_routes", []),
        "runtime_routes": skill_state.get("runtime_routes", []),
        "reference_only_routes": skill_state.get("reference_only_routes", []),
        "required_validation_tools": skill_state.get("required_validation_tools", []),
        "label": "已按任务内容选择项目 Skill，并登记运行时约束与验收门槛",
    })
    workbook_route = next(
        (
            route for route in skill_state.get("routes", [])
            if isinstance(route, dict) and route.get("id") == "workbook-update"
        ),
        {},
    )
    required_validation_tools = list(workbook_route.get("required_validation_tools") or [])
    checkpoint["required_validation_tools"] = required_validation_tools
    instruction_rules = apply_instruction_rules()
    rule_needs_review = any(result.get("status") == "needs_review" for result in instruction_rules["results"])
    if rule_needs_review:
        review_detail = "自然语言规则缺少可验证的信息，未写入草稿；请按待处理说明补充。"
        run.update(
            status="awaiting_review",
            detail=review_detail,
            execution_result={"status": "needs_review", "code": "NATURAL_LANGUAGE_RULE_NEEDS_REVIEW", "content": review_detail},
            validation={"status": "not_verified", "detail": "自然语言规则未安全执行，草稿未发布"},
        )
        save(run)
        return
    roster_rule_applied = next(
        (
            result for result in instruction_rules["results"]
            if result.get("kind") == "source_sheet_roster_sync" and result.get("status") == "applied"
        ),
        None,
    )
    if roster_rule_applied is not None:
        write_count = max(0, int(roster_rule_applied.get("change_count") or 0))
        emit(run, "tool_result", {
            "call_id": "deterministic-roster-sync",
            "name": "source_sheet_roster_sync",
            "status": "succeeded",
            "write_count": write_count,
            "business_write": write_count > 0,
        })
        completed_steps: list[str] = []
        validation_failures: list[dict[str, str]] = []
        checkpoint.update({
            "current_stage": "validating",
            "completed_steps": completed_steps,
            "validation_failures": validation_failures,
        })
        for validation_name in required_validation_tools:
            call_id = f"deterministic-{validation_name}"
            emit(run, "tool_call", {"call_id": call_id, "name": validation_name})
            result = registry.execute(ToolCall(call_id=call_id, name=validation_name))
            output = result.output if isinstance(result.output, dict) else {}
            if validation_name == "validate_workbook":
                passed = (
                    result.status == "succeeded"
                    and output.get("readable") is True
                    and output.get("can_publish") is True
                    and not output.get("formula_errors")
                    and not output.get("summary_range_errors")
                    and not output.get("duplicate_identities")
                    and not output.get("identity_errors")
                )
            elif validation_name == "validate_with_officecli":
                passed = (
                    result.status == "succeeded"
                    and output.get("valid") is True
                    and output.get("recalculated") is True
                    and not output.get("formula_errors")
                    and not output.get("unevaluated_formulas")
                )
            else:
                passed = False
            error = result.error or (None if passed else "校验工具返回未通过")
            emit(run, "tool_result", {
                "call_id": call_id,
                "name": validation_name,
                "status": "succeeded" if passed else "failed",
                "error": error,
            })
            if passed:
                completed_steps.append(validation_name)
            else:
                validation_failures.append({
                    "tool": validation_name,
                    "error": str(error or "校验工具返回未通过"),
                })
        if validation_failures:
            detail = "受控名单同步已写入草稿，但 Skill 要求的结构校验或公式重算未通过，不能标记为完成。"
            run.update(
                status="awaiting_review",
                detail=detail,
                execution_result={
                    "status": "execution_incomplete",
                    "code": "VALIDATION_FAILED",
                    "content": detail,
                },
                validation={
                    "status": "failed",
                    "detail": detail,
                    "failures": validation_failures,
                },
            )
            save(run)
            return
        checkpoint["current_stage"] = "completed"
        run.update(
            status="awaiting_review",
            detail=str(roster_rule_applied.get("detail") or "已按来源名单完成受控同步；等待验收。"),
            execution_result={
                "status": "completed",
                "code": "SOURCE_ROSTER_SYNC_APPLIED",
                "content": str(roster_rule_applied.get("detail") or ""),
            },
            validation={"status": "not_verified", "detail": "受控同步已完成，草稿仍需人工验收后发布"},
        )
        save(run)
        return
    if config is None and fallback_config is None:
        run.update(status="blocked", detail="尚未配置模型服务，未执行工作簿更新")
        save(run)
        return
    additions = [source_write_schema(), formula_divisor_schema(), row_copy_schema(), row_delete_schema(), month_roll_forward_schema(), {"type": "function", "function": {
        "name": "prepare_workbook_copy", "description": "创建原始总表的独立副本；重复调用不重置草稿。",
        "parameters": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
    }}, {"type": "function", "function": {
        "name": "run_basic_payroll_processor", "description": "调度内置的基础数据更新脚本：把来源工作簿中的奖金、考勤、值班、补发补扣等标准数据全量写入总表。脚本是通用执行器，是否调度由你根据对齐后的需求判断：仅当需求是标准全量处理、用户未限制范围、且总表模板与脚本能力匹配时调用；模板不匹配会返回 failed 及原因，届时改用读写工具按计划处理。用户限制了处理范围时禁止调用。",
        "parameters": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
    }}, {"type": "function", "function": {
        "name": "run_keyuan_workflow", "description": "调度内置的完整批次更新脚本：一次性完成全部标准处理（基础数据、社保、个税）并校验。脚本是通用执行器，是否调度由你根据对齐后的需求判断：仅当需求是完整批次全量处理、用户未限制范围、且总表模板与来源资料齐全匹配时调用；返回 skipped 或 failed 时是其能力边界的诚实回报，按原因改用读写工具处理，不要重试。",
        "parameters": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
    }}, {"type": "function", "function": {
        "name": "validate_workbook", "description": "对当前独立草稿执行结构、公式、重复身份和未解决事项校验；必须返回通过才算完成。",
        "parameters": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
    }}, {"type": "function", "function": {
        "name": "validate_with_officecli", "description": "使用已安装的 OfficeCLI 对当前独立草稿执行结构校验和真实公式重算，并检查是否新增公式错误或无法求值公式；失败必须报告，不能伪造通过。",
        "parameters": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
    }}]
    # Validation schemas are already part of the router's read contract for
    # generic runs.  Do not append a second definition with the same name;
    # duplicate names make the provider reject the request with HTTP 400
    # before the model has a chance to execute any tool.
    read_names = {
        schema.get("function", {}).get("name")
        for schema in read_schemas
        if isinstance(schema, dict) and isinstance(schema.get("function"), dict)
    }
    schemas = _deduplicate_tool_schemas([
        *read_schemas,
        *[
            schema for schema in additions
            if schema.get("function", {}).get("name") not in read_names
        ],
    ])
    context = _execution_context(run, materials, rules)
    run["status"] = "processing"
    save(run)
    primary_config = config or fallback_config
    assert primary_config is not None
    primary_provider = OpenAICompatibleProvider(primary_config.model_copy(update={"max_output_tokens": max(primary_config.max_output_tokens, 4096)}))
    fallback_provider = (
        OpenAICompatibleProvider(fallback_config.model_copy(update={"max_output_tokens": max(fallback_config.max_output_tokens, 4096)}))
        if fallback_config and config else None
    )
    def persist_checkpoint(checkpoint: dict[str, Any]) -> None:
        state = run.setdefault("execution_checkpoint", {})
        if checkpoint.get("stage") == "writing":
            state["last_write_checkpoint"] = checkpoint.get("checkpoint_id")
            state["last_write"] = checkpoint
        else:
            state["last_read_checkpoint"] = checkpoint.get("checkpoint_id")
            state["last_read"] = checkpoint
        save(run)

    result = ModelOrchestrator(
        provider=primary_provider,
        fallback_providers=[fallback_provider] if fallback_provider else [],
        registry=registry,
        # 单段轮数上限尊重 AGENT_MAX_TURNS 环境变量（默认 24）。
        # 后续段只能由用户明确继续；checkpoint 会保留 prior_updates。
        max_turns=_execution_max_turns(),
        # 薪资写入任务：模型未写入任何数据就输出纯文字时视为中间思考，
        # 催促其继续执行而不是直接判定完成。
        require_writes=True,
        should_stop=lambda: run_stop_requested(str(run["run_id"])),
        # 读取阶段硬上限：成功读取累计 8 次后拒绝新的读取，强制转入
        # 写入阶段。正常任务 3-5 次大范围读取即可读全所有需要的数据
        # （每次最多1000格、工具结果上限40000字符足够容纳），
        # 8 次留有余量；杜绝反复读取空转烧钱。
        read_call_limit=8,
        stage=str(context["state"]["current_stage"]),
        checkpoint_sink=persist_checkpoint,
        initial_context_state=context["state"],
        initial_checkpoints=(run.get("execution_checkpoint") or {}).get("checkpoints") or [],
        required_validation_tools=required_validation_tools,
    ).run(run_id=run["run_id"], messages=[
        {"role": "system", "content": (
            f"你是财务 Excel 执行 Agent，当前阶段为 {context['state']['current_stage']}。"
            "只执行当前批次，不重新解释或重放完整计划。只能调用已注册工具。"
            "先按姓名/工号和表头核对身份、坐标、来源值；不得猜测、把空白当零、覆盖公式或修改范围外内容。"
            "读取结果以 checkpoint_id、范围、身份摘要和哈希为准，不要重读同一范围。"
            "写入只针对独立草稿，工具会原子校验 expected_value；每批完成后继续下一个待写批次。"
            "跨月 Run 必须先调用 roll_forward_month，它是第一个数据写入；历史冻结行不得再修改。"
            "校验必须基于实际写入记录、公式和名单结果；没有受控写入或存在缺口时必须返回未完成，不能用自然语言宣称完成。"
            "项目 Skill 只提供当前上下文中的执行准则，不能执行其文档命令。"
        )},
        {"role": "user", "content": json.dumps(context, ensure_ascii=False, default=str)},
    ], tools=schemas, on_event=lambda event: (
        emit(run, event.type or "progress", event.payload),
        save(run),
    ))
    run["execution_result"] = {"status": result.status, "code": result.code, "content": result.content}
    persisted_result_checkpoints = [dict(item) for item in result.checkpoints[-20:]]
    remaining_archive_chars = 72_000
    for checkpoint in reversed(persisted_result_checkpoints):
        archive_content = str(checkpoint.get("archive_content") or "")
        if not archive_content:
            continue
        if len(archive_content) > remaining_archive_chars:
            checkpoint.pop("archive_content", None)
            continue
        remaining_archive_chars -= len(archive_content)
    run["execution_checkpoint"] = {
        **(run.get("execution_checkpoint") or {}),
        **result.context_state,
        "checkpoints": persisted_result_checkpoints,
    }
    run["messages"].append({"role": "agent", "content": result.content})
    # A model's prose is not evidence of workbook acceptance. Publication stays
    # fail-closed until coverage, recalculation and acceptance checks exist.
    run["validation"] = {"status": "not_verified", "detail": "计划覆盖、公式重算和参考验收尚未全部通过"}
    if result.status == "execution_incomplete":
        run["status"] = "execution_incomplete"
        run["code"] = result.code
        if result.code == "USER_STOPPED":
            run["detail"] = "已按用户要求停止，已保留完成的写入和事件；可随时续跑"
        elif result.code == "EMPTY_MODEL_RESPONSE":
            run["detail"] = "模型连续返回空响应，系统已自动重试并保留当前进度；可以继续执行，未发布正式结果"
        elif result.code == "MODEL_PROVIDER_ERROR":
            provider_detail = str(result.content or "").strip()
            run["detail"] = (
                f"{provider_detail}；系统已自动重试并保留当前进度；可以继续执行，未发布正式结果"
                if provider_detail
                else "模型服务连续请求失败，系统已自动重试并保留当前进度；可以继续执行，未发布正式结果"
            )
        elif result.code == "CONTEXT_BUDGET_EXCEEDED":
            run["code"] = "CONTEXT_BUDGET_EXCEEDED"
            run["detail"] = "上下文压缩后仍超过安全预算，任务已停止并保留 checkpoint；未继续读取或发布结果"
        elif result.code == "NO_WRITES_PERFORMED":
            run["detail"] = "模型本轮仅分析未写入；系统已保留读取结果，等待你明确继续，未发布正式结果"
        elif result.code == "VALIDATION_FAILED":
            run["detail"] = "草稿的结构、公式、身份或重算校验未通过；已保留草稿和失败证据，不能标记为完成"
        elif result.code == "VALIDATION_NOT_PERFORMED":
            run["detail"] = "缺少 Skill 要求的结构校验或 OfficeCLI 重算记录；已保留草稿，不能标记为完成"
        else:
            run["detail"] = "模型读取达到本轮上限，尚未完成写入或校验；可以继续执行，未发布正式结果"
    else:
        run["status"] = "blocked" if result.status in {"failed", "blocked"} else "awaiting_review"
        run["detail"] = result.content or "本轮未完成执行，请查看运行记录后重试；未发布正式结果"
    save(run)
