from __future__ import annotations

from copy import deepcopy
import hashlib
import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from fastapi import HTTPException
import openpyxl
import pytest

import backend.routers.agent as agent
from backend.models import Project, UploadFile
from core.document_agent.contracts import ToolCall
from core.document_agent.model import ModelConfig, ModelResponse


class _Query:
    def __init__(self, first_value: Any = None, all_values: list[Any] | None = None):
        self.first_value = first_value
        self.all_values = all_values or []

    def filter(self, *_args: Any) -> _Query:
        return self

    def first(self) -> Any:
        return self.first_value

    def all(self) -> list[Any]:
        return self.all_values


class _Db:
    def __init__(self, project: Any, files: list[Any]):
        self.project = project
        self.files = files

    def query(self, model: Any) -> _Query:
        if model is Project:
            return _Query(first_value=self.project)
        if model is UploadFile:
            return _Query(all_values=self.files)
        raise AssertionError(f"Unexpected database model: {model}")


@pytest.fixture
def scenario(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    upload_root = tmp_path / "uploads"
    project_dir = upload_root / "project-1"
    project_dir.mkdir(parents=True)
    files = []
    for filename, role, value in [
        ("总表.xlsx", "financial_master", 0),
        ("奖金来源.xlsx", "financial_source", 2000),
    ]:
        path = project_dir / filename
        workbook = openpyxl.Workbook()
        workbook.active.title = "工资核算"
        workbook.active.append(["工号", "姓名", "奖金"])
        workbook.active.append(["E001", "测试人员", value])
        workbook.save(path)
        workbook.close()
        files.append(SimpleNamespace(
            id=role, original_name=filename, file_type=role,
            stored_path=f"project-1/{filename}", file_size=path.stat().st_size,
        ))
    project = SimpleNamespace(
        id="project-1", salary_month="2026.07",
        owner=SimpleNamespace(tenant_id="tenant-a", tenant=SimpleNamespace(name="测试公司")),
    )
    user = SimpleNamespace(id="user-1", tenant_id="tenant-a")
    calls: list[str] = []

    def legacy_integration(*_args: Any, **_kwargs: Any) -> SimpleNamespace:
        calls.append("legacy_integration")
        return SimpleNamespace(model_dump=lambda: {
            "filename": "legacy-draft.xlsx", "issues": [], "updates": [], "matches": [],
        })

    def release(*_args: Any, **_kwargs: Any) -> SimpleNamespace:
        calls.append("release")
        return SimpleNamespace(model_dump=lambda: {"filename": "published.xlsx"})

    def model_complete(*_args: Any, **_kwargs: Any) -> Any:
        calls.append("model")
        raise AssertionError("An unconfirmed run must not invoke the model")

    monkeypatch.setattr(agent, "UPLOAD_DIR", str(upload_root))
    monkeypatch.setattr(agent, "RUN_DIR", tmp_path / "runs")
    monkeypatch.setattr(agent, "_result_path", lambda _project, filename: tmp_path / "results" / filename)
    monkeypatch.setattr(agent, "_load_material_index", lambda *_args: [])
    monkeypatch.setattr(agent.experience_store, "suggest_memory", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(agent.ModelConfig, "from_env", lambda: None)
    monkeypatch.setattr(agent.OpenAICompatibleProvider, "complete", model_complete)
    monkeypatch.setattr(agent, "create_financial_workbook_integration", legacy_integration)
    monkeypatch.setattr(agent, "release_latest_financial_workbook_integration", release)
    return SimpleNamespace(
        root=tmp_path, user=user, db=_Db(project, files), calls=calls,
        originals={path: path.read_bytes() for path in project_dir.glob("*.xlsx")},
    )


@pytest.mark.parametrize("auto_publish", [False, True])
def test_create_run_only_inventories_files_without_integrating_or_creating_draft(
    scenario: SimpleNamespace, auto_publish: bool,
) -> None:
    result = agent.create_agent_run(
        agent.AgentRunCreateIn(project_id="project-1", auto_publish=auto_publish),
        user=scenario.user, db=scenario.db,
    )

    assert scenario.calls == []
    assert not result.get("draft_filename")
    assert {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")} == scenario.originals


def test_new_run_returns_unconfirmed_plan_and_file_roles(scenario: SimpleNamespace) -> None:
    result = agent.create_agent_run(
        agent.AgentRunCreateIn(project_id="project-1"), user=scenario.user, db=scenario.db,
    )

    assert result["plan_confirmation"]["required"] is True
    assert result["plan_confirmation"]["confirmed"] is False
    assert {entry["filename"]: entry["role"] for entry in result["file_manifest"]} == {
        "总表.xlsx": "financial_master", "奖金来源.xlsx": "financial_source",
    }
    saved = agent._load_run(result["run_id"], scenario.user)
    assert saved["plan_confirmation"] == result["plan_confirmation"]


def test_keyuan_files_use_generic_agent_mode(scenario: SimpleNamespace) -> None:
    scenario.db.project.salary_month = "2026.09"
    scenario.db.files[0].original_name = "202608（所属月202607）-北京科园-鹤安-大药房工资核算总表-v2.xlsx"
    scenario.db.files[1].original_name = "薪资数据-科园-7月薪资（8.14发薪）.xlsx"

    result = agent.create_agent_run(
        agent.AgentRunCreateIn(project_id="project-1"), user=scenario.user, db=scenario.db,
    )

    assert result.get("execution_mode") is None
    assert "自动化流程" not in str(result.get("detail") or "")


@pytest.mark.parametrize("action", ["execute", "publish"])
def test_unconfirmed_plan_blocks_execution_and_publication_even_with_no_pending_items(
    scenario: SimpleNamespace, action: str,
) -> None:
    run_id = "a" * 32
    agent._save_run({
        "run_id": run_id, "tenant_id": "tenant-a", "project_id": "project-1",
        "salary_month": "2026.07", "status": "ready_to_publish", "items": [],
        "plan_confirmation": {"required": True, "confirmed": False},
        "month_confirmation": {"required": False, "confirmed": True},
    })

    with pytest.raises(HTTPException) as raised:
        if action == "execute":
            agent.execute_agent_run(run_id, user=scenario.user)
        else:
            agent.publish_agent_run(
                run_id, agent.AgentPublishIn(confirm_month=True), user=scenario.user, db=scenario.db,
            )

    assert raised.value.status_code == 409
    assert scenario.calls == []
    assert {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")} == scenario.originals


def _prepared_run(scenario: SimpleNamespace, *, confirmed: bool = False) -> dict[str, Any]:
    master = next(path for path in scenario.originals if path.name == "总表.xlsx")
    source = next(path for path in scenario.originals if path.name == "奖金来源.xlsx")
    run = {
        "run_id": "b" * 32, "tenant_id": "tenant-a", "project_id": "project-1",
        "salary_month": "2026.07", "instruction": "按手册核对奖金", "status": "awaiting_review",
        "master_file": master.name, "source_files": [source.name],
        "_master_path": str(master), "_source_paths": {source.name: str(source)},
        "file_manifest": [
            {"filename": path.name, "role": role, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
            for path, role in [(master, "financial_master"), (source, "financial_source")]
        ],
        "plan_confirmation": {"required": True, "confirmed": confirmed},
        "month_confirmation": {"required": False, "confirmed": True},
        "items": [], "events": [], "messages": [],
    }
    agent._save_run(run)
    return run


def _stub_plan_provider(monkeypatch: pytest.MonkeyPatch, plan: dict[str, Any]) -> list[dict[str, Any]]:
    captured: list[dict[str, Any]] = []
    config = ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1",
        api_key="test-secret", model="test-planner",
    )

    def complete(_self: Any, **kwargs: Any) -> ModelResponse:
        captured.append(kwargs)
        return ModelResponse(content=json.dumps(plan), request_id="plan-1")

    monkeypatch.setattr(agent.ModelConfig, "from_env", lambda: config)
    monkeypatch.setattr(agent.OpenAICompatibleProvider, "complete", complete)
    return captured


@pytest.mark.parametrize("instruction", [None, "E001奖金取来源表2000元，保留原始总表"])
def test_generating_plan_is_read_only_and_uses_user_answers_without_confirming(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch, instruction: str | None,
) -> None:
    run = _prepared_run(scenario)
    plan = {"summary": "核对并更新副本", "steps": ["核对E001奖金"], "questions": []}
    captured = _stub_plan_provider(monkeypatch, plan)
    result = agent.confirm_agent_plan(
        run["run_id"], agent.AgentPlanIn(confirm_plan=False, instruction=instruction), user=scenario.user,
    )

    assert {key: result["model_plan"][key] for key in plan} == plan
    assert result["plan_confirmation"]["confirmed"] is False
    assert not result.get("draft_filename")
    assert scenario.calls == []
    assert {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")} == scenario.originals
    context = json.loads(captured[0]["messages"][1]["content"])
    assert run["instruction"] in context["instruction"]
    if instruction:
        assert instruction in context["instruction"]
    saved = agent._load_run(run["run_id"], scenario.user)
    assert saved["model_plan"] == result["model_plan"]


@pytest.mark.parametrize("model_plan", [
    None,
    {"summary": "需要澄清", "steps": ["核对奖金"], "questions": ["同一人员两个奖金来源选哪个？"]},
])
def test_plan_confirmation_requires_existing_plan_without_open_questions(
    scenario: SimpleNamespace, model_plan: dict[str, Any] | None,
) -> None:
    run = _prepared_run(scenario)
    if model_plan is not None:
        run["model_plan"] = model_plan
        agent._save_run(run)

    with pytest.raises(HTTPException) as raised:
        agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=True), user=scenario.user)

    assert raised.value.status_code == 409
    assert scenario.calls == []
    assert agent._load_run(run["run_id"], scenario.user)["plan_confirmation"]["confirmed"] is False


def test_explicit_confirmation_accepts_complete_plan_with_unchanged_inputs(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    run = _prepared_run(scenario)
    _stub_plan_provider(monkeypatch, {"summary": "核对奖金", "steps": ["核对E001奖金"], "questions": []})
    agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=False), user=scenario.user)
    monkeypatch.setattr(agent.ModelConfig, "from_env", lambda: None)

    result = agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=True), user=scenario.user)

    assert result["plan_confirmation"]["confirmed"] is True
    assert scenario.calls == []
    assert not result.get("draft_filename")
    assert {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")} == scenario.originals


def test_plan_confirmation_keeps_only_selected_steps_and_records_custom_instruction(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    run = _prepared_run(scenario)
    steps = ["核对E001奖金", "核对E002考勤"]
    _stub_plan_provider(monkeypatch, {"summary": "核对并更新副本", "steps": steps, "questions": []})
    agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=False), user=scenario.user)

    result = agent.confirm_agent_plan(
        run["run_id"],
        agent.AgentPlanIn(
            confirm_plan=True,
            selected_steps=[steps[0]],
            instruction="奖金仅处理有来源依据的人员",
        ),
        user=scenario.user,
    )

    assert result["model_plan"]["steps"] == [steps[0]]
    saved = agent._load_run(run["run_id"], scenario.user)
    assert "奖金仅处理有来源依据的人员" in saved["instruction"]
    assert saved["events"][-1]["payload"]["selected_steps"] == [steps[0]]
    assert scenario.calls == []


@pytest.mark.parametrize("action", ["confirm", "execute"])
@pytest.mark.parametrize("filename", ["总表.xlsx", "奖金来源.xlsx"])
def test_changed_input_blocks_plan_confirmation_or_execution(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch, action: str, filename: str,
) -> None:
    run = _prepared_run(scenario)
    requests = _stub_plan_provider(monkeypatch, {"summary": "核对奖金", "steps": ["核对E001奖金"], "questions": []})
    agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=False), user=scenario.user)
    if action == "execute":
        agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=True), user=scenario.user)
    requests.clear()
    changed = next(path for path in scenario.originals if path.name == filename)
    workbook = openpyxl.load_workbook(changed)
    workbook.active["C2"] = 3000
    workbook.save(changed)
    workbook.close()
    current_files = {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")}

    with pytest.raises(HTTPException) as raised:
        if action == "confirm":
            agent.confirm_agent_plan(run["run_id"], agent.AgentPlanIn(confirm_plan=True), user=scenario.user)
        else:
            agent.execute_agent_run(run["run_id"], user=scenario.user)

    assert raised.value.status_code == 409
    assert "文件" in str(raised.value.detail)
    assert scenario.calls == []
    assert requests == []
    assert {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")} == current_files


def _create_and_confirm_plan(scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch) -> str:
    _stub_plan_provider(monkeypatch, {
        "summary": "按来源更新独立副本，原件不动", "steps": ["核对E001身份，将来源C2奖金复制到总表C2"], "questions": [],
    })
    created = agent.create_agent_run(
        agent.AgentRunCreateIn(project_id="project-1"), user=scenario.user, db=scenario.db,
    )
    run_id = created["run_id"]
    agent.confirm_agent_plan(run_id, agent.AgentPlanIn(confirm_plan=False), user=scenario.user)
    confirmed = agent.confirm_agent_plan(run_id, agent.AgentPlanIn(confirm_plan=True), user=scenario.user)
    assert confirmed["plan_confirmation"]["confirmed"] is True
    assert not confirmed.get("draft_filename")
    return run_id


def test_agent_main_flow_uses_model_tools_to_write_only_draft_and_cannot_self_approve(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    run_id = _create_and_confirm_plan(scenario, monkeypatch)
    snapshots: list[dict[str, Any]] = []

    class ExecutionProvider:
        def __init__(self, _config: Any):
            pass

        def complete(self, **kwargs: Any) -> ModelResponse:
            snapshots.append(deepcopy(kwargs))
            if len(snapshots) == 1:
                return ModelResponse(tool_calls=[
                    ToolCall(call_id="copy", name="prepare_workbook_copy"),
                    ToolCall(call_id="target", name="read_range", arguments={"sheet": "工资核算", "range": "A2:C2"}),
                    ToolCall(call_id="source", name="read_source_range", arguments={
                        "filename": "奖金来源.xlsx", "sheet": "工资核算", "range": "A2:C2",
                    }),
                ])
            if len(snapshots) == 2:
                return ModelResponse(tool_calls=[ToolCall(call_id="write", name="apply_source_cells", arguments={
                    "changes": [{
                        "source_file": "奖金来源.xlsx", "source_sheet": "工资核算", "source_cells": ["C2"],
                        "target_sheet": "工资核算", "target_cell": "C2", "expected_value": 0, "aggregation": "copy",
                    }],
                })])
            return ModelResponse(content="已将E001奖金2000写入草稿，仍需验收。")

    monkeypatch.setattr("backend.agent_execution.OpenAICompatibleProvider", ExecutionProvider)

    result = agent.execute_agent_run(run_id, user=scenario.user, db=scenario.db)

    assert scenario.calls == []
    assert len(snapshots) == 3
    advertised = {schema["function"]["name"] for schema in snapshots[0]["tools"]}
    assert {
        "prepare_workbook_copy",
        "read_range",
        "read_source_range",
        "apply_source_cells",
        "apply_formula_divisors",
        "insert_and_copy_row",
    } <= advertised
    read_results = {
        message["tool_call_id"]: json.loads(message["content"])
        for message in snapshots[1]["messages"] if message["role"] == "tool"
    }
    assert read_results["target"]["values"] == [["E001", "测试人员", 0]]
    assert read_results["source"]["values"] == [["E001", "测试人员", 2000]]
    draft = agent._result_path("project-1", result["draft_filename"])
    assert draft not in scenario.originals
    workbook = openpyxl.load_workbook(draft, data_only=False)
    try:
        assert workbook["工资核算"]["C2"].value == 2000
    finally:
        workbook.close()
    assert {path: path.read_bytes() for path in scenario.originals} == scenario.originals
    assert result["workbook_updates"] == [{
        "target_sheet": "工资核算", "target_cell": "C2", "old_value": 0, "new_value": 2000,
        "source_file": "奖金来源.xlsx", "source_sheet": "工资核算", "source_cells": ["C2"], "aggregation": "copy",
    }]
    tool_results = [event["payload"] for event in result["events"] if event["type"] == "tool_result"]
    assert [(item["name"], item["status"]) for item in tool_results] == [
        ("prepare_workbook_copy", "succeeded"), ("read_range", "succeeded"),
        ("read_source_range", "succeeded"), ("apply_source_cells", "succeeded"),
    ]
    assert result["validation"]["status"] == "not_verified"
    saved = agent._load_run(run_id, scenario.user)
    assert saved["workbook_updates"] == result["workbook_updates"]
    with pytest.raises(HTTPException) as raised:
        agent.publish_agent_run(run_id, agent.AgentPublishIn(), user=scenario.user, db=scenario.db)
    assert raised.value.status_code == 409
    assert scenario.calls == []


def test_matching_keyuan_source_runs_basic_processor_before_model(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    scenario.db.files[1].original_name = "薪资数据-科园-7月薪资.xlsx"
    run_id = _create_and_confirm_plan(scenario, monkeypatch)
    processor_calls: list[tuple[Path, Path, Path]] = []

    def fake_processor(master: Path, source: Path, output: Path) -> dict[str, Any]:
        processor_calls.append((master, source, output))
        workbook = openpyxl.load_workbook(output)
        workbook.active["C2"] = 2000
        workbook.save(output)
        workbook.close()
        return {
            "status": "passed",
            "changes": [{"sheet": "工资核算", "cell": "C2", "before": 0, "after": 2000}],
            "issues": [],
        }

    monkeypatch.setattr("scripts.keyuan_basic_processor.process_keyuan_basics", fake_processor)
    monkeypatch.setattr(
        "backend.agent_execution.find_basic_salary_source",
        lambda _paths: ("薪资数据-科园-7月薪资.xlsx", scenario.root / "uploads" / "project-1" / "奖金来源.xlsx"),
    )

    class ProseProvider:
        def __init__(self, _config: Any):
            pass

        def complete(self, **_kwargs: Any) -> ModelResponse:
            return ModelResponse(content="基础处理已完成，仍需验收。")

    monkeypatch.setattr("backend.agent_execution.OpenAICompatibleProvider", ProseProvider)
    result = agent.execute_agent_run(run_id, user=scenario.user, db=scenario.db)

    assert len(processor_calls) == 1
    master, source, output = processor_calls[0]
    assert master.name == "总表.xlsx"
    assert source.name == "奖金来源.xlsx"
    assert output == agent._result_path("project-1", result["draft_filename"])
    assert result["basic_processor"] == {
        "status": "passed", "change_count": 1, "issue_count": 0, "issues": [],
    }
    workbook = openpyxl.load_workbook(output, data_only=False)
    try:
        assert workbook.active["C2"].value == 2000
    finally:
        workbook.close()


def test_basic_source_detection_accepts_generic_monthly_filename(tmp_path: Path) -> None:
    from backend.agent_execution import find_basic_salary_source

    source = tmp_path / "更新表-2026-08.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "奖金-8月"
    workbook.create_sheet("考勤-8月")
    workbook.save(source)
    workbook.close()

    detected = find_basic_salary_source({"更新表-2026-08.xlsx": str(source)})

    assert detected == ("更新表-2026-08.xlsx", source)


def test_basic_processor_runs_before_model_configuration_gate(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    scenario.db.files[1].original_name = "薪资数据-科园-8月薪资.xlsx"
    run_id = _create_and_confirm_plan(scenario, monkeypatch)
    calls: list[Path] = []

    def fake_processor(master: Path, source: Path, output: Path) -> dict[str, Any]:
        calls.append(output)
        return {"status": "passed", "changes": [], "issues": []}

    monkeypatch.setattr("scripts.keyuan_basic_processor.process_keyuan_basics", fake_processor)
    monkeypatch.setattr("backend.agent_execution.ModelConfig.from_env", lambda: None)
    monkeypatch.setattr("backend.agent_execution.ModelConfig.fallback_from_env", lambda: None)

    result = agent.execute_agent_run(run_id, user=scenario.user, db=scenario.db)

    assert calls
    assert result["status"] == "blocked"
    assert result["code"] == "MODEL_CONFIGURATION_REQUIRED"
    assert "基础 Python 更新已完成" in result["detail"]


def test_background_workflow_runs_preflight_before_model_planning(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    run = _prepared_run(scenario)
    run["status"] = "planning"
    calls: list[str] = []

    monkeypatch.setattr(agent, "_load_run", lambda *_args, **_kwargs: run)
    monkeypatch.setattr(agent, "_save_run", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(agent, "_material_context", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(agent, "_rule_package_context", lambda *_args, **_kwargs: {})
    monkeypatch.setattr(agent, "SessionLocal", lambda: SimpleNamespace(close=lambda: None))
    monkeypatch.setattr(agent, "_release_run_worker", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(agent.ModelConfig, "from_env", lambda: None)
    monkeypatch.setattr(agent.ModelConfig, "fallback_from_env", lambda: None)
    monkeypatch.setattr(
        agent,
        "find_basic_salary_source",
        lambda _paths: ("奖金来源.xlsx", Path(run["_source_paths"]["奖金来源.xlsx"])),
    )

    def fake_preflight(current_run: dict[str, Any], **_kwargs: Any) -> dict[str, Any]:
        calls.append("preflight")
        current_run["basic_processor"] = {
            "status": "passed", "change_count": 1, "issue_count": 0, "issues": [],
        }
        current_run["preflight_draft_filename"] = "Agent草稿.xlsx"
        return current_run["basic_processor"]

    monkeypatch.setattr(agent, "run_basic_preflight", fake_preflight)
    monkeypatch.setattr(
        agent,
        "confirm_agent_plan",
        lambda *_args, **_kwargs: calls.append("plan"),
    )

    agent._run_agent_workflow(run["run_id"], "tenant-a")

    assert calls == ["preflight"]
    assert run["status"] == "blocked"
    assert run["code"] == "MODEL_CONFIGURATION_REQUIRED"
    assert "基础 Python 更新已完成" in run["detail"]


def test_model_prose_without_write_tool_calls_never_creates_or_approves_a_draft(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    run_id = _create_and_confirm_plan(scenario, monkeypatch)

    class ProseOnlyProvider:
        def __init__(self, _config: Any):
            pass

        def complete(self, **_kwargs: Any) -> ModelResponse:
            return ModelResponse(content="已全部写入并通过所有验收，可以发布。")

    monkeypatch.setattr("backend.agent_execution.OpenAICompatibleProvider", ProseOnlyProvider)

    result = agent.execute_agent_run(run_id, user=scenario.user, db=scenario.db)

    assert not result.get("draft_filename")
    assert not result.get("workbook_updates")
    assert {path: path.read_bytes() for path in scenario.root.rglob("*.xlsx")} == scenario.originals
    assert result["validation"]["status"] == "not_verified"
    assert scenario.calls == []
    with pytest.raises(HTTPException) as raised:
        agent.publish_agent_run(run_id, agent.AgentPublishIn(), user=scenario.user, db=scenario.db)
    assert raised.value.status_code == 409


def test_empty_model_execution_is_resumable_with_an_actionable_detail(
    scenario: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
) -> None:
    run_id = _create_and_confirm_plan(scenario, monkeypatch)

    class EmptyProvider:
        def __init__(self, _config: Any):
            pass

        def complete(self, **_kwargs: Any) -> ModelResponse:
            return ModelResponse()

    monkeypatch.setattr("backend.agent_execution.OpenAICompatibleProvider", EmptyProvider)

    result = agent.execute_agent_run(run_id, user=scenario.user, db=scenario.db)

    assert result["status"] == "execution_incomplete"
    assert result["execution_result"]["code"] == "EMPTY_MODEL_RESPONSE"
    assert "自动重试" in result["detail"]


def test_public_run_recovers_legacy_empty_execution_as_resumable() -> None:
    result = agent._public_run({
        "run_id": "run-1",
        "status": "awaiting_review",
        "detail": "本轮未完成执行，请查看运行记录后重试；未发布正式结果",
        "items": [],
        "workbook_updates": [],
        "execution_result": {"status": "completed", "code": None, "content": ""},
        "validation": {"status": "not_verified"},
    })

    assert result["status"] == "execution_incomplete"
    assert result["execution_result"]["code"] == "EMPTY_MODEL_RESPONSE"
    assert "继续执行" in result["detail"]


def test_public_run_exposes_actionable_provider_error_detail() -> None:
    result = agent._public_run({
        "run_id": "run-1",
        "status": "execution_incomplete",
        "items": [],
        "workbook_updates": [],
        "execution_result": {
            "status": "execution_incomplete",
            "code": "MODEL_PROVIDER_ERROR",
            "content": "模型 API Key 或服务权限认证失败，请检查模型地址和密钥",
        },
    })

    assert "API Key" in result["detail"]
    assert "请继续执行" in result["detail"]
