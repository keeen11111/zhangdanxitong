from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
from typing import Any

import openpyxl
import pytest

from backend.routers.agent import _build_run_tool_registry, _model_tool_schemas
from core.document_agent.contracts import ToolCall, ToolResult
from core.document_agent.model import ModelConfig, ModelProviderError, ModelResponse, OpenAICompatibleProvider
from core.document_agent.orchestrator import ModelOrchestrator, ToolRegistry


@pytest.mark.parametrize("name", ["prepare_workbook_copy", "apply_source_cells"])
def test_new_workbook_write_tool_names_are_validated_and_executable(name: str) -> None:
    registry = ToolRegistry()
    registry.register(name, lambda: {"executed": name})

    result = registry.execute(ToolCall(call_id="new-write-tool", name=name))

    assert result.name == name
    assert result.status == "succeeded"
    assert result.ok is True
    assert result.output == {"executed": name}
    assert ToolResult.model_validate_json(result.model_dump_json()) == result


@pytest.mark.parametrize(
    "schema", _model_tool_schemas(), ids=lambda schema: schema["function"]["name"],
)
def test_every_advertised_tool_survives_typed_registry_round_trip(
    schema: dict[str, Any],
) -> None:
    name = schema["function"]["name"]
    registry = ToolRegistry()
    registry.register(name, lambda **arguments: {"received": arguments})

    call = ToolCall(call_id=f"call-{name}", name=name, arguments={})
    result = registry.execute(call)

    assert result.status == "succeeded"
    assert result.ok is True
    assert result.call_id == call.call_id
    assert result.name == name
    assert result.output == {"received": {}}
    assert ToolResult.model_validate_json(result.model_dump_json()) == result


def test_officecli_validation_tool_is_exposed_without_command_arguments() -> None:
    schema = next(
        item for item in _model_tool_schemas()
        if item["function"]["name"] == "validate_with_officecli"
    )
    assert schema["function"]["parameters"] == {
        "type": "object",
        "properties": {},
        "required": [],
        "additionalProperties": False,
    }


def test_officecli_validation_is_scoped_to_the_current_draft(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft = tmp_path / "draft.xlsx"
    draft.write_bytes(b"placeholder")
    captured: dict[str, Any] = {}

    class Completed:
        returncode = 0
        stdout = "workbook is valid"
        stderr = ""

    def fake_run(args: list[str], **kwargs: Any) -> Completed:
        captured["args"] = args
        captured["kwargs"] = kwargs
        return Completed()

    monkeypatch.setattr("backend.routers.agent.subprocess.run", fake_run)
    registry = _build_run_tool_registry({
        "run_id": "run-1", "project_id": "project-1", "draft_filename": draft.name,
        "items": [],
    }, workbook_path=draft)

    result = registry.execute(ToolCall(
        call_id="officecli", name="validate_with_officecli", arguments={},
    ))

    assert result.ok is True
    assert result.output["valid"] is True
    assert Path(captured["args"][-1]) == draft
    assert captured["kwargs"]["shell"] is False


@pytest.mark.parametrize("name", ["inspect_source_file", "read_source_range"])
def test_advertised_source_tools_read_the_registered_run_source(
    name: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "tenant-a" / "奖金来源.xlsx"
    source.parent.mkdir()
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "奖金名单"
    sheet.append(["工号", "奖金"])
    sheet.append(["E001", 2000])
    workbook.save(source)
    workbook.close()
    monkeypatch.setattr("backend.routers.agent.UPLOAD_DIR", str(tmp_path))
    run = {
        "run_id": "run-1", "project_id": "project-1",
        "source_files": [source.name], "_source_paths": {source.name: str(source)},
    }
    registry = _build_run_tool_registry(run, workbook_path=tmp_path / "draft.xlsx")
    schema = next(item for item in _model_tool_schemas() if item["function"]["name"] == name)
    arguments = {"filename": source.name}
    if name == "read_source_range":
        arguments.update({"sheet": "奖金名单", "range": "A1:B2"})
    assert set(arguments) == set(schema["function"]["parameters"]["required"])

    result = registry.execute(ToolCall(call_id="source-call", name=name, arguments=arguments))

    assert result.ok is True, result.error
    assert result.status == "succeeded"
    assert result.output["filename"] == source.name
    if name == "inspect_source_file":
        assert result.output["sheets"] == [{
            "name": "奖金名单", "max_row": 2, "max_column": 2,
            "sample": [["工号", "奖金"], ["E001", 2000]],
        }]
    else:
        assert result.output["values"] == [["工号", "奖金"], ["E001", 2000]]
    assert result.observation == result.output


@pytest.mark.parametrize("name", ["inspect_source_file", "read_source_range"])
def test_responses_provider_accepts_advertised_source_function_calls(
    name: str, monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1",
        api_key="test-secret", model="finance-model", api_style="responses",
    ))
    arguments = {"filename": "奖金来源.xlsx"}
    if name == "read_source_range":
        arguments.update({"sheet": "奖金名单", "range": "A1:B2"})
    captured: dict[str, Any] = {}

    def fake_request(request: Any, **_kwargs: Any) -> _Response:
        captured["payload"] = json.loads(request.data)
        return _Response({
            "id": "response-source", "status": "completed",
            "output": [{
                "type": "function_call", "call_id": "source-call", "name": name,
                "arguments": json.dumps(arguments, ensure_ascii=False),
            }],
        })

    monkeypatch.setattr("urllib.request.urlopen", fake_request)
    response = provider.complete(
        messages=[{"role": "user", "content": "核对本次运行的奖金来源"}],
        tools=_model_tool_schemas(),
    )

    assert name in {tool["name"] for tool in captured["payload"]["tools"]}
    assert [call.model_dump() for call in response.tool_calls] == [{
        "call_id": "source-call", "name": name, "arguments": arguments,
    }]


def test_multi_tool_response_remains_one_assistant_turn_before_all_results() -> None:
    calls = [
        ToolCall(call_id="call-inspect", name="inspect_workbook"),
        ToolCall(call_id="call-read", name="read_range", arguments={"sheet": "奖金名单", "range": "A1:B2"}),
    ]
    snapshots: list[list[dict[str, Any]]] = []

    class Provider:
        def complete(self, *, messages: Any, tools: Any) -> ModelResponse:
            snapshots.append(deepcopy(messages))
            if len(snapshots) == 1:
                return ModelResponse(content="同时检查结构和奖金来源。", tool_calls=calls)
            return ModelResponse(content="已核对来源。")

    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda: {"sheets": ["奖金名单"]})
    registry.register("read_range", lambda **_arguments: {"values": [["E001", 2000]]})
    original_messages = [{"role": "user", "content": "核对奖金"}]

    result = ModelOrchestrator(provider=Provider(), registry=registry).run(
        run_id="run-1", messages=original_messages, tools=_model_tool_schemas(),
    )

    assert result.status == "completed"
    assert original_messages == [{"role": "user", "content": "核对奖金"}]
    follow_up = snapshots[1]
    assert [message["role"] for message in follow_up] == ["user", "assistant", "tool", "tool"]
    assert follow_up[1]["content"] == "同时检查结构和奖金来源。"
    assert [call["id"] for call in follow_up[1]["tool_calls"]] == ["call-inspect", "call-read"]
    assert [message["tool_call_id"] for message in follow_up[2:]] == ["call-inspect", "call-read"]
    assert [json.loads(message["content"]) for message in follow_up[2:]] == [
        {"sheets": ["奖金名单"]}, {"values": [["E001", 2000]]},
    ]
    assert [event.type for event in result.events] == [
        "model_request", "model_response", "tool_call", "tool_result", "tool_call", "tool_result",
        "model_request", "model_response",
    ]


def test_malformed_tool_arguments_are_retried_with_same_provider() -> None:
    """A transient provider formatting error must not abort the workbook run."""
    calls: list[list[dict[str, Any]]] = []

    class Provider:
        def complete(self, *, messages: Any, tools: Any) -> ModelResponse:
            calls.append(deepcopy(messages))
            if len(calls) == 1:
                raise ModelProviderError("模型工具调用参数无效")
            return ModelResponse(content="已继续完成核对。")

    result = ModelOrchestrator(
        provider=Provider(),
        registry=ToolRegistry(),
        fallback_providers=[object()],
    ).run(
        run_id="run-tool-retry",
        messages=[{"role": "user", "content": "继续处理"}],
        tools=_model_tool_schemas(),
    )

    assert result.status == "completed"
    assert result.content == "已继续完成核对。"
    assert len(calls) == 2
    assert calls[1][-1]["role"] == "user"
    assert "不是合法 JSON" in calls[1][-1]["content"]
    assert [event.payload.get("stage") for event in result.events if event.type == "model_request"] == [
        "model_call", "model_call",
    ]


class _Response:
    def __init__(self, payload: dict[str, Any]):
        self._body = json.dumps(payload).encode("utf-8")

    def read(self) -> bytes:
        return self._body

    def __enter__(self) -> _Response:
        return self

    def __exit__(self, *_args: object) -> None:
        return None
