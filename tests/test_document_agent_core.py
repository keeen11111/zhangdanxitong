from decimal import Decimal

from core.document_agent import (
    Decision,
    ExecutionPlan,
    WorkItem,
    choose_bonus_source,
)
from core.document_agent.contracts import ToolCall
from core.document_agent.model import ModelProviderError, ModelResponse
from core.document_agent.orchestrator import ModelOrchestrator, ToolRegistry
from core.document_agent.rules import Rule, RuleAction, RulePackage, RuleSource, resolve_rule_priority


def test_unique_nonzero_bonus_source_is_automatic() -> None:
    assert choose_bonus_source(2000, 0).model_dump() == {
        "status": "auto", "source": "J", "value": 2000,
        "reason": "exactly_one_nonzero",
    }
    assert choose_bonus_source(0, Decimal("2000")).source == "K"


def test_bonus_source_boundaries_require_review() -> None:
    assert choose_bonus_source(0, 0).status == "no_update"
    assert choose_bonus_source(2000, 3000).status == "needs_conversation"
    assert choose_bonus_source(-1, 0).status == "needs_conversation"
    assert choose_bonus_source("待确认", 0).status == "needs_conversation"


def test_execution_plan_is_person_scoped_and_decision_is_whitelisted() -> None:
    item = WorkItem(item_id="item-1", person_key="E001", person_category="店员", status="needs_review")
    plan = ExecutionPlan(
        run_id="run-1", company_id="tenant-a", salary_month="2026.07",
        rule_version="keyuan-1", items=[item], auto_publish=False,
    )
    assert plan.items[0].person_key == "E001"
    assert Decision(item_id="item-1", action="apply_proposed", value=2000).action == "apply_proposed"


def test_rule_package_is_company_scoped_and_uses_explicit_precedence() -> None:
    package = RulePackage(
        package_id="keyuan-payroll",
        company_id="tenant-a",
        version="2026.08.1",
        effective_period="2026-08",
        rules=[Rule(
            rule_id="bonus",
            condition={"category": "店员"},
            action=RuleAction(type="set_value_from_source", target_field="bonus", source_field="J"),
            source_refs=["手册.docx#p3"],
        )],
        sources=[
            RuleSource(kind="manual", ref="手册.docx#p3", excerpt="奖金从 J 列取值。"),
            RuleSource(kind="recording", ref="会议转写#00:10", excerpt="会议确认奖金取值规则。"),
        ],
    )
    assert package.rules[0].action.source_field == "J"
    assert resolve_rule_priority("explicit_instruction") == 0
    assert resolve_rule_priority("active_rule_package") == 1
    assert resolve_rule_priority("manual") == 2
    assert resolve_rule_priority("recording") == 3


def test_turn_limit_returns_an_incomplete_execution_result() -> None:
    class ReadOnlyProvider:
        def complete(self, *, messages, tools):
            return ModelResponse(content="仍在读取来源", tool_calls=[ToolCall(
                call_id=f"call-{len(messages)}", name="inspect_workbook",
            )])

    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda: {"sheets": []})
    result = ModelOrchestrator(provider=ReadOnlyProvider(), registry=registry, max_turns=1).run(
        run_id="run-1", messages=[{"role": "user", "content": "继续"}], tools=[],
    )

    assert result.status == "execution_incomplete"
    assert result.code == "MAX_TURNS_EXCEEDED"


def test_orchestrator_allows_a_project_to_finish_after_eight_tool_turns() -> None:
    class LongRunningProvider:
        def __init__(self) -> None:
            self.turns = 0

        def complete(self, *, messages, tools):
            self.turns += 1
            if self.turns <= 9:
                return ModelResponse(content="继续读取", tool_calls=[ToolCall(
                    call_id=f"call-{self.turns}", name="inspect_workbook", arguments={"attempt": self.turns},
                )])
            return ModelResponse(content="项目处理完成")

    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda **_: {"sheets": []})
    result = ModelOrchestrator(
        provider=LongRunningProvider(), registry=registry, max_turns=10,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    assert result.status == "completed"
    assert result.content == "项目处理完成"


def test_orchestrator_emits_each_event_before_the_next_model_turn() -> None:
    observed: list[tuple[str, str | None]] = []

    class StreamingProvider:
        def __init__(self) -> None:
            self.turns = 0

        def complete(self, *, messages, tools):
            self.turns += 1
            if self.turns == 1:
                return ModelResponse(content="正在读取", tool_calls=[ToolCall(
                    call_id="call-1", name="inspect_workbook",
                )])
            assert observed[-1] == ("model_request", None)
            return ModelResponse(content="已完成")

    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda: {"sheets": []})
    result = ModelOrchestrator(provider=StreamingProvider(), registry=registry).run(
        run_id="run-stream", messages=[{"role": "user", "content": "开始"}], tools=[],
        on_event=lambda event: observed.append((event.type, event.payload.get("name"))),
    )

    assert result.status == "completed"
    assert observed == [
        ("model_request", None),
        ("model_response", None),
        ("tool_call", "inspect_workbook"),
        ("tool_result", "inspect_workbook"),
        ("model_request", None),
        ("model_response", None),
    ]


def test_orchestrator_stops_repeated_tool_requests_before_repeating_a_write() -> None:
    class StuckProvider:
        def complete(self, *, messages, tools):
            return ModelResponse(content="继续", tool_calls=[ToolCall(
                call_id=f"call-{len(messages)}", name="apply_source_cells", arguments={"changes": []},
            )])

    registry = ToolRegistry()
    write_count = 0

    def apply_source_cells(changes):
        nonlocal write_count
        write_count += 1
        return {"updates": []}

    registry.register("apply_source_cells", apply_source_cells)
    result = ModelOrchestrator(provider=StuckProvider(), registry=registry).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "blocked"
    assert result.code == "NO_PROGRESS_DETECTED"
    assert write_count == 1


def test_orchestrator_reports_repeated_read_without_blocking_the_run() -> None:
    class AlternatingReadProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            calls = [
                ToolCall(call_id="read-a-1", name="inspect_workbook", arguments={"sheet": "工资"}),
                ToolCall(call_id="read-b", name="inspect_workbook", arguments={"sheet": "奖金"}),
                ToolCall(call_id="read-a-2", name="inspect_workbook", arguments={"sheet": "工资"}),
            ]
            return ModelResponse(content="继续读取", tool_calls=[calls[self.turn - 1]])

    reads: list[str] = []
    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda sheet: reads.append(sheet) or {"sheet": sheet})
    result = ModelOrchestrator(
        provider=AlternatingReadProvider(), registry=registry, max_turns=3,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    assert result.status == "execution_incomplete"
    assert result.code == "MAX_TURNS_EXCEEDED"
    assert reads == ["工资", "奖金"]
    assert any(
        event.type == "tool_result"
        and str(event.payload.get("error") or "").startswith("该读取范围已读取过")
        for event in result.events
    )


def test_orchestrator_keeps_model_context_bounded_for_long_projects() -> None:
    class LongProjectProvider:
        def __init__(self) -> None:
            self.turns = 0
            self.message_counts: list[int] = []

        def complete(self, *, messages, tools):
            self.turns += 1
            self.message_counts.append(len(messages))
            if self.turns <= 16:
                return ModelResponse(content="继续读取", tool_calls=[ToolCall(
                    call_id=f"call-{self.turns}", name="inspect_workbook", arguments={"attempt": self.turns},
                )])
            return ModelResponse(content="项目处理完成")

    provider = LongProjectProvider()
    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda **_: {"sheets": []})
    result = ModelOrchestrator(provider=provider, registry=registry).run(
        run_id="run-1", messages=[{"role": "system", "content": "规则"}, {"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "completed"
    assert max(provider.message_counts) <= 26


def test_orchestrator_retries_an_empty_model_response_and_continues() -> None:
    class EmptyThenCompleteProvider:
        def __init__(self) -> None:
            self.turns = 0

        def complete(self, *, messages, tools):
            self.turns += 1
            if self.turns == 1:
                return ModelResponse()
            return ModelResponse(content="项目处理完成")

    provider = EmptyThenCompleteProvider()
    result = ModelOrchestrator(provider=provider).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert provider.turns == 2
    assert result.status == "completed"
    assert result.content == "项目处理完成"


def test_orchestrator_retries_a_temporary_provider_error_and_continues() -> None:
    class UnavailableThenCompleteProvider:
        def __init__(self) -> None:
            self.turns = 0

        def complete(self, *, messages, tools):
            self.turns += 1
            if self.turns == 1:
                raise ModelProviderError("模型服务暂时不可用")
            return ModelResponse(content="项目处理完成")

    provider = UnavailableThenCompleteProvider()
    result = ModelOrchestrator(provider=provider).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert provider.turns == 2
    assert result.status == "completed"


def test_orchestrator_switches_to_fallback_provider_with_same_context() -> None:
    class PrimaryProvider:
        def complete(self, *, messages, tools):
            raise ModelProviderError("主模型不可用")

    class FallbackProvider:
        def __init__(self) -> None:
            self.messages = []

        def complete(self, *, messages, tools):
            self.messages.append(messages)
            return ModelResponse(content="备用模型已完成")

    fallback = FallbackProvider()
    result = ModelOrchestrator(
        provider=PrimaryProvider(), fallback_providers=[fallback],
    ).run(
        run_id="run-1", messages=[{"role": "user", "content": "继续处理"}], tools=[],
    )

    assert result.status == "completed"
    assert result.content == "备用模型已完成"
    assert fallback.messages[0][0]["content"] == "继续处理"
    assert any(event.payload.get("stage") == "model_failover" for event in result.events)


def test_orchestrator_marks_persistent_empty_responses_as_resumable() -> None:
    class EmptyProvider:
        def complete(self, *, messages, tools):
            return ModelResponse()

    result = ModelOrchestrator(provider=EmptyProvider()).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "execution_incomplete"
    assert result.code == "EMPTY_MODEL_RESPONSE"


def test_orchestrator_compacts_context_after_empty_response() -> None:
    requests = []

    class Provider:
        def complete(self, *, messages, tools):
            requests.append(messages)
            if len(requests) == 1:
                return ModelResponse()
            return ModelResponse(content="已从保存进度继续")

    result = ModelOrchestrator(provider=Provider()).run(
        run_id="run-1",
        messages=[{"role": "user", "content": "处理项目"}],
        tools=[],
    )

    assert result.status == "completed"
    assert len(requests) == 2
    assert len(requests[1]) == 2
    assert "当前已保存进度" in requests[1][-1]["content"]


def test_orchestrator_marks_persistent_provider_errors_as_resumable() -> None:
    class UnavailableProvider:
        def complete(self, *, messages, tools):
            raise ModelProviderError("模型服务暂时不可用")

    result = ModelOrchestrator(provider=UnavailableProvider()).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "execution_incomplete"
    assert result.code == "MODEL_PROVIDER_ERROR"


def test_orchestrator_compacts_context_after_http_400_and_retries_same_provider() -> None:
    requests = []

    class Provider:
        def complete(self, *, messages, tools):
            requests.append(messages)
            if len(requests) == 1:
                raise ModelProviderError("模型服务请求失败（HTTP 400）")
            return ModelResponse(content="已从保存进度继续")

    result = ModelOrchestrator(provider=Provider()).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[]
    )

    assert result.status == "completed"
    assert len(requests) == 2
    assert len(requests[1]) == 2
    assert "HTTP 400" in requests[1][-1]["content"]


def test_orchestrator_preserves_actionable_provider_error_details() -> None:
    class UnauthorizedProvider:
        def complete(self, *, messages, tools):
            raise ModelProviderError("模型 API Key 或服务权限认证失败，请检查模型地址和密钥")

    result = ModelOrchestrator(provider=UnauthorizedProvider()).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "execution_incomplete"
    assert result.code == "MODEL_PROVIDER_ERROR"
    assert "API Key" in result.content
    failure_events = [event for event in result.events if event.type == "needs_user_input"]
    assert failure_events[-1].payload["detail"] == result.content
