from decimal import Decimal
import json

import pytest

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


def test_repeated_structure_inspection_forces_the_model_toward_a_write() -> None:
    """A repeated inspection gets guidance instead of ending a write task."""

    class RepeatingInspectionProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn <= 3:
                return ModelResponse(content="继续检查结构", tool_calls=[ToolCall(
                    call_id=f"inspect-{self.turn}", name="inspect_workbook",
                )])
            if self.turn == 4:
                assert any(
                    message.get("role") == "tool" and "重复" in str(message.get("content"))
                    for message in messages
                )
                return ModelResponse(content="开始写入", tool_calls=[ToolCall(
                    call_id="write-1", name="apply_source_cells",
                    arguments={"changes": [{"sheet": "工资", "cell": "A1", "value": "已核对"}]},
                )])
            return ModelResponse(content="已完成写入")

    registry = ToolRegistry()
    inspections: list[bool] = []
    registry.register("inspect_workbook", lambda: inspections.append(True) or {"sheets": []})
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    result = ModelOrchestrator(
        provider=RepeatingInspectionProvider(), registry=registry, max_turns=6,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "继续"}], tools=[])

    assert result.status == "completed"
    assert inspections == [True, True]
    assert any(
        event.type == "tool_result" and "重复" in str(event.payload.get("error") or "")
        for event in result.events
    )


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


def test_orchestrator_read_call_limit_forces_write_phase_after_cap() -> None:
    class ReadForeverProvider:
        """Ignores guidance and keeps requesting data reads long past the cap."""

        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            return ModelResponse(content="再读一次", tool_calls=[ToolCall(
                call_id=f"read-{self.turn}", name="read_range",
                arguments={"sheet": f"表{self.turn}"},
            )])

    executed_reads: list[str] = []

    def read_range(sheet):
        executed_reads.append(sheet)
        return {"cells": [[sheet]]}

    registry = ToolRegistry()
    registry.register("read_range", read_range)
    result = ModelOrchestrator(
        provider=ReadForeverProvider(), registry=registry, max_turns=20, read_call_limit=3,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    # 前 3 次数据读取真实执行，之后的读取全部被拒绝且不再触达工具本体。
    assert executed_reads == ["表1", "表2", "表3"]
    assert any(
        event.type == "tool_result"
        and "读取阶段已结束" in str(event.payload.get("error") or "")
        for event in result.events
    )


def test_writing_stage_rejects_new_data_reads_and_keeps_durable_state() -> None:
    class ReadAfterWriteProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="read-after-write", name="read_range",
                    arguments={"sheet": "明细", "range": "A1:C3"},
                )])
            return ModelResponse(content="已停止重读")

    executed_reads: list[bool] = []
    registry = ToolRegistry()
    registry.register("read_range", lambda **_: executed_reads.append(True) or {})
    result = ModelOrchestrator(
        provider=ReadAfterWriteProvider(), registry=registry, stage="writing",
        initial_context_state={
            "current_stage": "writing",
            "completed_writes": [{"tool": "apply_source_cells", "count": 1}],
            "last_read_checkpoint": "cp-read",
            "last_write_checkpoint": "cp-write",
        },
    ).run(run_id="run-writing", messages=[{"role": "user", "content": "继续"}])

    assert result.status == "completed"
    assert executed_reads == []
    assert result.context_state["last_read_checkpoint"] == "cp-read"
    assert result.context_state["last_write_checkpoint"] == "cp-write"
    rejection = next(
        event for event in result.events
        if event.type == "tool_result" and event.payload.get("code") == "READ_FORBIDDEN_AFTER_WRITING"
    )
    assert rejection.payload["status"] == "failed"


def test_orchestrator_read_cap_ignores_structure_inspection() -> None:
    """inspect/find_table 等结构探查不占读取额度：额度只留给数据读取。"""

    class InspectForeverProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn > 5:
                return ModelResponse(content="结构核对完成")
            return ModelResponse(content="查看结构", tool_calls=[ToolCall(
                call_id=f"inspect-{self.turn}", name="inspect_workbook",
                arguments={"attempt": self.turn},
            )])

    executed = []

    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda **kwargs: executed.append(kwargs) or {"sheets": []})
    result = ModelOrchestrator(
        provider=InspectForeverProvider(), registry=registry, max_turns=10, read_call_limit=2,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    assert result.status == "completed"
    assert len(executed) == 5
    assert not any(
        event.type == "tool_result"
        and "读取阶段已结束" in str(event.payload.get("error") or "")
        for event in result.events
    )


def test_orchestrator_keeps_read_results_beyond_replay_window() -> None:
    """回放窗口外的数据读取结果通过存档保留：写入阶段仍然可见。"""

    class ReadThenWriteProvider:
        def __init__(self) -> None:
            self.turn = 0
            self.write_request_messages = None

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn <= 8:
                return ModelResponse(content="读取", tool_calls=[ToolCall(
                    call_id=f"read-{self.turn}", name="read_range",
                    arguments={"sheet": f"数据{self.turn}"},
                )])
            if self.turn == 9:
                self.write_request_messages = messages
                return ModelResponse(content="写入", tool_calls=[ToolCall(
                    call_id="write-1", name="apply_source_cells",
                    arguments={"changes": [{"sheet": "明细", "cell": "A1", "value": 1}]},
                )])
            return ModelResponse(content="已完成写入")

    def read_range(sheet):
        return {"cells": [[f"{sheet}-行1"], [f"{sheet}-行2"]]}

    registry = ToolRegistry()
    registry.register("read_range", read_range)
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    provider = ReadThenWriteProvider()
    result = ModelOrchestrator(provider=provider, registry=registry, max_turns=12).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "completed"
    assert provider.write_request_messages is not None
    # 第 1 轮的读取早已滑出 6 轮回放窗口，但其数据必须仍在写入请求的上下文中。
    flat = [str(message.get("content")) for message in provider.write_request_messages]
    assert any("数据1-行1" in content for content in flat)
    # 窗口内的读取不重复出现（存档只接管窗口外的条目）。
    assert sum("数据8-行1" in content for content in flat) == 1


def test_orchestrator_empty_response_retry_preserves_read_archive() -> None:
    """空响应后的紧凑重试只丢弃轮次回放，已读取数据必须保留。"""

    class ReadEmptyWriteProvider:
        def __init__(self) -> None:
            self.turn = 0
            self.retry_request_messages = None

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(content="读取", tool_calls=[ToolCall(
                    call_id="read-1", name="read_range", arguments={"sheet": "派遣"},
                )])
            if self.turn == 2:
                return ModelResponse()  # 空响应，触发紧凑重试
            if self.turn == 3:
                self.retry_request_messages = messages
                return ModelResponse(content="基于存档写入", tool_calls=[ToolCall(
                    call_id="write-1", name="apply_source_cells",
                    arguments={"changes": [{"sheet": "明细", "cell": "A1", "value": 1}]},
                )])
            return ModelResponse(content="已使用存档数据完成写入")

    registry = ToolRegistry()
    registry.register("read_range", lambda sheet: {"cells": [[f"{sheet}-人员A"], [f"{sheet}-人员B"]]})
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    provider = ReadEmptyWriteProvider()
    result = ModelOrchestrator(provider=provider, registry=registry, max_turns=6).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "completed"
    assert provider.retry_request_messages is not None
    flat = [str(message.get("content")) for message in provider.retry_request_messages]
    assert any("派遣-人员A" in content for content in flat)


def test_prepare_workbook_copy_does_not_count_as_data_write() -> None:
    """建草稿副本不是数据写入：零写入不能被纯文字总结成 completed。"""

    class CopyThenTalkProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(content="先建草稿", tool_calls=[ToolCall(
                    call_id="copy-1", name="prepare_workbook_copy",
                )])
            return ModelResponse(content="我认为已经处理完成了")

    registry = ToolRegistry()
    registry.register("prepare_workbook_copy", lambda: {"status": "created"})
    provider = CopyThenTalkProvider()
    result = ModelOrchestrator(
        provider=provider, registry=registry, max_turns=6, require_writes=True,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    # 连续 2 次催促后仍未写入，应保持可续跑而不是把 0 写入判定为完成。
    nudge_events = [
        event for event in result.events
        if event.payload.get("stage") == "nudge_no_tool_call"
    ]
    assert len(nudge_events) == 2
    assert provider.turn == 4  # 1 次建草稿 + 2 次被催促的纯文字 + 1 次最终接受
    assert result.status == "execution_incomplete"
    assert result.code == "NO_WRITES_PERFORMED"


@pytest.mark.parametrize(
    ("tool_name", "output"),
    [
        ("roll_forward_month", {"rolled": True}),
        ("run_basic_payroll_processor", {"status": "skipped", "reason": "模板不匹配"}),
    ],
)
def test_structural_or_skipped_tools_cannot_satisfy_required_business_writes(
    tool_name: str, output: dict[str, object],
) -> None:
    class Provider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="non-business-write", name=tool_name,
                    arguments={"change": {}} if tool_name == "roll_forward_month" else {},
                )])
            return ModelResponse(content="处理完成")

    registry = ToolRegistry()
    registry.register(tool_name, lambda **_arguments: output)
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, max_turns=5, require_writes=True,
    ).run(run_id="no-business-write", messages=[{"role": "user", "content": "更新数据"}])

    assert result.status == "execution_incomplete"
    assert result.code == "NO_WRITES_PERFORMED"


def test_required_skill_validations_must_execute_before_completion() -> None:
    class WriteValidateProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            calls = {
                1: ToolCall(call_id="write", name="apply_source_cells", arguments={"changes": [{}]}),
                3: ToolCall(call_id="structural", name="validate_workbook"),
                4: ToolCall(call_id="recalc", name="validate_with_officecli"),
            }
            if self.turn in calls:
                return ModelResponse(tool_calls=[calls[self.turn]])
            return ModelResponse(content="完成")

    registry = ToolRegistry()
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    registry.register("validate_workbook", lambda: {
        "readable": True, "unresolved_item_count": 0,
        "formula_errors": [], "summary_range_errors": [],
        "duplicate_identities": [], "identity_errors": [], "can_publish": True,
    })
    registry.register("validate_with_officecli", lambda: {
        "valid": True, "recalculated": True,
        "formula_errors": [], "unevaluated_formulas": [],
    })
    provider = WriteValidateProvider()
    result = ModelOrchestrator(
        provider=provider, registry=registry, require_writes=True,
        required_validation_tools=["validate_workbook", "validate_with_officecli"],
        max_turns=6,
    ).run(run_id="run-validation-gate", messages=[{"role": "user", "content": "处理"}])

    assert result.status == "completed"
    assert provider.turn == 5
    assert set(result.context_state["completed_steps"]) == {
        "validate_workbook", "validate_with_officecli",
    }
    assert any(event.payload.get("stage") == "nudge_validation" for event in result.events)


def test_schema_only_officecli_validation_cannot_satisfy_recalculation_gate() -> None:
    class Provider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            calls = {
                1: ToolCall(call_id="write", name="apply_source_cells", arguments={"changes": [{}]}),
                2: ToolCall(call_id="structural", name="validate_workbook"),
                3: ToolCall(call_id="schema-only", name="validate_with_officecli"),
            }
            if self.turn in calls:
                return ModelResponse(tool_calls=[calls[self.turn]])
            return ModelResponse(content="完成")

    registry = ToolRegistry()
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    registry.register("validate_workbook", lambda: {
        "readable": True, "unresolved_item_count": 0,
        "formula_errors": [], "summary_range_errors": [],
        "duplicate_identities": [], "identity_errors": [], "can_publish": True,
    })
    registry.register("validate_with_officecli", lambda: {"valid": True})
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, require_writes=True,
        required_validation_tools=["validate_workbook", "validate_with_officecli"],
        max_turns=5,
    ).run(run_id="schema-only-validation", messages=[{"role": "user", "content": "处理"}])

    assert result.status == "execution_incomplete"
    assert result.code == "VALIDATION_FAILED"
    assert "validate_with_officecli" not in result.context_state["completed_steps"]


def test_summary_range_errors_fail_structural_validation_gate() -> None:
    class Provider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="write", name="apply_source_cells", arguments={"changes": [{}]},
                )])
            if self.turn == 2:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="validate", name="validate_workbook",
                )])
            return ModelResponse(content="完成")

    registry = ToolRegistry()
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    registry.register("validate_workbook", lambda: {
        "readable": True, "unresolved_item_count": 0,
        "formula_errors": [], "summary_range_errors": [{"cell": "C4"}],
        "duplicate_identities": [], "identity_errors": [], "can_publish": True,
    })
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, require_writes=True,
        required_validation_tools=["validate_workbook"], max_turns=4,
    ).run(run_id="summary-range-invalid", messages=[{"role": "user", "content": "处理"}])

    assert result.status == "execution_incomplete"
    assert result.code == "VALIDATION_FAILED"
    assert "validate_workbook" not in result.context_state["completed_steps"]


def test_a_write_after_validation_requires_both_validations_to_run_again() -> None:
    class Provider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            calls = {
                1: ToolCall(call_id="write-1", name="apply_source_cells", arguments={"changes": [{}]}),
                2: ToolCall(call_id="structural-1", name="validate_workbook"),
                3: ToolCall(call_id="recalc-1", name="validate_with_officecli"),
                4: ToolCall(call_id="write-2", name="apply_source_cells", arguments={"changes": [{}]}),
                6: ToolCall(call_id="structural-2", name="validate_workbook"),
                7: ToolCall(call_id="recalc-2", name="validate_with_officecli"),
            }
            if self.turn in calls:
                return ModelResponse(tool_calls=[calls[self.turn]])
            return ModelResponse(content="完成")

    registry = ToolRegistry()
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    registry.register("validate_workbook", lambda: {
        "readable": True, "unresolved_item_count": 0,
        "formula_errors": [], "summary_range_errors": [],
        "duplicate_identities": [], "identity_errors": [], "can_publish": True,
    })
    registry.register("validate_with_officecli", lambda: {
        "valid": True, "recalculated": True,
        "formula_errors": [], "unevaluated_formulas": [],
    })
    provider = Provider()
    result = ModelOrchestrator(
        provider=provider, registry=registry, require_writes=True,
        required_validation_tools=["validate_workbook", "validate_with_officecli"],
        max_turns=9,
    ).run(run_id="post-validation-write", messages=[{"role": "user", "content": "处理"}])

    assert result.status == "completed"
    assert provider.turn == 8
    successful_validations = [
        event.payload["name"] for event in result.events
        if event.type == "tool_result"
        and event.payload.get("status") == "succeeded"
        and event.payload.get("name") in {"validate_workbook", "validate_with_officecli"}
    ]
    assert successful_validations == [
        "validate_workbook", "validate_with_officecli",
        "validate_workbook", "validate_with_officecli",
    ]


def test_failed_machine_validation_cannot_be_summarized_as_complete() -> None:
    class InvalidProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="write", name="apply_source_cells", arguments={"changes": [{}]},
                )])
            if self.turn == 2:
                return ModelResponse(tool_calls=[ToolCall(call_id="validate", name="validate_workbook")])
            return ModelResponse(content="已经完成")

    registry = ToolRegistry()
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    registry.register("validate_workbook", lambda: {
        "readable": True, "unresolved_item_count": 0,
        "formula_errors": [{"cell": "A1", "error": "#REF!"}],
        "duplicate_identities": [], "can_publish": False,
    })
    result = ModelOrchestrator(
        provider=InvalidProvider(), registry=registry, require_writes=True,
        required_validation_tools=["validate_workbook"], max_turns=4,
    ).run(run_id="run-validation-failed", messages=[{"role": "user", "content": "处理"}])

    assert result.status == "execution_incomplete"
    assert result.code == "VALIDATION_FAILED"
    assert result.context_state["validation_failures"]


def test_incomplete_post_write_summary_is_continued_until_remaining_write() -> None:
    """A preliminary row deletion must not end a multi-step workbook run."""

    class DeleteThenWriteProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn == 1:
                return ModelResponse(content="删除多余人员", tool_calls=[ToolCall(
                    call_id="delete-1", name="delete_rows", arguments={"change": {"sheet": "明细"}},
                )])
            if self.turn == 2:
                return ModelResponse(content="The rows were deleted. Now I need to update wage fields.")
            if self.turn == 3:
                return ModelResponse(content="写入工资字段", tool_calls=[ToolCall(
                    call_id="write-1", name="apply_source_cells", arguments={"changes": [{"cell": "A1"}]},
                )])
            return ModelResponse(content="已完成全部写入和校验")

    registry = ToolRegistry()
    registry.register("delete_rows", lambda change: {"deleted": True})
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    result = ModelOrchestrator(
        provider=DeleteThenWriteProvider(), registry=registry, max_turns=6, require_writes=True,
    ).run(run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    assert result.status == "completed"
    assert any(event.payload.get("stage") == "nudge_incomplete_summary" for event in result.events)


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
                ToolCall(call_id="read-a-1", name="read_range", arguments={"sheet": "工资", "range": "A1:C3"}),
                ToolCall(call_id="read-b", name="read_range", arguments={"sheet": "奖金", "range": "A1:C3"}),
                ToolCall(call_id="read-a-2", name="read_range", arguments={"sheet": "工资", "range": "A1:C3"}),
            ]
            return ModelResponse(content="继续读取", tool_calls=[calls[self.turn - 1]])

    reads: list[str] = []
    registry = ToolRegistry()
    registry.register("read_range", lambda sheet, range: reads.append(sheet) or {"sheet": sheet, "range": range})
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


def test_orchestrator_allows_repeated_structural_inspection() -> None:
    """结构探查（inspect/find_table）隔轮重复不拦截：

    空响应紧凑重试后模型需要先确认草稿/结构才能继续写入；
    拦截会造成“想看状态被拒→空响应”死循环（实发案例：rev 55，
    turn 2 的 inspect 在 turn 7 重复时被误拦）。
    """

    class MixedInspectProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            plan = [
                ("inspect-1", "inspect_workbook", {}),
                ("read-1", "read_range", {"sheet": "工资", "range": "A1:C3"}),
                ("inspect-2", "inspect_workbook", {}),
                ("read-2", "read_range", {"sheet": "奖金", "range": "A1:C3"}),
            ]
            if self.turn > len(plan):
                return ModelResponse(content="处理完成")
            call_id, name, arguments = plan[self.turn - 1]
            return ModelResponse(content="继续", tool_calls=[ToolCall(call_id=call_id, name=name, arguments=arguments)])

    executed: list[str] = []
    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda: executed.append("inspect") or {"sheets": []})
    registry.register("read_range", lambda sheet, range: executed.append("read") or {"sheet": sheet, "range": range})
    result = ModelOrchestrator(
        provider=MixedInspectProvider(), registry=registry, max_turns=5,
    ).run(run_id="run-inspect", messages=[{"role": "user", "content": "处理项目"}], tools=[])

    # 两次 inspect 都真实执行（隔轮重复不拦截），读取正常执行
    assert executed == ["inspect", "read", "inspect", "read"]
    assert not any(
        event.type == "tool_result" and "已读取过" in str(event.payload.get("error") or "")
        for event in result.events
    )


def test_orchestrator_never_deduplicates_structural_inspection() -> None:
    """相同结构探查可连续执行；它是恢复状态确认，不是大数据重读。"""

    class InspectTwiceProvider:
        def __init__(self) -> None:
            self.turn = 0

        def complete(self, *, messages, tools):
            self.turn += 1
            if self.turn <= 2:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id=f"inspect-{self.turn}", name="inspect_workbook", arguments={},
                )])
            return ModelResponse(content="结构确认完成")

    executed: list[str] = []
    registry = ToolRegistry()
    registry.register("inspect_workbook", lambda: executed.append("inspect") or {"sheets": []})

    result = ModelOrchestrator(provider=InspectTwiceProvider(), registry=registry, max_turns=3).run(
        run_id="run-inspect-twice", messages=[{"role": "user", "content": "继续"}], tools=[],
    )

    assert result.status == "completed"
    assert executed == ["inspect", "inspect"]


def test_orchestrator_keeps_model_context_bounded_for_long_projects() -> None:
    class LongProjectProvider:
        def __init__(self) -> None:
            self.turns = 0
            self.message_counts: list[int] = []

        def complete(self, *, messages, tools):
            self.turns += 1
            self.message_counts.append(len(messages))
            if self.turns <= 12:
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


def test_orchestrator_enforces_total_context_char_budget() -> None:
    """会话总字符预算：多轮大读取结果不得把请求撑爆模型上下文窗口。

    实发案例（run 955a）：一轮 5 个 40K 读取 + 6 轮回放 + 100K 存档，
    请求超 128K token 窗口后 DeepSeek 返回空 content，run 卡死 0 写入。
    """

    class BigReadThenWriteProvider:
        def __init__(self) -> None:
            self.turns = 0
            self.total_chars: list[int] = []

        def complete(self, *, messages, tools):
            self.turns += 1
            self.total_chars.append(sum(
                len(str(m.get("content") or ""))
                + sum(len(str((tc.get("function") or {}).get("arguments") or "")) for tc in m.get("tool_calls") or [])
                for m in messages
            ))
            if self.turns <= 8:
                return ModelResponse(content="继续读取", tool_calls=[ToolCall(
                    call_id=f"read-{self.turns}", name="read_range",
                    arguments={"sheet": f"表{self.turns}", "range": f"A{self.turns}:C{self.turns + 10}"},
                )])
            return ModelResponse(content="数据已读完，开始写入")

    big_payload = {"rows": [["数值" + str(i)] * 8 for i in range(600)]}  # 约 30K 字符
    registry = ToolRegistry()
    registry.register("read_range", lambda sheet, range: big_payload)
    provider = BigReadThenWriteProvider()
    result = ModelOrchestrator(provider=provider, registry=registry, max_turns=10).run(
        run_id="run-budget", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "completed"
    # 总字符（含 tool_call arguments）始终受预算约束（允许少量超额来自最新轮无条件保留）
    from core.document_agent.orchestrator import CONVERSATION_TOTAL_CHAR_BUDGET
    assert max(provider.total_chars) <= CONVERSATION_TOTAL_CHAR_BUDGET + 50_000


def test_orchestrator_strictly_bounds_five_large_reads_from_one_turn() -> None:
    """单轮多个 40K 读取也不能穿透总预算（run 955a 的直接复现）。"""

    class FiveReadsProvider:
        def __init__(self) -> None:
            self.calls = 0
            self.request_chars: list[int] = []

        def complete(self, *, messages, tools):
            self.calls += 1
            self.request_chars.append(sum(len(str(message.get("content") or "")) for message in messages))
            if self.calls == 1:
                return ModelResponse(tool_calls=[
                    ToolCall(
                        call_id=f"read-{index}", name="read_range",
                        arguments={"sheet": f"表{index}", "range": "A1:Z1000"},
                    )
                    for index in range(5)
                ])
            return ModelResponse(content="已读取并安全结束")

    registry = ToolRegistry()
    registry.register("read_range", lambda sheet, range: {"data": "数" * 40_000})
    provider = FiveReadsProvider()

    result = ModelOrchestrator(provider=provider, registry=registry, max_turns=2).run(
        run_id="run-five-reads", messages=[{"role": "user", "content": "处理大表"}], tools=[],
    )

    from core.document_agent.orchestrator import CONVERSATION_TOTAL_CHAR_BUDGET
    assert result.status == "completed"
    assert max(provider.request_chars) <= CONVERSATION_TOTAL_CHAR_BUDGET


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


def test_failover_after_empty_response_uses_compact_continuation() -> None:
    """Provider failover must not replay the stale read transcript."""

    class PrimaryProvider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            self.calls += 1
            if self.calls == 1:
                return ModelResponse(content="读取", tool_calls=[ToolCall(
                    call_id="read-1", name="read_range",
                    arguments={"sheet": "明细", "range": "A1:C3"},
                )])
            return ModelResponse()

    class FallbackProvider:
        def __init__(self) -> None:
            self.messages = []

        def complete(self, *, messages, tools):
            self.messages.append(messages)
            return ModelResponse(content="备用模型已继续")

    fallback = FallbackProvider()
    registry = ToolRegistry()
    registry.register("read_range", lambda **_: {
        "sheet": "明细", "range": "A1:C3", "values": [["工号", "姓名", "工资"], ["E1", "甲", 1]],
        "padding": "STALE-RAW-TRANSCRIPT" + ("x" * 1000),
    })
    result = ModelOrchestrator(
        provider=PrimaryProvider(), fallback_providers=[fallback], max_turns=3,
    ).run(run_id="run-empty-failover", messages=[{"role": "user", "content": "处理"}])

    assert result.status == "completed"
    assert fallback.messages
    serialized = json.dumps(fallback.messages[0], ensure_ascii=False)
    assert "compact_context" in serialized
    assert "STALE-RAW-TRANSCRIPT" not in serialized


def test_orchestrator_marks_persistent_empty_responses_as_resumable() -> None:
    class EmptyProvider:
        def complete(self, *, messages, tools):
            return ModelResponse()

    result = ModelOrchestrator(provider=EmptyProvider()).run(
        run_id="run-1", messages=[{"role": "user", "content": "处理项目"}], tools=[],
    )

    assert result.status == "execution_incomplete"
    assert result.code == "EMPTY_MODEL_RESPONSE"


def test_run_8c92aaa_large_tool_result_is_replaced_by_checkpoint(monkeypatch) -> None:
    """8c92... repeatedly replayed large range reads until responses became empty."""

    monkeypatch.setenv("AGENT_INPUT_TOKEN_BUDGET", "12000")

    class Provider:
        def __init__(self) -> None:
            self.turn = 0
            self.requests = []

        def complete(self, *, messages, tools):
            self.turn += 1
            self.requests.append(json.dumps(messages, ensure_ascii=False))
            if self.turn == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="read-8c92", name="read_source_range",
                    arguments={"filename": "source.xlsx", "sheet": "派遣", "range": "A1:T24"},
                )])
            if self.turn == 2:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="write-8c92", name="apply_source_cells",
                    arguments={"changes": [{"target_sheet": "明细", "target_cell": "A3", "expected_value": "旧"}]},
                )])
            return ModelResponse(content="done")

    marker = "RAW-SALARY-CELL-DO-NOT-REPLAY"
    registry = ToolRegistry()
    registry.register("read_source_range", lambda **_: {
        "filename": "source.xlsx", "sheet": "派遣", "range": "A1:T24",
        "values": [["工号", "姓名", "基本工资"], ["E001", "张三", 100]],
        "padding": marker + ("x" * 30000),
    })
    registry.register("apply_source_cells", lambda changes: {"updates": changes})
    provider = Provider()

    result = ModelOrchestrator(
        provider=provider, registry=registry, max_turns=4, require_writes=True,
    ).run(run_id="8c92aaa6f1e54aef84762aaeba5a13b1", messages=[{"role": "user", "content": "只处理派遣"}])

    assert result.status == "completed"
    assert marker not in provider.requests[1]
    assert marker not in provider.requests[2]
    assert "checkpoint_id" in provider.requests[2]
    # The continuation carries semantic checkpoint evidence only; the raw
    # range ``values`` payload must never be replayed after the read.
    assert '"values"' not in provider.requests[1]
    assert '"values"' not in provider.requests[2]


def test_run_cd9e97_length_compacts_instead_of_replaying_request(monkeypatch) -> None:
    """cd9e... had many empty turns after accumulated tool transcripts."""

    monkeypatch.setenv("AGENT_INPUT_TOKEN_BUDGET", "12000")

    class Provider:
        def __init__(self) -> None:
            self.requests = []

        def complete(self, *, messages, tools):
            self.requests.append(json.dumps(messages, ensure_ascii=False))
            if len(self.requests) == 1:
                return ModelResponse(content="analysis" * 1000, finish_reason="length", output_tokens=512)
            return ModelResponse(content="stopped after compact")

    provider = Provider()
    events = []
    result = ModelOrchestrator(provider=provider, max_turns=2).run(
        run_id="cd9e97cabed445d198912da3ae76251d",
        messages=[{"role": "user", "content": "处理派遣"}], on_event=events.append,
    )

    assert result.status == "completed"
    assert provider.requests[0] != provider.requests[1]
    assert "compact_context" in provider.requests[1]
    responses = [event for event in events if event.type == "model_response"]
    assert responses[0].payload["finish_reason"] == "length"
    assert responses[0].payload["reached_output_limit"] is True


def test_run_57ad_empty_response_stops_after_one_compaction_and_keeps_checkpoint() -> None:
    """57ad... must not turn repeated empty responses into another read loop."""

    class Provider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            self.calls += 1
            if self.calls == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="read-57ad", name="read_range", arguments={"sheet": "明细", "range": "A1:C3"},
                )])
            return ModelResponse()

    registry = ToolRegistry()
    registry.register("read_range", lambda **_: {"sheet": "明细", "range": "A1:C3", "values": [["工号", "姓名", "工资"], ["E1", "甲", 1]]})
    events = []
    result = ModelOrchestrator(provider=Provider(), registry=registry, max_turns=8).run(
        run_id="57ad71eb2cfc41e28b9fc32a63863b64",
        messages=[{"role": "user", "content": "处理"}], on_event=events.append,
    )

    assert result.status == "execution_incomplete"
    assert result.code == "EMPTY_MODEL_RESPONSE"
    assert sum(event.type == "model_response" for event in events) == 3
    checkpoint_events = [event for event in events if event.payload.get("checkpoint_id")]
    assert checkpoint_events


def test_run_460878_budget_exceeded_fails_closed_before_provider(monkeypatch) -> None:
    """460878... used a huge initial task and then repeatedly replanned the workbook."""

    monkeypatch.setenv("AGENT_INPUT_TOKEN_BUDGET", "40")

    class Provider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            self.calls += 1
            return ModelResponse(content="should not run")

    provider = Provider()
    result = ModelOrchestrator(provider=provider).run(
        run_id="46087850599045b4aaecf717125a719e",
        messages=[{"role": "system", "content": "s" * 2000}, {"role": "user", "content": "u" * 2000}],
        tools=[{"type": "function", "function": {"name": "read_range", "description": "d" * 2000, "parameters": {}}}],
    )

    assert provider.calls == 0
    assert result.status == "execution_incomplete"
    assert result.code == "CONTEXT_BUDGET_EXCEEDED"
    request = next(event for event in result.events if event.type == "model_request")
    assert request.payload["estimated_input_tokens"] > request.payload["input_token_budget"]


def test_model_request_and_response_share_local_correlation_id() -> None:
    class Provider:
        def complete(self, *, messages, tools):
            return ModelResponse(content="done", request_id="provider-response-123")

    result = ModelOrchestrator(provider=Provider(), stage="writing").run(
        run_id="run-correlation", messages=[{"role": "user", "content": "write"}], tools=[],
    )

    request = next(event for event in result.events if event.type == "model_request")
    response = next(event for event in result.events if event.type == "model_response")
    assert request.payload["request_id"]
    assert response.payload["request_id"] == request.payload["request_id"]
    assert response.payload["provider_request_id"] == "provider-response-123"
    assert request.payload["stage"] == "writing"


def test_checkpoint_preserves_arbitrary_required_write_evidence() -> None:
    from core.document_agent.orchestrator import _checkpoint_summary

    rows = [[f"H{column}" for column in range(1, 11)]]
    rows.extend([[f"R{row}C{column}" for column in range(1, 11)] for row in range(2, 41)])
    checkpoint = _checkpoint_summary(
        "read_source_range",
        {"filename": "source.xlsx", "sheet": "派遣", "range": "A1:J40"},
        {"filename": "source.xlsx", "sheet": "派遣", "range": "A1:J40", "rows": rows},
        sequence=1,
        requirements=[{
            "source_file": "source.xlsx", "source_sheet": "派遣", "source_cells": ["J30"],
            "target_sheet": "明细", "target_cell": "Q30", "expected_value": 123,
        }],
    )

    evidence = checkpoint["write_evidence"][0]
    assert evidence["source_coordinates"] == [{"cell": "J30", "value": "R30C10"}]
    assert evidence["target_coordinate"] == "Q30"
    assert evidence["expected_value"] == 123
    assert checkpoint["required_fields_missing"] == []


def test_resumed_segment_restores_read_checkpoint_for_a_pending_write() -> None:
    from core.document_agent.orchestrator import _checkpoint_summary

    change = {
        "source_file": "source.xlsx", "source_sheet": "派遣", "source_cells": ["C2"],
        "target_sheet": "明细", "target_cell": "D2", "expected_value": 0,
        "aggregation": "copy",
    }
    checkpoint = _checkpoint_summary(
        "read_source_range",
        {"filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2"},
        {
            "filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2",
            "rows": [["工号", "姓名", "金额"], ["E001", "李楠", 1200]],
        },
        sequence=1,
        requirements=[change],
    )

    class Provider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            self.calls += 1
            if self.calls == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="write-resumed", name="apply_source_cells",
                    arguments={"changes": [change]},
                )])
            return ModelResponse(content="写入和检查已完成")

    registry = ToolRegistry()
    registry.register("apply_source_cells", lambda changes: changes)
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, max_turns=2, require_writes=True,
        stage="writing",
        initial_context_state={
            "current_stage": "writing",
            "pending_writes": [change],
            "last_read_checkpoint": checkpoint["checkpoint_id"],
        },
        initial_checkpoints=[checkpoint],
    ).run(
        run_id="resumed-run", messages=[{"role": "user", "content": "从检查点继续写入"}],
    )

    assert result.status == "completed"
    assert result.context_state["completed_writes"] == [{"tool": "apply_source_cells", "count": 1}]


def test_large_read_payload_is_replaced_by_semantic_checkpoint_on_next_turn() -> None:
    sentinel = "LAST-SOURCE-VALUE-DO-NOT-DROP"
    rows = [
        [f"value-{row}-{column}-xxxxxxxx" for column in range(24)]
        for row in range(25)
    ]
    rows[-1][-1] = sentinel
    saw_sentinel = False

    class Provider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            nonlocal saw_sentinel
            self.calls += 1
            if self.calls == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="read-large", name="read_source_range",
                    arguments={"filename": "source.xlsx", "sheet": "数据", "range": "A1:X25"},
                )])
            saw_sentinel = sentinel in json.dumps(messages, ensure_ascii=False)
            return ModelResponse(content="读取完成")

    registry = ToolRegistry()
    registry.register("read_source_range", lambda **_arguments: {
        "filename": "source.xlsx", "sheet": "数据", "range": "A1:X25", "values": rows,
    })
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, max_turns=2,
    ).run(run_id="large-read", messages=[{"role": "user", "content": "读取来源"}])

    assert result.status == "completed"
    # Raw range payloads are intentionally kept server-side.  The next turn
    # receives bounded checkpoint evidence instead of replaying the complete
    # table (including the sentinel cell).
    assert saw_sentinel is False


def test_resumed_segment_restores_bounded_read_values_not_only_their_hash() -> None:
    sentinel = "RESUME-SOURCE-VALUE"
    rows = [["工号", "姓名", "金额"], ["E001", "李楠", sentinel]]

    class ReadProvider:
        def complete(self, *, messages, tools):
            return ModelResponse(tool_calls=[ToolCall(
                call_id="read-before-interrupt", name="read_source_range",
                arguments={"filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2"},
            )])

    first_registry = ToolRegistry()
    first_registry.register("read_source_range", lambda **_arguments: {
        "filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2", "values": rows,
    })
    interrupted = ModelOrchestrator(
        provider=ReadProvider(), registry=first_registry, max_turns=1,
    ).run(run_id="read-interrupted", messages=[{"role": "user", "content": "读取后继续"}])
    resumed_messages = []

    class ResumeProvider:
        def complete(self, *, messages, tools):
            resumed_messages.extend(messages)
            return ModelResponse(content="已恢复")

    resumed = ModelOrchestrator(
        provider=ResumeProvider(), max_turns=1,
        initial_context_state=interrupted.context_state,
        initial_checkpoints=interrupted.checkpoints,
    ).run(run_id="read-resumed", messages=[{"role": "user", "content": "继续"}])

    assert interrupted.code == "MAX_TURNS_EXCEEDED"
    assert resumed.status == "completed"
    assert sentinel in json.dumps(resumed_messages, ensure_ascii=False)


def test_month_roll_forward_does_not_discard_source_values_needed_for_writes() -> None:
    sentinel = 1200
    source_visible_after_roll = False
    change = {
        "source_file": "source.xlsx", "source_sheet": "派遣", "source_cells": ["C2"],
        "target_sheet": "明细", "target_cell": "D2", "expected_value": 0,
        "aggregation": "copy",
    }

    class Provider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            nonlocal source_visible_after_roll
            self.calls += 1
            if self.calls == 1:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="read-source", name="read_source_range",
                    arguments={"filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2"},
                )])
            if self.calls == 2:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="roll-month", name="roll_forward_month", arguments={"change": {}},
                )])
            if self.calls == 3:
                source_visible_after_roll = str(sentinel) in json.dumps(messages, ensure_ascii=False)
                return ModelResponse(tool_calls=[ToolCall(
                    call_id="write-source", name="apply_source_cells",
                    arguments={"changes": [change]},
                )])
            return ModelResponse(content="写入完成")

    registry = ToolRegistry()
    registry.register("read_source_range", lambda **_arguments: {
        "filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2",
        "values": [["工号", "姓名", "金额"], ["E001", "李楠", sentinel]],
    })
    registry.register("roll_forward_month", lambda change: {"rolled": True})
    registry.register("apply_source_cells", lambda changes: changes)
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, max_turns=4, require_writes=True,
    ).run(run_id="month-write", messages=[{"role": "user", "content": "新增月份并写入"}])

    assert result.status == "completed"
    assert source_visible_after_roll is True
    assert [entry["tool"] for entry in result.context_state["completed_writes"]] == [
        "roll_forward_month", "apply_source_cells",
    ]


def test_resumed_segment_never_reexecutes_an_archived_read_range() -> None:
    from core.document_agent.orchestrator import _checkpoint_summary

    arguments = {"filename": "source.xlsx", "sheet": "派遣", "range": "A1:C2"}
    fingerprint = json.dumps(
        {"name": "read_source_range", "arguments": arguments},
        ensure_ascii=False, sort_keys=True,
    )
    checkpoint = _checkpoint_summary(
        "read_source_range", arguments,
        {**arguments, "values": [["工号", "姓名", "金额"], ["E001", "李楠", 1200]]},
        sequence=1,
    )
    checkpoint.update({
        "call_id": "previous-read", "read_fingerprint": fingerprint,
        "archive_content": json.dumps({
            **checkpoint, "values": [["工号", "姓名", "金额"], ["E001", "李楠", 1200]],
        }, ensure_ascii=False),
    })

    class Provider:
        def __init__(self) -> None:
            self.calls = 0

        def complete(self, *, messages, tools):
            self.calls += 1
            if self.calls <= 3:
                return ModelResponse(tool_calls=[ToolCall(
                    call_id=f"duplicate-{self.calls}", name="read_source_range", arguments=arguments,
                )])
            return ModelResponse(content="停止重复读取")

    executions = 0

    def read_source_range(**_arguments):
        nonlocal executions
        executions += 1
        return {**arguments, "values": []}

    registry = ToolRegistry()
    registry.register("read_source_range", read_source_range)
    result = ModelOrchestrator(
        provider=Provider(), registry=registry, max_turns=4,
        initial_checkpoints=[checkpoint],
        initial_context_state={"current_stage": "reading"},
    ).run(run_id="resume-repeat", messages=[{"role": "user", "content": "继续"}])

    assert result.status == "completed"
    assert executions == 0
    duplicate_results = [
        event for event in result.events
        if event.type == "tool_result" and event.payload.get("name") == "read_source_range"
    ]
    assert len(duplicate_results) == 3
    assert all(event.payload.get("status") == "failed" for event in duplicate_results)


def test_legacy_checkpoint_archive_is_sanitized_before_prompt_replay() -> None:
    from core.document_agent.orchestrator import _semantic_checkpoint_content

    content = json.dumps({
        "checkpoint_id": "cp-1",
        "sheet": "明细",
        "range": "A1:C2",
        "current_cells": [{"row": 0, "column": 0, "value": "E001"}],
        "values": [["E001", "敏感工资", 999999]],
    }, ensure_ascii=False)

    sanitized = json.loads(_semantic_checkpoint_content(content))
    assert sanitized["checkpoint_id"] == "cp-1"
    assert sanitized["current_cells"][0]["value"] == "E001"
    assert "values" not in sanitized


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


def test_compact_context_fits_token_budget_after_large_initial_context(monkeypatch) -> None:
    monkeypatch.setenv("AGENT_INPUT_TOKEN_BUDGET", "1200")
    requests = []

    class Provider:
        def complete(self, *, messages, tools):
            requests.append(messages)
            if "compact_context" not in json.dumps(messages, ensure_ascii=False):
                return ModelResponse()
            return ModelResponse(content="已从压缩状态继续")

    result = ModelOrchestrator(provider=Provider(), max_turns=2).run(
        run_id="run-compact-token-fit",
        messages=[
            {"role": "system", "content": "系统规则" + "工资" * 2500},
            {"role": "user", "content": "用户要求" + "人员" * 2500},
        ],
        tools=[{"type": "function", "function": {
            "name": "read_range", "description": "读取", "parameters": {"type": "object"},
        }}],
    )

    assert result.status == "completed"
    from core.document_agent.model import estimate_token_count
    compact_tokens = estimate_token_count(json.dumps(requests[-1], ensure_ascii=False))
    tool_tokens = estimate_token_count(json.dumps([{"type": "function", "function": {
        "name": "read_range", "description": "读取", "parameters": {"type": "object"},
    }}], ensure_ascii=False))
    assert compact_tokens + tool_tokens <= 1200
    assert "工资" * 1000 not in json.dumps(requests[-1], ensure_ascii=False)


def test_compact_context_drops_server_side_archive_payload(monkeypatch) -> None:
    monkeypatch.setenv("AGENT_INPUT_TOKEN_BUDGET", "1200")
    requests = []
    checkpoint = {
        "checkpoint_id": "cp-legacy",
        "version": 1,
        "tool": "read_range",
        "sheet": "明细",
        "range": "A1:C2",
        "current_cells": [{"row": 0, "column": 0, "value": "E001"}],
        "data_hash": "hash",
        "archive_content": "RAW-ARCHIVE-" + ("x" * 100_000),
    }

    class Provider:
        def complete(self, *, messages, tools):
            requests.append(messages)
            if len(requests) == 1:
                return ModelResponse()
            return ModelResponse(content="已从 checkpoint 继续")

    result = ModelOrchestrator(
        provider=Provider(), max_turns=2, initial_checkpoints=[checkpoint],
    ).run(run_id="compact-archive", messages=[{"role": "user", "content": "继续"}], tools=[])

    assert result.status == "completed"
    compacted = json.dumps(requests[-1], ensure_ascii=False)
    assert "RAW-ARCHIVE-" not in compacted
    from core.document_agent.model import estimate_token_count
    assert estimate_token_count(compacted) <= 1200


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
