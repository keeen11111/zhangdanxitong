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
