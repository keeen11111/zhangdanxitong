import pytest
from pydantic import ValidationError

from core.document_agent.rule_compiler import (
    RuleCandidate,
    RuleCompilationError,
    compile_rule_package,
    validate_rule_package,
)
from core.document_agent.rules import Rule, RuleAction, RulePackage, RuleSource


def _candidate(
    *,
    rule_id: str = "bonus-source",
    source_kind: str = "manual",
    source_ref: str = "manual-v1.docx#p12",
    excerpt: str = "奖金从 J、K 两列中唯一非零列取值。",
    revision: int = 1,
    action_type: str = "choose_unique_nonzero",
) -> RuleCandidate:
    action: dict[str, object] = {
        "type": action_type,
        "target_field": "bonus",
    }
    if action_type == "choose_unique_nonzero":
        action["source_fields"] = ["J", "K"]
    if action_type == "set_value_from_source":
        action["source_field"] = "J"
    return RuleCandidate.model_validate(
        {
            "rule_id": rule_id,
            "condition": {"person_category": "店员"},
            "action": action,
            "source": {
                "kind": source_kind,
                "ref": source_ref,
                "excerpt": excerpt,
                "revision": revision,
            },
        }
    )


def _compile(*candidates: RuleCandidate) -> RulePackage:
    return compile_rule_package(
        package_id="keyuan-payroll",
        company_id="tenant-a",
        version="2026.08.1",
        effective_period="2026-08",
        candidates=candidates,
    )


def test_compiler_applies_required_source_precedence() -> None:
    recording = _candidate(
        source_kind="recording",
        source_ref="meeting.m4a#00:18:20",
        excerpt="当时说优先选 K 列。",
        action_type="set_value_from_source",
    )
    manual = _candidate()
    active_package = _candidate(
        source_kind="active_rule_package",
        source_ref="rules-2026.07.json#/rules/bonus-source",
        excerpt="历史生效规则指定 K 列。",
        action_type="set_value_from_source",
    )
    explicit = _candidate(
        source_kind="explicit_instruction",
        source_ref="run-42#message-7",
        excerpt="本次明确按 J、K 唯一非零选择。",
    )

    package = _compile(recording, manual, active_package, explicit)

    assert package.rules[0].action.type == "choose_unique_nonzero"
    assert package.rules[0].source_refs == ["run-42#message-7"]
    assert [source.ref for source in package.sources] == ["run-42#message-7"]


def test_compiler_uses_latest_manual_revision() -> None:
    old_manual = _candidate(action_type="set_value_from_source")
    latest_manual = _candidate(
        source_ref="manual-v2.docx#p9",
        excerpt="新版手册改为 J、K 唯一非零。",
        revision=2,
    )

    package = _compile(old_manual, latest_manual)

    assert package.rules[0].source_refs == ["manual-v2.docx#p9"]


def test_identical_same_priority_rules_merge_evidence() -> None:
    first = _candidate(source_ref="manual.docx#p9")
    second = _candidate(source_ref="manual.docx#table3")

    package = _compile(second, first)

    assert package.rules[0].source_refs == ["manual.docx#p9", "manual.docx#table3"]
    assert [source.ref for source in package.sources] == [
        "manual.docx#p9",
        "manual.docx#table3",
    ]


def test_conflicting_same_priority_rules_are_rejected() -> None:
    choose_nonzero = _candidate(source_ref="manual-a.docx#p4")
    choose_j = _candidate(
        source_ref="manual-b.docx#p7",
        excerpt="奖金始终从 J 列取值。",
        action_type="set_value_from_source",
    )

    with pytest.raises(RuleCompilationError, match="同优先级规则冲突"):
        _compile(choose_nonzero, choose_j)


@pytest.mark.parametrize(
    "action",
    [
        {"type": "run_python", "code": "open('payroll.xlsx', 'wb')"},
        {"type": "set_formula", "formula": "=SUM(J2:K2)"},
        {"type": "set_value_from_source", "source_field": "=J2+K2"},
    ],
)
def test_actions_reject_code_formula_and_non_whitelisted_types(action: dict[str, str]) -> None:
    with pytest.raises(ValidationError):
        RuleAction.model_validate(action)


def test_candidate_requires_locatable_evidence() -> None:
    with pytest.raises(ValidationError):
        RuleCandidate.model_validate(
            {
                "rule_id": "bonus-source",
                "condition": {},
                "action": {"type": "keep_current"},
                "source": {"kind": "manual", "ref": "manual.docx#p3", "excerpt": ""},
            }
        )


def test_package_validation_rejects_missing_or_unknown_evidence() -> None:
    with pytest.raises(ValidationError):
        Rule(
            rule_id="bonus-source",
            action=RuleAction(type="keep_current"),
            source_refs=[],
        )

    invalid = RulePackage.model_construct(
        package_id="keyuan-payroll",
        company_id="tenant-a",
        version="2026.08.1",
        effective_period="2026-08",
        status="candidate",
        rules=[
            Rule(
                rule_id="bonus-source",
                action=RuleAction(type="keep_current"),
                source_refs=["missing#ref"],
            )
        ],
        sources=[
            RuleSource(kind="manual", ref="manual.docx#p3", excerpt="保留当前值。")
        ],
    )
    with pytest.raises(ValueError, match="未知来源"):
        validate_rule_package(invalid)


def test_condition_rejects_executable_payloads() -> None:
    with pytest.raises(ValidationError, match="不允许可执行字段"):
        RuleCandidate.model_validate(
            {
                "rule_id": "unsafe",
                "condition": {"python": "import os"},
                "action": {"type": "skip"},
                "source": {
                    "kind": "explicit_instruction",
                    "ref": "run-42#message-8",
                    "excerpt": "跳过该人员。",
                },
            }
        )
