from __future__ import annotations

from pathlib import Path


def test_project_skill_context_reports_installed_sources_and_bounded_policy(tmp_path: Path) -> None:
    """The runtime receives a compact policy, never entire third-party manuals."""
    from backend.project_agent_skills import render_project_skill_context

    skill_root = tmp_path / "agent_skills"
    for name in (
        "officecli-xlsx", "spreadsheets", "document-xlsx",
        "cli-anything-openrefine", "docling", "call-agent",
    ):
        location = skill_root / name
        location.mkdir(parents=True)
        (location / "SKILL.md").write_text("# external source\n", encoding="utf-8")

    context = render_project_skill_context(skill_root)

    assert context["installed_skills"] == [
        "officecli-xlsx",
        "spreadsheets",
        "document-xlsx",
        "cli-anything-openrefine",
        "docling",
        "call-agent",
    ]
    assert "先概览再读必要范围" in context["execution_policy"]
    assert "原子" in context["execution_policy"]
    assert "空响应" in context["execution_policy"]
    assert len(context["execution_policy"]) < 4000
    routes = {route["id"]: route for route in context["task_routing"]}
    assert routes["workbook-update"]["skill_names"] == [
        "officecli-xlsx", "spreadsheets", "document-xlsx",
    ]
    assert "草稿" in routes["workbook-update"]["workflow"]
    assert routes["workbook-update"]["required_validation_tools"] == [
        "validate_workbook", "validate_with_officecli",
    ]
    assert "来源值" in routes["data-cleaning"]["workflow"]
    assert routes["document-evidence"]["runtime_mode"] == "reference_only"


def test_project_skill_context_omits_missing_skill_directories(tmp_path: Path) -> None:
    from backend.project_agent_skills import render_project_skill_context

    skill_root = tmp_path / "agent_skills"
    location = skill_root / "spreadsheets"
    location.mkdir(parents=True)
    (location / "SKILL.md").write_text("# external source\n", encoding="utf-8")

    context = render_project_skill_context(skill_root)

    assert context["installed_skills"] == ["spreadsheets"]
    assert context["task_routing"] == []


def test_select_project_skills_marks_runtime_gates_and_reference_only_routes() -> None:
    from backend.project_agent_skills import select_project_skills

    selected = select_project_skills({
        "instruction": "核对姓名不一致并完成本月工资结算",
        "file_manifest": [{"filename": "source.xlsx"}, {"filename": "contract.pdf"}],
        "items": [{"issue_type": "unmatched_person", "status": "needs_review"}],
    })

    ids = {route["id"] for route in selected["routes"]}
    assert "workbook-update" in ids
    assert "data-cleaning" in ids
    assert "month-end-close" in ids
    assert "document-evidence" in ids
    assert selected["required_validation_tools"] == ["validate_with_officecli", "validate_workbook"]
    assert "workbook-update" in selected["runtime_routes"]
    assert "document-evidence" in selected["reference_only_routes"]
