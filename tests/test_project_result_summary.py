"""项目列表的结果状态摘要测试。"""

import json

from backend.routers import projects as projects_router


def test_project_result_summary_uses_completed_export_metadata(tmp_path, monkeypatch) -> None:
    """历史项目有可用结果时，应带回结果时间和仍待处理的人工项数。"""
    monkeypatch.setattr(projects_router, "SESSION_DIR", str(tmp_path))
    project_id = "completed-project"
    (tmp_path / f"{project_id}_export_meta.json").write_text(
        json.dumps(
            {
                "status": "completed_with_issues",
                "filename": "工资核算_2026.06.xlsx",
                "completed_at": "2026-08-25T14:36:00",
                "issue_count": 12,
            }
        ),
        encoding="utf-8",
    )

    summary = projects_router._project_result_summary(project_id)

    assert summary == {
        "has_result": True,
        "result_completed_at": "2026-08-25T14:36:00",
        "pending_issue_count": 12,
    }


def test_project_result_summary_ignores_processing_or_broken_metadata(tmp_path, monkeypatch) -> None:
    """处理中或缺少结果文件的元数据不能把项目错误地导向结果页。"""
    monkeypatch.setattr(projects_router, "SESSION_DIR", str(tmp_path))
    project_id = "processing-project"
    (tmp_path / f"{project_id}_export_meta.json").write_text(
        json.dumps({"status": "processing", "filename": "临时文件.xlsx", "issue_count": 9}),
        encoding="utf-8",
    )

    assert projects_router._project_result_summary(project_id) == {
        "has_result": False,
        "result_completed_at": None,
        "pending_issue_count": 0,
    }


def test_project_result_summary_includes_a_generic_financial_draft(tmp_path, monkeypatch) -> None:
    """通用财务草稿也必须在项目列表显示为可继续处理的结果。"""
    monkeypatch.setattr(projects_router, "SESSION_DIR", str(tmp_path / "sessions"))
    monkeypatch.setattr(projects_router, "DATA_DIR", str(tmp_path))
    project_id = "financial-project"
    metadata_dir = tmp_path / "financial-workbook-integrations" / project_id
    metadata_dir.mkdir(parents=True)
    (metadata_dir / "latest.json").write_text(
        json.dumps(
            {
                "status": "review_required",
                "filename": "应收账款草稿.xlsx",
                "completed_at": "2026-08-25T16:20:00",
                "issues": [{}, {}, {}],
            }
        ),
        encoding="utf-8",
    )

    assert projects_router._project_result_summary(project_id) == {
        "has_result": True,
        "result_completed_at": "2026-08-25T16:20:00",
        "pending_issue_count": 3,
    }
