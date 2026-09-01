"""Integration progress persistence and API contract tests."""

from types import SimpleNamespace

import backend.routers.pipeline as pipeline_module


def test_integration_progress_persists_completed_stages(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(pipeline_module, "SESSION_DIR", str(tmp_path))
    stages = [
        {"key": "base", "label": "加载原始数据库", "status": "completed"},
        {"key": "source", "label": "解析来源文件", "status": "running"},
    ]

    saved = pipeline_module._save_integration_progress(
        "project-1",
        status="processing",
        stages=stages,
        detail="正在解析来源文件",
    )

    loaded = pipeline_module._load_integration_progress("project-1")

    assert saved["status"] == "processing"
    assert loaded["stages"] == stages
    assert loaded["detail"] == "正在解析来源文件"
    assert loaded["updated_at"]


def test_integration_progress_defaults_to_idle_when_no_run_exists(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(pipeline_module, "SESSION_DIR", str(tmp_path))

    progress = pipeline_module._load_integration_progress("project-1")

    assert progress["status"] == "idle"
    assert [stage["status"] for stage in progress["stages"]] == ["pending"] * 4


def test_progress_endpoint_loads_only_the_authorized_project(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(pipeline_module, "SESSION_DIR", str(tmp_path))
    monkeypatch.setattr(
        pipeline_module,
        "_load_project_or_404",
        lambda project_id, user, db: SimpleNamespace(id=project_id),
    )
    pipeline_module._save_integration_progress(
        "project-1",
        status="completed",
        stages=[{"key": "base", "label": "加载原始数据库", "status": "completed"}],
    )

    progress = pipeline_module.get_integration_progress("project-1", object(), object())

    assert progress.status == "completed"
    assert progress.stages[0]["key"] == "base"
