from __future__ import annotations

import hashlib
import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from fastapi import HTTPException
import openpyxl
import pytest

from backend.agent_demo import demo_directory, finish_demo, load_demo


def test_keyuan_demo_script_is_not_the_agent_execution_path() -> None:
    """The old desktop automation script remains only as a legacy utility."""
    import backend.routers.agent as agent

    assert not hasattr(agent, "_is_keyuan_deterministic_run")
    assert not hasattr(agent, "_execute_keyuan_deterministic_run")


@pytest.fixture
def configured_demo(tmp_path: Path) -> dict[str, Any]:
    root = tmp_path / "demo-references"
    directory = demo_directory(root, "tenant-a", "project-1")
    directory.mkdir(parents=True, exist_ok=True)
    reference = directory / "result.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "客户已完成结果"
    workbook.active["A1"] = "预置成品演示，不代表Agent计算或验收"
    workbook.active["C2"] = 2000
    workbook.save(reference)
    workbook.close()
    original = reference.read_bytes()
    digest = hashlib.sha256(original).hexdigest()
    (directory / "manifest.json").write_text(json.dumps({
        "filename": "客户已完成参考.xlsx", "sha256": digest,
    }, ensure_ascii=False), encoding="utf-8")
    return {"root": root, "reference": reference, "original": original, "sha256": digest}


def _confirmed_run() -> dict[str, Any]:
    return {
        "run_id": "a" * 32, "tenant_id": "tenant-a", "project_id": "project-1",
        "plan_confirmation": {"required": True, "confirmed": True},
    }


@pytest.mark.parametrize("tenant_id,project_id", [("tenant-b", "project-1"), ("tenant-a", "project-2")])
def test_demo_reference_is_not_visible_to_another_tenant_or_project(
    configured_demo: dict[str, Any], tenant_id: str, project_id: str,
) -> None:
    root = configured_demo["root"]
    own = demo_directory(root, "tenant-a", "project-1")
    other = demo_directory(root, tenant_id, project_id)

    assert own != other
    assert own.resolve().is_relative_to(root.resolve())
    assert other.resolve().is_relative_to(root.resolve())
    assert load_demo(root, tenant_id, project_id) is None


def test_missing_demo_configuration_returns_none(tmp_path: Path) -> None:
    assert load_demo(tmp_path / "not-configured", "tenant-a", "project-1") is None


def test_load_demo_returns_verified_reference_metadata(configured_demo: dict[str, Any]) -> None:
    config = load_demo(configured_demo["root"], "tenant-a", "project-1")

    assert config is not None
    assert config["filename"] == "客户已完成参考.xlsx"
    assert config["sha256"] == configured_demo["sha256"]
    assert Path(config["_reference_path"]).resolve() == configured_demo["reference"].resolve()
    assert configured_demo["reference"].read_bytes() == configured_demo["original"]


def test_load_demo_rejects_reference_that_does_not_match_manifest(configured_demo: dict[str, Any]) -> None:
    configured_demo["reference"].write_bytes(b"changed after manifest")

    with pytest.raises(HTTPException) as raised:
        load_demo(configured_demo["root"], "tenant-a", "project-1")

    assert raised.value.status_code == 409


def test_finish_demo_copies_exact_bytes_and_marks_demo_not_real_agent_acceptance(
    configured_demo: dict[str, Any], tmp_path: Path,
) -> None:
    config = load_demo(configured_demo["root"], "tenant-a", "project-1")
    run = _confirmed_run()
    destination = tmp_path / "演示输出.xlsx"

    finish_demo(run, config, destination)

    assert destination.read_bytes() == configured_demo["original"]
    assert configured_demo["reference"].read_bytes() == configured_demo["original"]
    assert run["execution_mode"] == "demo"
    assert run["status"] == "completed"
    assert run["draft_filename"] == destination.name
    assert run["demo_result"]["filename"] == config["filename"]
    assert run["demo_result"]["sha256"] == hashlib.sha256(destination.read_bytes()).hexdigest()
    assert run["demo_result"]["provenance"] == "user_supplied_completed_workbook"
    assert run["validation"]["status"] == "demo_reference_match"
    assert run["validation"]["status"] != "passed"


def test_finish_demo_rejects_reference_as_output(configured_demo: dict[str, Any]) -> None:
    config = load_demo(configured_demo["root"], "tenant-a", "project-1")

    with pytest.raises(HTTPException) as raised:
        finish_demo(_confirmed_run(), config, configured_demo["reference"])

    assert raised.value.status_code == 409
    assert configured_demo["reference"].read_bytes() == configured_demo["original"]


def test_finish_demo_rechecks_hash_before_copying(configured_demo: dict[str, Any], tmp_path: Path) -> None:
    config = load_demo(configured_demo["root"], "tenant-a", "project-1")
    configured_demo["reference"].write_bytes(b"changed after load_demo")
    destination = tmp_path / "must-not-exist.xlsx"

    with pytest.raises(HTTPException) as raised:
        finish_demo(_confirmed_run(), config, destination)

    assert raised.value.status_code == 409
    assert not destination.exists()


def test_finish_demo_does_not_output_before_plan_confirmation(
    configured_demo: dict[str, Any], tmp_path: Path,
) -> None:
    config = load_demo(configured_demo["root"], "tenant-a", "project-1")
    run = _confirmed_run()
    run["plan_confirmation"]["confirmed"] = False
    destination = tmp_path / "unconfirmed.xlsx"

    with pytest.raises(HTTPException) as raised:
        finish_demo(run, config, destination)

    assert raised.value.status_code == 409
    assert not destination.exists()
    assert configured_demo["reference"].read_bytes() == configured_demo["original"]


def test_completed_demo_cannot_be_published_as_a_real_financial_result(
    configured_demo: dict[str, Any], tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    import backend.routers.agent as agent

    config = load_demo(configured_demo["root"], "tenant-a", "project-1")
    run = _confirmed_run()
    finish_demo(run, config, tmp_path / "演示输出.xlsx")
    monkeypatch.setattr(agent, "RUN_DIR", tmp_path / "runs")
    agent._save_run(run)
    releases: list[str] = []

    def release(*_args: Any, **_kwargs: Any) -> None:
        releases.append("released")
        raise AssertionError("Demo results must not enter real financial publication")

    monkeypatch.setattr(agent, "release_latest_financial_workbook_integration", release)
    user = SimpleNamespace(id="user-1", tenant_id="tenant-a")

    with pytest.raises(HTTPException) as raised:
        agent.publish_agent_run(run["run_id"], agent.AgentPublishIn(), user=user, db=None)

    assert raised.value.status_code == 409
    assert releases == []
    saved = agent._load_run(run["run_id"], user)
    assert saved["status"] == "completed"
    assert saved["validation"]["status"] == "demo_reference_match"
