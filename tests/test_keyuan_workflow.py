from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest


def _source_paths(tmp_path: Path) -> dict[str, str]:
    names = [
        "薪资数据-科园-7月薪资（8.14发薪）.xlsx",
        "益药科园2026.07.xlsx",
        "鹤安长泰2026.07.xlsx",
        "202608_税款计算_工资薪金所得-益药科园.xls",
        "202608_税款计算_劳务报酬-益药科园.xls",
        "202608_税款计算_工资薪金所得-鹤安长泰.xls",
        "202608_税款计算_劳务报酬-鹤安长泰.xls",
    ]
    values: dict[str, str] = {}
    for index, name in enumerate(names):
        path = tmp_path / f"normalized-{index}.xlsx"
        path.write_bytes(b"PK\x03\x04")
        values[name] = str(path)
    return values


def test_detect_keyuan_batch_requires_the_complete_source_set(tmp_path: Path) -> None:
    from backend.keyuan_workflow import detect_keyuan_batch

    sources = _source_paths(tmp_path)
    batch = detect_keyuan_batch(
        "202607（所属月202606)-北京科园-鹤安-大药房工资核算总表-v2.xlsx",
        sources,
    )

    assert batch is not None
    assert batch.payroll_period == "2026.07"
    assert batch.payment_period == "2026.08"
    assert batch.output_filename == "202608（所属月202607）-北京科园-鹤安-大药房工资核算总表-v2.xlsx"
    assert batch.salary_source.name.startswith("normalized-")
    assert len(batch.social_sources) == 2
    assert len(batch.tax_sources) == 4

    sources.pop("202608_税款计算_劳务报酬-鹤安长泰.xls")
    assert detect_keyuan_batch(
        "202607（所属月202606)-北京科园-鹤安-大药房工资核算总表-v2.xlsx",
        sources,
    ) is None


def test_detect_keyuan_batch_does_not_claim_other_months_or_templates(tmp_path: Path) -> None:
    from backend.keyuan_workflow import detect_keyuan_batch

    sources = _source_paths(tmp_path)
    assert detect_keyuan_batch("通用工资总表.xlsx", sources) is None
    sources["202609_税款计算_工资薪金所得-益药科园.xls"] = sources.pop(
        "202608_税款计算_工资薪金所得-益药科园.xls"
    )
    assert detect_keyuan_batch(
        "202607（所属月202606)-北京科园-鹤安-大药房工资核算总表-v2.xlsx",
        sources,
    ) is None


def test_keyuan_batch_requires_the_project_to_use_the_payroll_period(tmp_path: Path) -> None:
    from backend.keyuan_workflow import detect_keyuan_batch

    batch = detect_keyuan_batch(
        "202607（所属月202606)-北京科园-鹤安-大药房工资核算总表-v2.xlsx",
        _source_paths(tmp_path),
    )

    assert batch is not None
    assert batch.matches_project_month("2026.07") is True
    assert batch.matches_project_month("2026.08") is False
    assert batch.matches_project_month("2026.09") is False


def test_execute_keyuan_batch_stages_normalized_inputs_and_returns_audited_output(
    tmp_path: Path,
) -> None:
    from backend.keyuan_workflow import detect_keyuan_batch, execute_keyuan_batch
    import scripts.keyuan_python_executor as executor

    sources = _source_paths(tmp_path)
    master = tmp_path / "normalized-master.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "工资核算"
    workbook.save(master)
    workbook.close()
    batch = detect_keyuan_batch(
        "202607（所属月202606)-北京科园-鹤安-大药房工资核算总表-v2.xlsx",
        sources,
    )
    assert batch is not None
    output = tmp_path / "result" / "draft.xlsx"

    def fake_process(master_path: Path, salary_path: Path, social_paths: object, tax_paths: object, output_path: Path) -> dict[str, object]:
        sample_dir = master_path.parent
        assert (sample_dir / batch.master_filename).is_file()
        assert (sample_dir / batch.salary_source_name).is_file()
        assert all((sample_dir / name).is_file() for name, _path in batch.social_sources)
        tax_dir = sample_dir / "回复：202608科园、鹤安应发工资的全部附件20260811"
        assert all((tax_dir / name).is_file() for name, _path in batch.tax_sources)
        workbook = openpyxl.load_workbook(master_path)
        workbook["工资核算"].insert_cols(1)
        workbook["工资核算"]["A2"] = "班制"
        workbook["工资核算"].cell(1, 99).value = "末列"
        workbook.create_sheet("台账")
        workbook.create_sheet("个税导出核对")
        workbook.create_sheet("OA请款及审批")
        workbook.save(output_path)
        workbook.close()
        return {"status": "passed", "changes": [{"sheet": "工资核算", "cell": "A2", "before": None, "after": "班制"}], "issues": []}

    monkeypatch = pytest.MonkeyPatch()
    monkeypatch.setattr(executor, "process_keyuan_python", fake_process)

    result = execute_keyuan_batch(batch, master_path=master, source_paths=sources, output_path=output, work_root=tmp_path / "work")
    monkeypatch.undo()

    assert result["status"] == "passed"
    assert result["change_count"] == 1
    assert result["download_filename"] == batch.output_filename
    assert result["structural"] == {
        "sheet_names": True,
        "payroll_columns": True,
        "payroll_header": True,
        "formula_references": True,
    }
    assert output.is_file()
    assert master.is_file()


def test_complete_keyuan_run_uses_fixed_executor_without_model(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import backend.agent_execution as execution
    from core.document_agent.orchestrator import ToolRegistry

    output = tmp_path / "Agent草稿.xlsx"
    captured: dict[str, object] = {}

    def fake_executor(batch: object, **kwargs: object) -> dict[str, object]:
        captured["batch"] = batch
        output_path = Path(kwargs["output_path"])
        workbook = openpyxl.Workbook()
        workbook.active.title = "工资核算"
        workbook["工资核算"]["A2"] = "班制"
        workbook["工资核算"].cell(1, 99).value = "末列"
        for name in ("台账", "个税导出核对", "OA请款及审批"):
            workbook.create_sheet(name)
        workbook.save(output_path)
        workbook.close()
        return {"status": "passed", "change_count": 2, "changes": [{"sheet": "工资核算", "cell": "A2"}], "issues": [], "structural": {}}

    monkeypatch.setattr(execution, "execute_keyuan_batch", fake_executor)
    monkeypatch.setattr(execution.ModelConfig, "from_env", lambda: None)
    monkeypatch.setattr(execution.ModelConfig, "fallback_from_env", lambda: None)
    monkeypatch.setattr(execution, "DATA_DIR", tmp_path)

    source_paths = _source_paths(tmp_path)
    master = tmp_path / "master.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "工资核算"
    workbook.save(master)
    workbook.close()
    run = {
        "run_id": "r" * 32,
        "master_file": "202607（所属月202606)-北京科园-鹤安-大药房工资核算总表-v2.xlsx",
        "salary_month": "2026.07",
        "_master_path": str(master),
        "_source_paths": source_paths,
        "workbook_updates": [],
        "messages": [],
        "model_plan": {"steps": []},
    }

    execution.execute_model_plan(
        run,
        draft=output,
        registry=ToolRegistry(),
        read_schemas=[],
        materials=[],
        rules={},
        save=lambda _run: None,
        emit=lambda *_args, **_kwargs: None,
    )

    assert captured["batch"] is not None
    assert run["keyuan_workflow"]["status"] == "passed"
    assert run["status"] == "completed"
    assert run["validation"]["status"] == "passed"
    assert output.is_file()
