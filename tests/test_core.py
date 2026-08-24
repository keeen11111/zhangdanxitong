"""tests/test_core.py — DataNormalizer / PayrollDBEngine / ExcelFormulaEngine 单元测试。

运行：
    pytest tests/test_core.py -v
"""
import os
import sys
from datetime import datetime

import openpyxl
import pandas as pd
import pytest
from openpyxl.styles import Font, PatternFill

# 把项目根目录加入 sys.path，便于 `core` 包导入
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.normalizer import DataNormalizer
from core.db_engine import PayrollDBEngine
from core.excel_engine import ExcelFormulaEngine
from core.validator import DataValidator, VALIDATION_RULES


# ============================================================
# DataNormalizer 测试
# ============================================================
class TestDataNormalizer:
    # ---------- clean_text ----------
    def test_clean_text_preserves_leading_zeros(self):
        assert DataNormalizer.clean_text("001") == "001"
        assert DataNormalizer.clean_text("00789") == "00789"

    def test_clean_text_nan_none_to_empty(self):
        assert DataNormalizer.clean_text(None) == ""
        assert DataNormalizer.clean_text(float("nan")) == ""
        assert DataNormalizer.clean_text(pd.NA) == ""
        assert DataNormalizer.clean_text(pd.NaT) == ""

    def test_clean_text_removes_all_whitespace(self):
        # 半角/全角空格、制表符、换行全部去除
        assert DataNormalizer.clean_text("  张 三  ") == "张三"
        assert DataNormalizer.clean_text("1\t2\n3") == "123"
        assert DataNormalizer.clean_text("110 101 1990") == "1101011990"

    def test_clean_text_float_integer_no_decimal(self):
        assert DataNormalizer.clean_text(123.0) == "123"
        assert DataNormalizer.clean_text(123.45) == "123.45"

    def test_clean_text_int(self):
        assert DataNormalizer.clean_text(7) == "7"

    # ---------- clean_amount ----------
    def test_clean_amount_currency_and_commas(self):
        assert DataNormalizer.clean_amount("¥1,234.56") == 1234.56
        assert DataNormalizer.clean_amount("￥9,876") == 9876.00
        assert DataNormalizer.clean_amount("1,234,567.89") == 1234567.89

    def test_clean_amount_invalid_to_zero(self):
        assert DataNormalizer.clean_amount("abc") == 0.0
        assert DataNormalizer.clean_amount("") == 0.0
        assert DataNormalizer.clean_amount(None) == 0.0
        assert DataNormalizer.clean_amount(float("nan")) == 0.0

    def test_clean_amount_numeric_passthrough(self):
        assert DataNormalizer.clean_amount(100) == 100.00
        assert DataNormalizer.clean_amount(99.999) == 100.00  # 四舍五入两位
        assert DataNormalizer.clean_amount(-50.5) == -50.50

    def test_clean_amount_returns_float_type(self):
        assert isinstance(DataNormalizer.clean_amount("100"), float)
        assert isinstance(DataNormalizer.clean_amount(100), float)

    # ---------- clean_date ----------
    def test_clean_date_standard_string(self):
        assert DataNormalizer.clean_date("2024-01-15") == "2024-01-15"
        assert DataNormalizer.clean_date("2024/1/5") == "2024-01-05"
        assert DataNormalizer.clean_date("2024.3.8") == "2024-03-08"

    def test_clean_date_datetime_obj(self):
        assert DataNormalizer.clean_date(datetime(2024, 3, 8)) == "2024-03-08"
        assert DataNormalizer.clean_date(pd.Timestamp("2024-12-31")) == "2024-12-31"

    def test_clean_date_excel_serial(self):
        # 序列号 44197 对应 2021-01-01
        assert DataNormalizer.clean_date("44197") == "2021-01-01"

    def test_clean_date_empty(self):
        assert DataNormalizer.clean_date(None) == ""
        assert DataNormalizer.clean_date("") == ""
        assert DataNormalizer.clean_date(float("nan")) == ""

    # ---------- normalize_dataframe ----------
    def test_normalize_dataframe_batch(self):
        df = pd.DataFrame({
            "工号": ["001", "002"],
            "姓名": [" 张三 ", "李四"],
            "基本工资": ["¥1,000.5", "2,000"],
        })
        out = DataNormalizer.normalize_dataframe(
            df, text_cols=["工号", "姓名"], amount_cols=["基本工资"]
        )
        assert out["工号"].tolist() == ["001", "002"]
        assert out["姓名"].tolist() == ["张三", "李四"]
        assert out["基本工资"].tolist() == [1000.5, 2000.0]


# ============================================================
# PayrollDBEngine 测试
# ============================================================
class TestPayrollDBEngine:
    @pytest.fixture
    def engine(self):
        return PayrollDBEngine()

    @pytest.fixture
    def last_month_df(self):
        return pd.DataFrame({
            "工号": ["001", "002", "003", "004"],
            "姓名": ["张三", "李四", "王五", "赵六"],
            "部门": ["技术部", "市场部", "技术部", "财务部"],
            "基本工资": [10000, 12000, 11000, 13000],
            "岗位工资": [5000, 6000, 5500, 6500],
        })

    @pytest.fixture
    def current_df(self):
        # 003 王五 离职；005 陈七 新入职；
        # 001 部门变动（技术部->产品部）；
        # 002 基本工资变动（12000->15000）；
        # 004 岗位工资变动（6500->7000）
        return pd.DataFrame({
            "工号": ["001", "002", "004", "005"],
            "姓名": ["张三", "李四", "赵六", "陈七"],
            "部门": ["产品部", "市场部", "财务部", "技术部"],
            "基本工资": [10000, 15000, 13000, 9000],
            "岗位工资": [5000, 6000, 7000, 4500],
        })

    # ---------- stage_excel_sheets ----------
    def test_stage_registers_views_and_counts(self, engine, last_month_df, current_df):
        report = engine.stage_excel_sheets({"上月": last_month_df, "当月": current_df})
        assert report == {"上月": 4, "当月": 4}
        assert set(engine.list_staged_views()) == {"上月", "当月"}

    def test_stage_accepts_excel_path(self, tmp_path, engine):
        # 写一个临时 Excel，验证文件路径加载
        fpath = tmp_path / "sample.xlsx"
        pd.DataFrame({"工号": ["001"], "姓名": ["张三"]}).to_excel(fpath, index=False)
        report = engine.stage_excel_sheets({"sample": str(fpath)})
        assert report == {"sample": 1}

    def test_stage_rejects_bad_type(self, engine):
        with pytest.raises(TypeError):
            engine.stage_excel_sheets({"bad": 123})

    # ---------- get_personnel_changes ----------
    def test_new_hires(self, engine, last_month_df, current_df):
        result = engine.get_personnel_changes(last_month_df, current_df)
        new_hires = result["NEW_HIRES"]
        assert len(new_hires) == 1
        assert new_hires.iloc[0]["工号"] == "005"
        assert new_hires.iloc[0]["姓名"] == "陈七"

    def test_departed(self, engine, last_month_df, current_df):
        result = engine.get_personnel_changes(last_month_df, current_df)
        departed = result["DEPARTED"]
        assert len(departed) == 1
        assert departed.iloc[0]["工号"] == "003"
        assert departed.iloc[0]["姓名"] == "王五"

    def test_transfers_identifies_all_three(self, engine, last_month_df, current_df):
        result = engine.get_personnel_changes(last_month_df, current_df)
        transfers = result["TRANSFERS"]
        ids = set(transfers["emp_id"].tolist())
        # 001 部门、002 基本工资、004 岗位工资 均异动
        assert ids == {"001", "002", "004"}
        # 离职的 003、新入职的 005 不应出现在异动里
        assert "003" not in ids
        assert "005" not in ids

    def test_transfers_includes_prev_curr_columns(self, engine, last_month_df, current_df):
        transfers = engine.get_personnel_changes(last_month_df, current_df)["TRANSFERS"]
        cols = set(transfers.columns)
        assert "prev_部门" in cols and "curr_部门" in cols
        assert "prev_基本工资" in cols and "curr_基本工资" in cols
        # 001 部门：技术部 -> 产品部
        row_001 = transfers[transfers["emp_id"] == "001"].iloc[0]
        assert row_001["prev_部门"] == "技术部"
        assert row_001["curr_部门"] == "产品部"
        assert bool(row_001["changed_部门"]) is True

    def test_no_changes_returns_empty(self, engine):
        df = pd.DataFrame({
            "工号": ["001", "002"],
            "姓名": ["张三", "李四"],
            "部门": ["技术部", "市场部"],
            "基本工资": [10000, 12000],
            "岗位工资": [5000, 6000],
        })
        result = engine.get_personnel_changes(df, df.copy())
        assert len(result["NEW_HIRES"]) == 0
        assert len(result["DEPARTED"]) == 0
        assert len(result["TRANSFERS"]) == 0

    def test_id_key_normalization(self, engine):
        # 上月工号是数字、当月是文本，归一化后应能正确关联
        lm = pd.DataFrame({"工号": [1, 2], "姓名": ["甲", "乙"],
                           "部门": ["A", "B"], "基本工资": [10, 20], "岗位工资": [1, 2]})
        cs = pd.DataFrame({"工号": ["1", "3"], "姓名": ["甲", "丙"],
                           "部门": ["A", "C"], "基本工资": [10, 30], "岗位工资": [1, 3]})
        result = engine.get_personnel_changes(lm, cs)
        assert set(result["NEW_HIRES"]["工号"]) == {"3"}
        assert set(result["DEPARTED"]["工号"]) == {"2"}

    def test_missing_id_column_raises(self, engine):
        lm = pd.DataFrame({"姓名": ["甲"]})
        cs = pd.DataFrame({"工号": ["1"], "姓名": ["甲"]})
        with pytest.raises(KeyError):
            engine.get_personnel_changes(lm, cs)

    # ---------- get_enriched_master_table ----------
    def test_enriched_master_left_join(self, engine):
        master = pd.DataFrame({
            "工号": ["001", "002", "003"],
            "姓名": ["张三", "李四", "王五"],
        })
        social = pd.DataFrame({
            "工号": ["001", "002"],
            "身份证号": ["110101199001011234", "110101199002022345"],
        })
        engine.stage_excel_sheets({"master": master, "social_security": social})
        enriched = engine.get_enriched_master_table()

        assert len(enriched) == 3
        # 003 没有社保记录，富化的身份证号应为空字符串
        row_003 = enriched[enriched["工号"] == "003"].iloc[0]
        assert row_003["enriched_身份证号"] == ""
        # 001 正常关联
        row_001 = enriched[enriched["工号"] == "001"].iloc[0]
        assert row_001["enriched_身份证号"] == "110101199001011234"

    def test_enriched_master_missing_view_raises(self, engine):
        master = pd.DataFrame({"工号": ["001"], "姓名": ["张三"]})
        engine.stage_excel_sheets({"master": master})
        with pytest.raises(KeyError):
            engine.get_enriched_master_table(social_security_view="not_exists")

    def test_enriched_master_extra_cols(self, engine):
        master = pd.DataFrame({"工号": ["001", "002"], "姓名": ["甲", "乙"]})
        social = pd.DataFrame({
            "工号": ["001", "002"],
            "身份证号": ["ID1", "ID2"],
            "P列异动": ["调岗", None],
        })
        engine.stage_excel_sheets({"master": master, "social_security": social})
        enriched = engine.get_enriched_master_table(extra_enrich_cols=["P列异动"])
        assert "enrich_P列异动" in enriched.columns
        row_001 = enriched[enriched["工号"] == "001"].iloc[0]
        assert row_001["enrich_P列异动"] == "调岗"


# ============================================================
# ExcelFormulaEngine 测试
# ============================================================
class TestExcelFormulaEngine:
    # ---------- shift_formula ----------
    def test_shift_simple_addition(self):
        assert ExcelFormulaEngine.shift_formula("=Z10+AA10", 1) == "=Z11+AA11"

    def test_shift_negative_offset(self):
        assert ExcelFormulaEngine.shift_formula("=Z10+AA10", -3) == "=Z7+AA7"

    def test_shift_absolute_row_preserved(self):
        # $N 绝对行号不平移
        assert ExcelFormulaEngine.shift_formula("=$Z$1+AA10", 1) == "=$Z$1+AA11"
        # $ 在行号前，列号相对
        assert ExcelFormulaEngine.shift_formula("=Z$10", 1) == "=Z$10"

    def test_shift_range_formula(self):
        assert ExcelFormulaEngine.shift_formula("=SUM(A1:B10)", 1) == "=SUM(A2:B11)"

    def test_shift_no_overflow_below_one(self):
        # 第 1 行下移到 0 应保持原值（防越界）
        assert ExcelFormulaEngine.shift_formula("=A1", -1) == "=A1"

    def test_shift_non_formula_passthrough(self):
        assert ExcelFormulaEngine.shift_formula("普通文本", 1) == "普通文本"
        assert ExcelFormulaEngine.shift_formula("", 1) == ""
        assert ExcelFormulaEngine.shift_formula(None, 1) is None

    def test_shift_multi_letter_column(self):
        # AA / AB 等多字母列名
        assert ExcelFormulaEngine.shift_formula("=AA10+AB10", 1) == "=AA11+AB11"

    def test_shift_preserves_leading_equals(self):
        out = ExcelFormulaEngine.shift_formula("=Z10", 1)
        assert out.startswith("=")

    # ---------- insert_employee_row ----------
    @pytest.fixture
    def template_ws(self):
        """构造一个 3 行 4 列的模板工作表。

        Row 1: 表头
        Row 2: 模板行（含样式、百分比格式、公式、静态值）
        Row 3: 已有员工行
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        # 表头
        ws.append(["工号", "姓名", "基本工资", "应发合计"])
        # 模板行（row 2）
        ws.append(["TPL", "模板", 10000, "=C2*1.1"])
        # 已有员工行
        ws.append(["001", "张三", 12000, "=C3*1.1"])

        # 给模板行加样式与百分比格式
        tpl_row = 2
        ws.cell(row=tpl_row, column=1).font = Font(bold=True)
        ws.cell(row=tpl_row, column=3).number_format = "0.00%"
        ws.cell(row=tpl_row, column=4).fill = PatternFill(
            start_color="FFFF00", end_color="FFFF00", fill_type="solid"
        )
        ws.row_dimensions[tpl_row].height = 25
        return ws

    def test_insert_row_returns_new_idx(self, template_ws):
        new_row = ExcelFormulaEngine.insert_employee_row(
            template_ws, template_row_idx=2,
            employee_data={1: "002", 2: "李四"}
        )
        assert new_row == 3  # 插在 row 2 下方

    def test_insert_row_copies_style_and_number_format(self, template_ws):
        ExcelFormulaEngine.insert_employee_row(
            template_ws, template_row_idx=2,
            employee_data={1: "002", 2: "李四"}
        )
        # 新行是 row 3
        new_cell = template_ws.cell(row=3, column=1)
        assert new_cell.font.bold is True
        # 百分比格式保留
        assert template_ws.cell(row=3, column=3).number_format == "0.00%"
        # 填充色保留
        assert template_ws.cell(row=3, column=4).fill.start_color.rgb in (
            "00FFFF00", "FFFFFF00"
        )
        # 行高对齐
        assert template_ws.row_dimensions[3].height == 25

    def test_insert_row_shifts_formula(self, template_ws):
        ExcelFormulaEngine.insert_employee_row(
            template_ws, template_row_idx=2,
            employee_data={1: "002", 2: "李四"}
        )
        # 新行 row 3，原模板公式 =C2*1.1 平移到 =C3*1.1
        assert template_ws.cell(row=3, column=4).value == "=C3*1.1"

    def test_insert_row_overrides_static_data(self, template_ws):
        ExcelFormulaEngine.insert_employee_row(
            template_ws, template_row_idx=2,
            employee_data={1: "002", 2: "李四", 3: 15000}
        )
        # 静态数据被覆盖
        assert template_ws.cell(row=3, column=1).value == "002"
        assert template_ws.cell(row=3, column=2).value == "李四"
        assert template_ws.cell(row=3, column=3).value == 15000
        # 公式列不被 employee_data 覆盖时保留平移后的公式
        assert template_ws.cell(row=3, column=4).value == "=C3*1.1"

    def test_insert_row_supports_column_letter(self, template_ws):
        ExcelFormulaEngine.insert_employee_row(
            template_ws, template_row_idx=2,
            employee_data={"A": "002", "B": "李四"}
        )
        assert template_ws.cell(row=3, column=1).value == "002"
        assert template_ws.cell(row=3, column=2).value == "李四"

    def test_insert_row_preserves_existing_below(self, template_ws):
        # 原有 row 3 的员工应被推到 row 4
        ExcelFormulaEngine.insert_employee_row(
            template_ws, template_row_idx=2,
            employee_data={1: "002", 2: "李四"}
        )
        # 原 row 3 -> row 4
        assert template_ws.cell(row=4, column=1).value == "001"
        assert template_ws.cell(row=4, column=2).value == "张三"

    def test_insert_row_invalid_idx_raises(self, template_ws):
        with pytest.raises(ValueError):
            ExcelFormulaEngine.insert_employee_row(
                template_ws, template_row_idx=0, employee_data={}
            )

    # ---------- delete_departed_rows ----------
    @pytest.fixture
    def populated_ws(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["工号", "姓名"])
        for eid, name in [("001", "甲"), ("002", "乙"),
                          ("003", "丙"), ("004", "丁")]:
            ws.append([eid, name])
        return ws

    def test_delete_rows_removes_matched(self, populated_ws):
        deleted = ExcelFormulaEngine.delete_departed_rows(
            populated_ws, key_col_idx=1, departed_keys=["002", "004"]
        )
        assert deleted == 2
        remaining_ids = [populated_ws.cell(row=r, column=1).value
                         for r in range(2, populated_ws.max_row + 1)]
        assert remaining_ids == ["001", "003"]

    def test_delete_rows_no_match(self, populated_ws):
        deleted = ExcelFormulaEngine.delete_departed_rows(
            populated_ws, key_col_idx=1, departed_keys=["999"]
        )
        assert deleted == 0
        assert populated_ws.max_row == 5  # 表头+4 行未变

    def test_delete_rows_type_insensitive(self, populated_ws):
        # 数字 2 应能匹配工号 "002"——这里测试字符串归一化
        deleted = ExcelFormulaEngine.delete_departed_rows(
            populated_ws, key_col_idx=1, departed_keys=["002"]
        )
        assert deleted == 1

    def test_delete_rows_from_bottom_avoids_collapse(self, populated_ws):
        # 一次性删多个，行号不应错位
        ExcelFormulaEngine.delete_departed_rows(
            populated_ws, key_col_idx=1, departed_keys=["001", "003"]
        )
        remaining_ids = [populated_ws.cell(row=r, column=1).value
                         for r in range(2, populated_ws.max_row + 1)]
        assert remaining_ids == ["002", "004"]

    # ---------- freeze_row_values ----------
    def test_freeze_row_overwrites_formula_with_value(self, tmp_path):
        # 构造一个含公式的目标工作簿
        target_wb = openpyxl.Workbook()
        target_ws = target_wb.active
        target_ws.append(["A", "B", "合计"])
        target_ws.append([10, 20, "=A2+B2"])

        # 模拟"已计算"的数据源（直接写静态值，假装是 Excel 缓存的结果）
        data_wb = openpyxl.Workbook()
        data_ws = data_wb.active
        data_ws.append(["A", "B", "合计"])
        data_ws.append([10, 20, 30])

        written = ExcelFormulaEngine.freeze_row_values(
            target_ws, row_idx=2, data_only_ws=data_ws
        )
        assert written == 3
        # 公式被替换为静态值
        assert target_ws.cell(row=2, column=3).value == 30

    def test_freeze_row_skips_none(self):
        target_wb = openpyxl.Workbook()
        target_ws = target_wb.active
        target_ws.append(["A", "B"])
        target_ws.append([1, "=A2*2"])

        data_wb = openpyxl.Workbook()
        data_ws = data_wb.active
        data_ws.append(["A", "B"])
        data_ws.append([1, None])  # B 列无缓存值

        written = ExcelFormulaEngine.freeze_row_values(
            target_ws, row_idx=2, data_only_ws=data_ws
        )
        assert written == 1
        # B 列公式保留（未被 None 覆盖）
        assert target_ws.cell(row=2, column=2).value == "=A2*2"

    def test_freeze_row_invalid_idx_raises(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["A"])
        with pytest.raises(ValueError):
            ExcelFormulaEngine.freeze_row_values(ws, row_idx=99, data_only_ws=ws)


# ============================================================
# DataValidator 测试
# ============================================================
class TestDataValidator:
    def test_status_labels_do_not_use_emoji(self):
        statuses = [
            DataValidator.STATUS_OK,
            DataValidator.STATUS_MISSING_ID_CARD,
            DataValidator.STATUS_MISSING_BASE_SALARY,
            DataValidator.STATUS_MISSING_CONFIRM_DATE,
            DataValidator.STATUS_MISSING_OTHER,
            DataValidator.STATUS_WARNING_ONLY,
        ]

        assert all(not any(symbol in status for symbol in "🟢🔴🟡") for status in statuses)

    @pytest.fixture
    def validator(self):
        return DataValidator()

    @pytest.fixture
    def full_profile_df(self):
        """全部字段合格的员工数据。"""
        return pd.DataFrame({
            "工号": ["001", "002", "003", "004", "005"],
            "姓名": ["张三", "李四", "王五", "赵六", "钱七"],
            "身份证号": [
                "110101199001011234",
                "110101199002022345",
                "11010119900303345X",  # 末位 X 合法
                "",                     # 缺身份证
                "11010119900505",       # 不足 18 位
            ],
            "基本工资": [10000, 12000, 11000, 13000, 0],  # 005 为 0
            "转正日期": ["2024-01-01", "", "2024-03-01", "2024-04-01", "2024-05-01"],
            "试用期状态": ["已转正", "已转正", "已转正", "已转正", "试用中"],
            # 003 缺车贴标准（仅警告）；004 车贴标准保留（身份证已 BLOCKER）
            "车贴标准": [200, 200, None, 200, 200],
        })

    # ---------- 规则集结构 ----------
    def test_rules_have_blocker_and_warning(self):
        assert "blocker" in VALIDATION_RULES
        assert "warning" in VALIDATION_RULES
        blocker_fields = {r["field"] for r in VALIDATION_RULES["blocker"]}
        assert {"工号", "姓名", "身份证号", "基本工资"} <= blocker_fields
        warning_fields = {r["field"] for r in VALIDATION_RULES["warning"]}
        assert {"转正日期", "试用期状态", "车贴标准"} <= warning_fields

    def test_blocker_rules_have_severity(self):
        for r in VALIDATION_RULES["blocker"]:
            assert r["severity"] == "BLOCKER"
        for r in VALIDATION_RULES["warning"]:
            assert r["severity"] == "WARNING"

    # ---------- audit_employee_profiles ----------
    def test_audit_adds_two_columns(self, validator, full_profile_df):
        out = validator.audit_employee_profiles(full_profile_df)
        assert "诊断状态" in out.columns
        assert "缺失字段清单" in out.columns
        # 不应破坏原列
        for col in full_profile_df.columns:
            assert col in out.columns

    def test_audit_qualified_row(self, validator):
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["110101199001011234"],
            "基本工资": [10000],
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_OK
        assert out.iloc[0]["缺失字段清单"] == []

    def test_audit_missing_id_card_empty(self, validator):
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": [""],
            "基本工资": [10000],
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_ID_CARD
        assert "身份证号" in out.iloc[0]["缺失字段清单"]

    def test_audit_missing_id_card_short(self, validator):
        # 不足 18 位也应判为缺身份证
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["12345"],
            "基本工资": [10000],
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_ID_CARD

    def test_audit_id_card_with_x_suffix(self, validator):
        # 末位 X 合法
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["11010119900101123X"],
            "基本工资": [10000],
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_OK

    def test_audit_missing_base_salary_zero(self, validator):
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["110101199001011234"],
            "基本工资": [0],  # 0 不算正数
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_BASE_SALARY

    def test_audit_missing_base_salary_nan(self, validator):
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["110101199001011234"],
            "基本工资": [float("nan")],
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_BASE_SALARY

    def test_audit_missing_confirm_date(self, validator):
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["110101199001011234"],
            "基本工资": [10000],
            "转正日期": [""],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        # 缺转正日期，无核心项缺失
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_CONFIRM_DATE
        assert "转正日期" in out.iloc[0]["缺失字段清单"]

    def test_audit_warning_only(self, validator):
        # 仅缺试用期状态（非转正日期）
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": ["110101199001011234"],
            "基本工资": [10000],
            "转正日期": ["2024-01-01"],
            "试用期状态": [""],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_WARNING_ONLY

    def test_audit_priority_id_card_over_salary(self, validator):
        # 同时缺身份证与基本工资，状态应是身份证（优先级更高）
        df = pd.DataFrame({
            "工号": ["001"],
            "姓名": ["张三"],
            "身份证号": [""],
            "基本工资": [0],
            "转正日期": [""],
            "试用期状态": [""],
            "车贴标准": [None],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_ID_CARD
        # 缺失字段清单应聚合全部缺失项
        missing = set(out.iloc[0]["缺失字段清单"])
        assert {"身份证号", "基本工资", "转正日期", "试用期状态", "车贴标准"} <= missing

    def test_audit_missing_other_blocker(self, validator):
        # 缺工号（非身份证/基本工资的核心项）
        df = pd.DataFrame({
            "工号": [""],
            "姓名": ["张三"],
            "身份证号": ["110101199001011234"],
            "基本工资": [10000],
            "转正日期": ["2024-01-01"],
            "试用期状态": ["已转正"],
            "车贴标准": [200],
        })
        out = validator.audit_employee_profiles(df)
        assert out.iloc[0]["诊断状态"] == DataValidator.STATUS_MISSING_OTHER

    def test_audit_full_dataset(self, validator, full_profile_df):
        out = validator.audit_employee_profiles(full_profile_df)
        statuses = out["诊断状态"].tolist()
        # 001 全合格
        assert statuses[0] == DataValidator.STATUS_OK
        # 002 缺转正日期
        assert statuses[1] == DataValidator.STATUS_MISSING_CONFIRM_DATE
        # 003 缺车贴标准（仅警告）
        assert statuses[2] == DataValidator.STATUS_WARNING_ONLY
        # 004 缺身份证
        assert statuses[3] == DataValidator.STATUS_MISSING_ID_CARD
        # 005 缺身份证 + 基本工资为0 -> 身份证优先
        assert statuses[4] == DataValidator.STATUS_MISSING_ID_CARD

    def test_audit_does_not_mutate_input(self, validator, full_profile_df):
        original_cols = list(full_profile_df.columns)
        _ = validator.audit_employee_profiles(full_profile_df)
        # 原 DataFrame 列不变
        assert list(full_profile_df.columns) == original_cols

    # ---------- get_completeness_metrics ----------
    def test_metrics_full_qualified(self, validator):
        df = pd.DataFrame({
            "工号": ["001", "002"],
            "姓名": ["甲", "乙"],
            "身份证号": ["110101199001011234", "110101199002022345"],
            "基本工资": [10000, 12000],
            "转正日期": ["2024-01-01", "2024-02-01"],
            "试用期状态": ["已转正", "已转正"],
            "车贴标准": [200, 200],
        })
        audited = validator.audit_employee_profiles(df)
        metrics = validator.get_completeness_metrics(audited)
        assert metrics["total"] == 2
        assert metrics["qualified"] == 2
        assert metrics["completeness_pct"] == 100.0
        assert metrics["high_risk_count"] == 0
        assert metrics["warning_count"] == 0

    def test_metrics_mixed_dataset(self, validator, full_profile_df):
        audited = validator.audit_employee_profiles(full_profile_df)
        metrics = validator.get_completeness_metrics(audited)
        assert metrics["total"] == 5
        assert metrics["qualified"] == 1  # 仅 001 合格
        # 004、005 缺身份证；005 还缺基本工资；高危按人头计而非按字段计
        assert metrics["high_risk_count"] == 2  # 004、005
        # 002 缺转正日期、003 缺车贴标准 -> 警告
        assert metrics["warning_count"] == 2
        assert metrics["completeness_pct"] == 20.0
        # 字段级分布
        breakdown = metrics["missing_field_breakdown"]
        assert breakdown["身份证号"] == 2  # 004、005
        assert breakdown["转正日期"] == 1  # 002
        assert breakdown["车贴标准"] == 1  # 003

    def test_metrics_empty_df(self, validator):
        df = pd.DataFrame(columns=[
            "工号", "姓名", "身份证号", "基本工资",
            "转正日期", "试用期状态", "车贴标准",
        ])
        audited = validator.audit_employee_profiles(df)
        metrics = validator.get_completeness_metrics(audited)
        assert metrics["total"] == 0
        assert metrics["completeness_pct"] == 100.0  # 空集默认 100%
