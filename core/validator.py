"""员工数据完整性校验引擎：规则集 + 行级诊断 + 完整度指标。

设计原则：
- 规则集声明式定义：核心项（Blocker）/ 警告项（Warning），便于 UI 高亮渲染。
- 行级诊断：返回带 `诊断状态` 与 `缺失字段清单` 的 DataFrame，支持网格双击补齐。
- 完整度指标：合格率 + 高危缺项数，供控制塔仪表盘直接消费。
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd

from core.normalizer import DataNormalizer


# ============================================================
# 校验规则集
# ============================================================
VALIDATION_RULES: dict[str, list[dict]] = {
    # 🔴 核心项 (Blocker)：缺失或不合规将阻断后续处理
    "blocker": [
        {
            "field": "工号",
            "label": "工号",
            "check": "not_empty",
            "severity": "BLOCKER",
            "message": "工号缺失",
        },
        {
            "field": "姓名",
            "label": "姓名",
            "check": "not_empty",
            "severity": "BLOCKER",
            "message": "姓名缺失",
        },
        {
            "field": "身份证号",
            "label": "身份证号",
            "check": "id_card_18",
            "severity": "BLOCKER",
            "message": "身份证号缺失或非18位",
        },
        {
            "field": "基本工资",
            "label": "基本工资",
            "check": "amount_positive",
            "severity": "BLOCKER",
            "message": "基本工资缺失或非正数",
        },
    ],
    # 🟡 警告项 (Warning)：缺失会标记但不阻断
    "warning": [
        {
            "field": "转正日期",
            "label": "转正日期",
            "check": "not_empty",
            "severity": "WARNING",
            "message": "转正日期缺失",
        },
        {
            "field": "试用期状态",
            "label": "试用期状态",
            "check": "not_empty",
            "severity": "WARNING",
            "message": "试用期状态缺失",
        },
        {
            "field": "车贴标准",
            "label": "车贴标准",
            "check": "not_empty",
            "severity": "WARNING",
            "message": "车贴标准缺失",
        },
    ],
}


# 18 位身份证号基础格式校验（前 17 位数字 + 末位数字或 X）
_ID_CARD_RE = re.compile(r"^\d{17}[\dXx]$")


class DataValidator:
    """员工数据完整性校验引擎。"""

    # 诊断状态枚举
    STATUS_OK = "合格"
    STATUS_MISSING_ID_CARD = "阻断 · 缺身份证"
    STATUS_MISSING_BASE_SALARY = "阻断 · 缺基本工资"
    STATUS_MISSING_CONFIRM_DATE = "提醒 · 缺转正日期"
    STATUS_MISSING_OTHER = "阻断 · 缺核心项"
    STATUS_WARNING_ONLY = "提醒 · 仅警告项缺失"

    def __init__(self, rules: dict[str, list[dict]] | None = None):
        self.rules = rules if rules is not None else VALIDATION_RULES

    # ------------------------------------------------------------------
    # 行级校验
    # ------------------------------------------------------------------
    def audit_employee_profiles(self, df: pd.DataFrame) -> pd.DataFrame:
        """对每一行进行诊断，返回带 `诊断状态` 与 `缺失字段清单` 的副本。

        :param df: 员工数据 DataFrame
        :return:   新增两列的 DataFrame
                   - `诊断状态`：枚举值
                   - `缺失字段清单`：list[str]，缺失的具体字段名
        """
        out = df.copy()
        statuses: list[str] = []
        missing_lists: list[list[str]] = []

        blocker_rules = self.rules.get("blocker", [])
        warning_rules = self.rules.get("warning", [])

        for _, row in df.iterrows():
            missing_blockers: list[str] = []
            missing_warnings: list[str] = []

            # 逐条评估 Blocker 规则
            for rule in blocker_rules:
                field = rule["field"]
                if not self._check_rule(row, field, rule["check"]):
                    missing_blockers.append(rule["label"])

            # 逐条评估 Warning 规则
            for rule in warning_rules:
                field = rule["field"]
                if not self._check_rule(row, field, rule["check"]):
                    missing_warnings.append(rule["label"])

            # 优先级：Blocker > Warning
            status, missing = self._resolve_status(
                missing_blockers, missing_warnings
            )
            statuses.append(status)
            missing_lists.append(missing)

        out["诊断状态"] = statuses
        out["缺失字段清单"] = missing_lists
        return out

    # ------------------------------------------------------------------
    # 完整度指标
    # ------------------------------------------------------------------
    @staticmethod
    def get_completeness_metrics(audited_df: pd.DataFrame) -> dict:
        """计算数据完整度指标。

        :return: {
            "total":           总人数,
            "qualified":       合格人数（🟢）,
            "completeness_pct": 完整度百分比（0-100，保留两位小数）,
            "high_risk_count": 高危缺项人数（任意 🔴 状态）,
            "warning_count":   仅警告缺项人数（🟡）,
            "missing_field_breakdown": {字段名: 缺失人数} 详细分布,
        }
        """
        total = len(audited_df)
        if total == 0:
            return {
                "total": 0,
                "qualified": 0,
                "completeness_pct": 100.0,
                "high_risk_count": 0,
                "warning_count": 0,
                "missing_field_breakdown": {},
            }

        status_col = audited_df["诊断状态"]
        qualified = int((status_col == DataValidator.STATUS_OK).sum())
        blocker_statuses = {
            DataValidator.STATUS_MISSING_ID_CARD,
            DataValidator.STATUS_MISSING_BASE_SALARY,
            DataValidator.STATUS_MISSING_OTHER,
        }
        warning_statuses = {
            DataValidator.STATUS_MISSING_CONFIRM_DATE,
            DataValidator.STATUS_WARNING_ONLY,
        }
        high_risk = int(status_col.isin(blocker_statuses).sum())
        warning = int(status_col.isin(warning_statuses).sum())

        completeness_pct = round(qualified / total * 100, 2)

        # 字段级缺失分布
        breakdown: dict[str, int] = {}
        for lst in audited_df["缺失字段清单"]:
            for f in lst:
                breakdown[f] = breakdown.get(f, 0) + 1

        return {
            "total": total,
            "qualified": qualified,
            "completeness_pct": completeness_pct,
            "high_risk_count": high_risk,
            "warning_count": warning,
            "missing_field_breakdown": breakdown,
        }

    # ------------------------------------------------------------------
    # 内部：单条规则评估
    # ------------------------------------------------------------------
    @staticmethod
    def _check_rule(row: pd.Series, field: str, check: str) -> bool:
        """返回 True 表示通过，False 表示缺失/不合规。"""
        val = row.get(field)
        if check == "not_empty":
            text = DataNormalizer.clean_text(val)
            return text != ""
        if check == "id_card_18":
            text = DataNormalizer.clean_text(val)
            if not text:
                return False
            return bool(_ID_CARD_RE.match(text))
        if check == "amount_positive":
            amt = DataNormalizer.clean_amount(val)
            return amt > 0
        # 未知 check 类型默认通过
        return True

    # ------------------------------------------------------------------
    # 内部：状态汇总
    # ------------------------------------------------------------------
    @classmethod
    def _resolve_status(
        cls,
        missing_blockers: list[str],
        missing_warnings: list[str],
    ) -> tuple[str, list[str]]:
        """根据 Blocker / Warning 缺失清单解析最终状态与缺失字段聚合清单。

        优先级（从高到低）：
            1. 缺身份证号         -> 🔴 缺身份证
            2. 缺基本工资         -> 🔴 缺基本工资
            3. 其他核心项缺失     -> 🔴 缺核心项
            4. 缺转正日期         -> 🟡 缺转正日期
            5. 其他警告项缺失     -> 🟡 仅警告项缺失
            6. 全部通过           -> 🟢 合格
        """
        all_missing = missing_blockers + missing_warnings

        if missing_blockers:
            # 优先级 1/2：身份证号、基本工资 单独标签便于 UI 聚焦
            if "身份证号" in missing_blockers:
                return cls.STATUS_MISSING_ID_CARD, all_missing
            if "基本工资" in missing_blockers:
                return cls.STATUS_MISSING_BASE_SALARY, all_missing
            return cls.STATUS_MISSING_OTHER, all_missing

        if missing_warnings:
            if "转正日期" in missing_warnings:
                return cls.STATUS_MISSING_CONFIRM_DATE, all_missing
            return cls.STATUS_WARNING_ONLY, all_missing

        return cls.STATUS_OK, all_missing
