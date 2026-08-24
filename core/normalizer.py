"""数据归一化引擎：文本 / 金额 / 日期统一清洗。

设计原则：
- 文本：保留前导零（如工号 '001'），去除所有空白，NaN/None -> ''。
- 金额：清除货币符号与千分位，异常 -> 0.0，保留两位小数。
- 日期：统一输出 'YYYY-MM-DD'，兼容字符串、datetime、Excel 序列号。
"""
import re
from datetime import datetime, date

import pandas as pd


class DataNormalizer:
    """静态方法集合，无状态，可在 DataFrame.apply 中安全复用。"""

    # 业务关键字段名（与 DuckDB 引擎保持一致）
    EMP_ID_COL = "工号"
    EMP_NAME_COL = "姓名"
    DEPT_COL = "部门"
    BASE_SALARY_COL = "基本工资"
    POSITION_SALARY_COL = "岗位工资"
    ID_CARD_COL = "身份证号"

    # 空白字符正则：半角空格、全角空格、制表符、换行
    _WHITESPACE_RE = re.compile(r"[\s\u3000]+")

    # ------------------------------------------------------------------
    # 文本
    # ------------------------------------------------------------------
    @staticmethod
    def clean_text(val) -> str:
        """转字符串，保留前导零，去除所有空白，NaN/None -> ''。"""
        if val is None:
            return ""
        # 统一处理 NaN / NaT / NA
        try:
            if pd.isna(val):
                return ""
        except (TypeError, ValueError):
            # pd.isna 对某些类型（如 list）会抛错，降级处理
            pass

        # float 整数值避免变成 '1.0'，直接转 int 字符串
        if isinstance(val, float):
            if val.is_integer():
                s = str(int(val))
            else:
                s = repr(val)
        else:
            s = str(val)

        # 去除所有空白（含全角空格、制表符、换行）
        s = DataNormalizer._WHITESPACE_RE.sub("", s)
        return s

    # ------------------------------------------------------------------
    # 金额
    # ------------------------------------------------------------------
    @staticmethod
    def clean_amount(val) -> float:
        """转 float，清除 ¥/，/空格，异常 -> 0.0，保留两位小数。"""
        if val is None:
            return 0.0
        try:
            if pd.isna(val):
                return 0.0
        except (TypeError, ValueError):
            pass

        # 数值类型直接走
        if isinstance(val, (int, float)):
            try:
                return round(float(val), 2)
            except (ValueError, TypeError, OverflowError):
                return 0.0

        s = str(val).strip()
        if not s:
            return 0.0
        # 去掉货币符号、千分位逗号（半角/全角）、空白
        s = (
            s.replace("¥", "")
            .replace("￥", "")
            .replace(",", "")
            .replace("，", "")
        )
        s = DataNormalizer._WHITESPACE_RE.sub("", s)
        # 去掉可能的前缀正负号空白
        s = s.strip()
        try:
            return round(float(s), 2)
        except (ValueError, TypeError):
            return 0.0

    # ------------------------------------------------------------------
    # 日期
    # ------------------------------------------------------------------
    @staticmethod
    def clean_date(val) -> str:
        """统一格式化为 'YYYY-MM-DD'，无法识别时返回原字符串。"""
        if val is None:
            return ""
        try:
            if pd.isna(val):
                return ""
        except (TypeError, ValueError):
            pass

        # 已是日期/时间对象
        if isinstance(val, (datetime, date, pd.Timestamp)):
            ts = pd.Timestamp(val)
            if pd.isna(ts):
                return ""
            return ts.strftime("%Y-%m-%d")

        s = str(val).strip()
        if not s:
            return ""

        # Excel 数字序列号：5 位数字通常对应 1982 年以后的日期
        # （序列号 30000 ≈ 1982-08-04，80000 ≈ 2119 年）
        if re.fullmatch(r"\d{5}", s):
            try:
                serial = int(s)
                if 30000 <= serial <= 80000:
                    return (
                        pd.Timestamp("1899-12-30") + pd.Timedelta(days=serial)
                    ).strftime("%Y-%m-%d")
            except (ValueError, OverflowError):
                pass

        # 兼容常见分隔符：/ . 统一替换为 -
        normalized = s.replace("/", "-").replace(".", "-")
        try:
            ts = pd.to_datetime(normalized, errors="raise")
            if pd.isna(ts):
                return s
            return ts.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return s  # 兜底返回原值，便于后续校验引擎标记异常

    # ------------------------------------------------------------------
    # 便捷批量方法
    # ------------------------------------------------------------------
    @classmethod
    def normalize_dataframe(cls, df: pd.DataFrame, text_cols=None,
                            amount_cols=None, date_cols=None) -> pd.DataFrame:
        """对 DataFrame 指定列批量归一化，返回副本。"""
        out = df.copy()
        if text_cols:
            for c in text_cols:
                if c in out.columns:
                    out[c] = out[c].apply(cls.clean_text)
        if amount_cols:
            for c in amount_cols:
                if c in out.columns:
                    out[c] = out[c].apply(cls.clean_amount)
        if date_cols:
            for c in date_cols:
                if c in out.columns:
                    out[c] = out[c].apply(cls.clean_date)
        return out
