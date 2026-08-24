"""DuckDB 内存数据中台：多源 Excel 视图注册、异动计算、主表富化。

核心能力：
1. stage_excel_sheets      : 把多张 Excel/DataFrame 注册为 DuckDB 视图。
2. get_personnel_changes   : SQL 计算当月新增入职 / 确认离职 / 薪资部门异动。
3. get_enriched_master_table : LEFT JOIN 主表与社保表，拼接身份证号等富化字段。
"""
from __future__ import annotations

import duckdb
import pandas as pd

from core.normalizer import DataNormalizer


class PayrollDBEngine:
    """基于 DuckDB 内存库的数据中台。"""

    # 默认业务字段（与 DataNormalizer 保持一致）
    EMP_ID_COL = DataNormalizer.EMP_ID_COL
    EMP_NAME_COL = DataNormalizer.EMP_NAME_COL
    DEPT_COL = DataNormalizer.DEPT_COL
    BASE_SALARY_COL = DataNormalizer.BASE_SALARY_COL
    POSITION_SALARY_COL = DataNormalizer.POSITION_SALARY_COL
    ID_CARD_COL = DataNormalizer.ID_CARD_COL

    # 默认参与异动比对的字段
    DEFAULT_COMPARE_COLS = (
        EMP_NAME_COL,
        DEPT_COL,
        BASE_SALARY_COL,
        POSITION_SALARY_COL,
    )

    def __init__(self):
        # 内存数据库，进程结束即释放，适合 Streamlit 会话级暂存
        self.con = duckdb.connect(database=":memory:")
        self._staged_views: list[str] = []

    # ------------------------------------------------------------------
    # 视图注册
    # ------------------------------------------------------------------
    def stage_excel_sheets(self, file_dict: dict) -> dict:
        """把上传的多个 Excel 表格加载为 DuckDB 视图。

        :param file_dict: {view_name: DataFrame | excel_file_path}
        :return: {view_name: row_count} 注册报告
        """
        report: dict[str, int] = {}
        for view_name, payload in file_dict.items():
            if isinstance(payload, str):
                df = pd.read_excel(payload)
            elif isinstance(payload, pd.DataFrame):
                df = payload.copy()
            else:
                raise TypeError(
                    f"stage_excel_sheets 仅支持 DataFrame 或文件路径，收到 {type(payload)}"
                )
            # 列名规整为字符串并去空白
            df.columns = [str(c).strip() for c in df.columns]
            # 视图名安全化（去空格、去点号）
            safe_name = str(view_name).replace(" ", "_").replace(".", "_")
            self.con.register(safe_name, df)
            self._staged_views.append(safe_name)
            report[safe_name] = len(df)
        return report

    def list_staged_views(self) -> list[str]:
        return list(self._staged_views)

    # ------------------------------------------------------------------
    # 异动计算
    # ------------------------------------------------------------------
    def get_personnel_changes(
        self,
        last_month_df: pd.DataFrame,
        current_sources_df: pd.DataFrame,
        id_col: str | None = None,
        compare_cols: list[str] | None = None,
    ) -> dict[str, pd.DataFrame]:
        """SQL 计算当月人员异动。

        :return: {
            "NEW_HIRES": 当月新增入职（当月有、上月无），
            "DEPARTED":  当月确认离职（上月有、当月无），
            "TRANSFERS": 薪资/部门异动（两边都有但比对字段有差异），
        }
        """
        id_col = id_col or self.EMP_ID_COL
        if compare_cols is None:
            compare_cols = list(self.DEFAULT_COMPARE_COLS)

        lm = last_month_df.copy()
        cs = current_sources_df.copy()
        lm.columns = [str(c).strip() for c in lm.columns]
        cs.columns = [str(c).strip() for c in cs.columns]

        if id_col not in lm.columns:
            raise KeyError(f"上月数据缺少主键列：{id_col}")
        if id_col not in cs.columns:
            raise KeyError(f"当月数据缺少主键列：{id_col}")

        # 工号归一化为字符串，避免数值/文本比较错位
        lm[id_col] = lm[id_col].apply(DataNormalizer.clean_text)
        cs[id_col] = cs[id_col].apply(DataNormalizer.clean_text)

        # 仅保留两边都存在的比对列
        effective_compare = [
            c for c in compare_cols if c in lm.columns and c in cs.columns
        ]

        self.con.register("_last_month", lm)
        self.con.register("_current", cs)

        # a) NEW_HIRES：当月有、上月无
        new_hires = self.con.execute(
            f"""
            SELECT c.* FROM _current c
            LEFT JOIN _last_month l ON c."{id_col}" = l."{id_col}"
            WHERE l."{id_col}" IS NULL
            """
        ).fetchdf()

        # b) DEPARTED：上月有、当月无
        departed = self.con.execute(
            f"""
            SELECT l.* FROM _last_month l
            LEFT JOIN _current c ON l."{id_col}" = c."{id_col}"
            WHERE c."{id_col}" IS NULL
            """
        ).fetchdf()

        # c) TRANSFERS：两边都有，但比对列有差异
        if effective_compare:
            select_prev_curr = ", ".join(
                f'l."{c}" AS prev_{c}, c."{c}" AS curr_{c}'
                for c in effective_compare
            )
            change_flags = ", ".join(
                f'(l."{c}" IS DISTINCT FROM c."{c}") AS changed_{c}'
                for c in effective_compare
            )
            where_changed = " OR ".join(
                f'(l."{c}" IS DISTINCT FROM c."{c}")' for c in effective_compare
            )
            transfers = self.con.execute(
                f"""
                SELECT l."{id_col}" AS emp_id,
                       {select_prev_curr},
                       {change_flags}
                FROM _last_month l
                JOIN _current c ON l."{id_col}" = c."{id_col}"
                WHERE {where_changed}
                """
            ).fetchdf()
        else:
            transfers = pd.DataFrame()

        return {
            "NEW_HIRES": new_hires,
            "DEPARTED": departed,
            "TRANSFERS": transfers,
        }

    # ------------------------------------------------------------------
    # 主表富化
    # ------------------------------------------------------------------
    def get_enriched_master_table(
        self,
        master_view: str = "master",
        social_security_view: str = "social_security",
        id_col: str | None = None,
        id_card_col: str | None = None,
        extra_enrich_cols: list[str] | None = None,
    ) -> pd.DataFrame:
        """SQL LEFT JOIN 关联主表与社保表，拼接身份证号（BX列）等富化字段。

        :param master_view: 主表视图名
        :param social_security_view: 社保表视图名
        :param id_col: 关联主键（默认 工号）
        :param id_card_col: 社保表中的身份证号列名（默认 身份证号，对应模板 BX 列）
        :param extra_enrich_cols: 额外需要从社保表带出的列（如 P/Q 列异动数据）
        :return: 富化后的 DataFrame
        """
        id_col = id_col or self.EMP_ID_COL
        id_card_col = id_card_col or self.ID_CARD_COL

        # 校验视图已注册
        registered = {r[0] for r in self.con.execute(
            "SELECT table_name FROM information_schema.tables WHERE table_schema='main'"
        ).fetchall()}
        # register() 注册的是 view，information_schema 可能不显示，改用内部跟踪
        registered = registered.union(set(self._staged_views))
        if master_view not in registered:
            raise KeyError(f"主表视图未注册：{master_view}")
        if social_security_view not in registered:
            raise KeyError(f"社保表视图未注册：{social_security_view}")

        extra_select = ""
        if extra_enrich_cols:
            extra_select = ", " + ", ".join(
                f'COALESCE(s."{c}", NULL) AS enrich_{c}' for c in extra_enrich_cols
            )

        sql = f"""
            SELECT m.*,
                   COALESCE(s."{id_card_col}", '') AS enriched_{id_card_col}{extra_select}
            FROM "{master_view}" m
            LEFT JOIN "{social_security_view}" s
              ON m."{id_col}" = s."{id_col}"
        """
        return self.con.execute(sql).fetchdf()

    # ------------------------------------------------------------------
    # 便捷接口
    # ------------------------------------------------------------------
    def query(self, sql: str) -> pd.DataFrame:
        """便捷 SQL 查询接口。"""
        return self.con.execute(sql).fetchdf()

    def close(self):
        self.con.close()

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass
