"""AutoPayroll-Pro 可视化薪资控制塔 —— Streamlit 主界面（向导式 SaaS）。

流程：
  Step 1 📂 数据源导入   —— 上传后即时预览，不强求三表齐全
  Step 2 ⚙️ 字段映射与异动 —— 主键列嗅探+手动确认，按需异动计算
  Step 3 📊 异动看板 / ⚠️ 诊断补齐 —— 两大工作 Tab
  Step 4 🚀 模板渲染与导出

核心改进：
  - 主键列不再硬编码"工号"，支持别名嗅探（员工编号/工 号/ID 等）+ 手动选择
  - 单表也可进入诊断流程（无需上月+本月配对）
  - 上传后即时预览，所见即所得
"""
from __future__ import annotations

import io
import os
import sys
import traceback
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.db_engine import PayrollDBEngine
from core.excel_engine import ExcelFormulaEngine
from core.normalizer import DataNormalizer
from core.validator import DataValidator

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="AutoPayroll-Pro 薪资控制塔",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 1.2rem 1.5rem; border-radius: 12px; color: white;
        margin-bottom: 1rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.6rem; }
    .main-header p { color: rgba(255,255,255,0.85); margin: 0.3rem 0 0; font-size: 0.9rem; }

    [data-testid="stMetric"] {
        background: #f8f9ff; border-radius: 10px; padding: 1rem;
        border-left: 4px solid #667eea; box-shadow: 0 2px 6px rgba(0,0,0,0.04);
    }
    .blocker-box {
        background: #fff5f5; border: 2px solid #ff6b6b; border-radius: 8px;
        padding: 1rem; color: #c92a2a;
    }
    .warning-box {
        background: #fffbe6; border: 2px solid #ffd43b; border-radius: 8px;
        padding: 1rem; color: #866a04;
    }
    .success-box {
        background: #f0fff4; border: 2px solid #51cf66; border-radius: 8px;
        padding: 1rem; color: #2b8a3e;
    }
    /* 向导步骤条 */
    .stepper {
        display: flex; gap: 0.5rem; margin: 0.5rem 0 1rem;
    }
    .step-pill {
        flex: 1; padding: 0.5rem 0.8rem; border-radius: 20px; text-align: center;
        font-size: 0.85rem; font-weight: 600;
    }
    .step-done   { background: #d3f9d8; color: #2b8a3e; }
    .step-active { background: #667eea; color: white; }
    .step-todo   { background: #e9ecef; color: #868e96; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 主键列别名嗅探
# ============================================================
_ID_ALIASES = [
    "工号", "员工工号", "员工编号", "工 号", "編號", "员工ID", "ID",
    "Employee ID", "Emp ID", "工号 ", " 工号", "emp_no", "工牌号",
]


def _detect_id_column(df: pd.DataFrame) -> str | None:
    """嗅探 DataFrame 中最可能的主键列，返回列名或 None。"""
    if df is None or df.empty:
        return None
    cols = [str(c).strip() for c in df.columns]
    # 1. 精确匹配别名
    for alias in _ID_ALIASES:
        for c in cols:
            if c == alias:
                return c
    # 2. 包含"工号"/"编号"/"ID"的列
    for c in cols:
        if any(k in c for k in ("工号", "编号", "ID", "id")):
            return c
    # 3. 第一列兜底
    return cols[0] if cols else None


def _detect_name_column(df: pd.DataFrame) -> str | None:
    """嗅探姓名列。"""
    if df is None or df.empty:
        return None
    for c in df.columns:
        cs = str(c).strip()
        if cs in ("姓名", "员工姓名", "名字", "Name"):
            return cs
    return None


# ============================================================
# Session State
# ============================================================
def _init_state():
    defaults = {
        # 原始上传
        "last_month_raw": None,     # BytesIO（保留用于导出时重新加载公式）
        "last_month_df": None,
        "current_df": None,
        "social_dfs": {},
        "social_raws": {},
        # 配置
        "salary_month": "2026.06",
        "id_col_lm": None,          # 上月主键列
        "id_col_cs": None,          # 本月主键列
        "compare_cols": [],
        # 处理结果
        "diff_result": None,
        "master_df": None,          # 合并后的主表（诊断用）
        "audited_df": None,
        "edited_df": None,
        "log_lines": [],
        "export_buffer": None,
        "export_filename": None,
        # 流程阶段
        "stage": "import",          # import -> config -> work -> export
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()


def _log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.log_lines.append(f"[{ts}] {msg}")


def _set_stage(s: str):
    st.session_state.stage = s


# ============================================================
# 顶部标题 + 向导步骤条
# ============================================================
st.markdown(
    """
    <div class="main-header">
        <h1>🚀 AutoPayroll-Pro 可视化薪资控制塔</h1>
        <p>向导式数据导入 · 主键智能映射 · 异动 Diff · 在线补齐 · 公式保持导出</p>
    </div>
    """,
    unsafe_allow_html=True,
)

_stage_labels = {
    "import": ("1.数据导入", "config", "2.字段映射", "work", "3.诊断补齐", "export", "4.导出"),
}


def _render_stepper():
    stages = [("import", "1.数据导入"), ("config", "2.字段映射"),
              ("work", "3.诊断补齐"), ("export", "4.导出")]
    cur = st.session_state.stage
    order = ["import", "config", "work", "export"]
    cur_idx = order.index(cur) if cur in order else 0
    pills = ""
    for i, (key, label) in enumerate(stages):
        if i < cur_idx:
            cls = "step-done"
            icon = "✅ "
        elif i == cur_idx:
            cls = "step-active"
            icon = "▶ "
        else:
            cls = "step-todo"
            icon = ""
        pills += f'<div class="step-pill {cls}">{icon}{label}</div>'
    st.markdown(f'<div class="stepper">{pills}</div>', unsafe_allow_html=True)


_render_stepper()


# ============================================================
# 侧边栏：文件上传区（始终可见）
# ============================================================
with st.sidebar:
    st.markdown("### 📂 数据源上传")
    st.caption("上传后即时预览，可单独导入某一张表")

    last_month_file = st.file_uploader(
        "【上月主 Excel 模板】", type=["xlsx", "xls"], key="upload_last_month",
        help="带公式与历史数据的 Master 模板",
    )
    current_file = st.file_uploader(
        "【本月薪资异动表】", type=["xlsx", "xls"], key="upload_current",
        help="本月人员/薪资异动明细（可选，仅诊断时可不传）",
    )
    social_files = st.file_uploader(
        "【多张社保公积金表】", type=["xlsx", "xls"], accept_multiple_files=True,
        key="upload_social", help="用于跨表关联补齐身份证号（可选）",
    )

    # 即时读入上传文件（不等点按钮）
    if last_month_file is not None:
        try:
            df = pd.read_excel(last_month_file)
            if st.session_state.last_month_df is None or len(df) != len(st.session_state.last_month_df):
                st.session_state.last_month_df = df
                # 保留原始字节用于导出
                last_month_file.seek(0)
                st.session_state.last_month_raw = io.BytesIO(last_month_file.read())
                # 嗅探主键列
                st.session_state.id_col_lm = _detect_id_column(df)
                _log(f"上月主表导入：{len(df)} 行 × {len(df.columns)} 列，嗅探主键列=[{st.session_state.id_col_lm}]")
        except Exception as e:
            st.error(f"上月主表读取失败：{e}")

    if current_file is not None:
        try:
            df = pd.read_excel(current_file)
            if st.session_state.current_df is None or len(df) != len(st.session_state.current_df):
                st.session_state.current_df = df
                st.session_state.id_col_cs = _detect_id_column(df)
                _log(f"本月异动表导入：{len(df)} 行 × {len(df.columns)} 列，嗅探主键列=[{st.session_state.id_col_cs}]")
        except Exception as e:
            st.error(f"本月异动表读取失败：{e}")

    if social_files:
        new_social = {}
        new_raws = {}
        for f in social_files:
            try:
                df = pd.read_excel(f)
                name = f.name.rsplit(".", 1)[0]
                new_social[name] = df
                f.seek(0)
                new_raws[name] = io.BytesIO(f.read())
            except Exception as e:
                st.error(f"社保表 {f.name} 读取失败：{e}")
        if new_social:
            st.session_state.social_dfs = new_social
            st.session_state.social_raws = new_raws
            _log(f"社保表导入：{list(new_social.keys())}")

    st.markdown("---")
    st.session_state.salary_month = st.text_input(
        "当前薪资月份", value=st.session_state.salary_month)

    # 导入状态摘要
    st.markdown("### 📋 导入状态")
    s_ok, s_icon = (True, "✅") if st.session_state.last_month_df is not None else (False, "⬜")
    st.write(f"{s_icon} 上月主表" + (f"（{len(st.session_state.last_month_df)} 行）" if s_ok else ""))
    c_ok, c_icon = (True, "✅") if st.session_state.current_df is not None else (False, "⬜")
    st.write(f"{c_icon} 本月异动表" + (f"（{len(st.session_state.current_df)} 行）" if c_ok else "（可选）"))
    sc_icon = "✅" if st.session_state.social_dfs else "⬜"
    st.write(f"{sc_icon} 社保表" + (f"（{len(st.session_state.social_dfs)} 张）" if st.session_state.social_dfs else "（可选）"))


# ============================================================
# 主区域：按阶段渲染
# ============================================================

# ---------- Step 1: 数据导入预览 ----------
if st.session_state.stage == "import":
    st.markdown("### 📂 Step 1 · 数据源导入与预览")
    st.caption("上传文件后此处即时预览。确认数据无误后，点击下方按钮进入字段映射。")

    has_any = (st.session_state.last_month_df is not None
               or st.session_state.current_df is not None)
    if not has_any:
        st.info("👈 请在左侧至少上传一张 Excel 表格开始")

    # 预览上月表
    if st.session_state.last_month_df is not None:
        with st.expander(f"📊 上月主表预览（{len(st.session_state.last_month_df)} 行）", expanded=True):
            st.dataframe(st.session_state.last_month_df.head(20),
                         use_container_width=True, hide_index=True)
            st.caption(f"列名：{list(st.session_state.last_month_df.columns)}")

    # 预览本月表
    if st.session_state.current_df is not None:
        with st.expander(f"📋 本月异动表预览（{len(st.session_state.current_df)} 行）", expanded=True):
            st.dataframe(st.session_state.current_df.head(20),
                         use_container_width=True, hide_index=True)
            st.caption(f"列名：{list(st.session_state.current_df.columns)}")

    # 预览社保表
    for name, df in st.session_state.social_dfs.items():
        with st.expander(f"🪪 社保表 [{name}] 预览（{len(df)} 行）", expanded=False):
            st.dataframe(df.head(20), use_container_width=True, hide_index=True)
            st.caption(f"列名：{list(df.columns)}")

    # 进入下一步
    if has_any:
        st.markdown("---")
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("▶ 进入字段映射", type="primary", use_container_width=True):
                _set_stage("config")
                st.rerun()
        with c2:
            if st.session_state.log_lines:
                with st.expander("📜 导入日志"):
                    st.code("\n".join(st.session_state.log_lines), language="text")


# ---------- Step 2: 字段映射与异动配置 ----------
elif st.session_state.stage == "config":
    st.markdown("### ⚙️ Step 2 · 字段映射与异动配置")
    st.caption("确认主键列与比对列，系统据此计算人员异动。")

    lm = st.session_state.last_month_df
    cs = st.session_state.current_df

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 上月主表主键列")
        if lm is not None:
            default_idx = 0
            if st.session_state.id_col_lm in lm.columns:
                default_idx = list(lm.columns).index(st.session_state.id_col_lm)
            st.session_state.id_col_lm = st.selectbox(
                "选择主键列（用于人员匹配）", list(lm.columns),
                index=default_idx, key="sel_id_lm")
            st.caption(f"已选：**{st.session_state.id_col_lm}**")
        else:
            st.warning("未上传上月主表")

    with col_b:
        st.markdown("#### 本月异动表主键列")
        if cs is not None:
            default_idx = 0
            if st.session_state.id_col_cs in cs.columns:
                default_idx = list(cs.columns).index(st.session_state.id_col_cs)
            st.session_state.id_col_cs = st.selectbox(
                "选择主键列", list(cs.columns), index=default_idx, key="sel_id_cs")
            st.caption(f"已选：**{st.session_state.id_col_cs}**")
        else:
            st.info("未上传本月异动表（可跳过异动计算，直接诊断上月表）")

    st.markdown("---")
    st.markdown("#### 异动比对列")
    st.caption("选择需检测变动的字段（部门/薪资等）。仅当上月+本月都存在时生效。")

    # 比对列候选：取两表交集列名，排除主键列
    if lm is not None and cs is not None:
        common = [c for c in lm.columns if c in cs.columns
                  and c != st.session_state.id_col_lm]
        defaults = [c for c in ["姓名", "部门", "基本工资", "岗位工资"] if c in common]
        st.session_state.compare_cols = st.multiselect(
            "勾选比对列", common, default=defaults, key="sel_compare")
    else:
        st.session_state.compare_cols = []
        st.caption("⚠ 仅有一张表，无法进行异动比对，将跳过异动计算")

    st.markdown("---")
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        if st.button("← 返回导入", use_container_width=True):
            _set_stage("import")
            st.rerun()
    with c2:
        # 模式判断
        can_diff = (lm is not None and cs is not None
                    and st.session_state.id_col_lm and st.session_state.id_col_cs)
        if st.button("🔄 计算异动并进入诊断" if can_diff else "▶ 直接进入诊断",
                     type="primary", use_container_width=True):
            try:
                # 构建主表：上月表为底，无上月表则用本月表
                if lm is not None:
                    master = lm.copy()
                elif cs is not None:
                    master = cs.copy()
                else:
                    st.error("无可用数据"); st.stop()

                # 异动计算（需两表齐全）
                if can_diff:
                    engine = PayrollDBEngine()
                    diff = engine.get_personnel_changes(
                        lm, cs,
                        id_col=st.session_state.id_col_lm,
                        compare_cols=st.session_state.compare_cols or None,
                    )
                    st.session_state.diff_result = diff
                    _log(f"异动计算：新增 {len(diff['NEW_HIRES'])} / "
                         f"离职 {len(diff['DEPARTED'])} / "
                         f"异动 {len(diff['TRANSFERS'])}")
                    # 新增入职并入主表
                    hires = diff["NEW_HIRES"].copy()
                    for c in master.columns:
                        if c not in hires.columns:
                            hires[c] = None
                    hires = hires[master.columns]
                    master = pd.concat([master, hires], ignore_index=True)
                    engine.close()
                else:
                    st.session_state.diff_result = None
                    _log("仅单表，跳过异动计算，直接诊断主表")

                # 跨表关联补身份证号（主表主键列用用户选择/嗅探的列名）
                master_key_col = st.session_state.id_col_lm or (
                    _detect_id_column(master) if master is not None else "工号")
                if st.session_state.social_dfs and "身份证号" not in master.columns:
                    master["身份证号"] = ""
                for sname, sdf in st.session_state.social_dfs.items():
                    # 社保表主键列也用嗅探
                    s_key = _detect_id_column(sdf)
                    if s_key and "身份证号" in sdf.columns:
                        id_map = dict(zip(
                            sdf[s_key].apply(DataNormalizer.clean_text),
                            sdf["身份证号"].apply(DataNormalizer.clean_text)))
                        before = int((master.get("身份证号", pd.Series(dtype=str)) == "").sum())
                        master["身份证号"] = master.apply(
                            lambda r: id_map.get(
                                DataNormalizer.clean_text(r.get(master_key_col)),
                                r.get("身份证号", "")), axis=1)
                        after = int((master["身份证号"] == "").sum())
                        _log(f"社保表 [{sname}] 补齐身份证：{before - after} 条（社保主键列={s_key}，主表主键列={master_key_col}）")

                # 诊断
                validator = DataValidator()
                audited = validator.audit_employee_profiles(master)
                st.session_state.master_df = master
                st.session_state.audited_df = audited
                st.session_state.edited_df = audited.copy()
                m = validator.get_completeness_metrics(audited)
                _log(f"诊断完成：合格 {m['qualified']}/{m['total']} ({m['completeness_pct']}%)")
                _set_stage("work")
                st.rerun()
            except Exception as e:
                _log(f"❌ 处理失败：{e}")
                st.error(f"处理失败：{e}")
                with st.expander("详细堆栈"):
                    st.code(traceback.format_exc())
    with c3:
        if st.session_state.log_lines:
            with st.expander("📜 日志"):
                st.code("\n".join(st.session_state.log_lines), language="text")


# ---------- Step 3: 异动看板 + 诊断补齐 ----------
elif st.session_state.stage == "work":
    tab1, tab2 = st.tabs(["📊 异动看板", "⚠️ 诊断与在线补齐"])

    # ===== Tab 1: 异动看板 =====
    with tab1:
        diff = st.session_state.diff_result
        audited = st.session_state.audited_df
        validator = DataValidator()
        metrics = validator.get_completeness_metrics(audited)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("总计薪人数", metrics["total"])
        with c2:
            nh = len(diff["NEW_HIRES"]) if diff else 0
            st.metric("新增入职", nh, delta=f"+{nh}" if nh else None,
                      delta_color="off")
        with c3:
            dp = len(diff["DEPARTED"]) if diff else 0
            st.metric("离职删行", dp, delta=f"-{dp}" if dp else None,
                      delta_color="inverse")
        with c4:
            st.metric("数据完整度", f"{metrics['completeness_pct']}%",
                      delta=f"{metrics['qualified']}/{metrics['total']}",
                      delta_color="normal" if metrics["completeness_pct"] >= 80 else "inverse")

        st.markdown("---")

        if diff:
            sub1, sub2, sub3 = st.tabs(["🆕 新增入职", "🗑️ 离职清理", "🪪 身份证缺失预警"])
            with sub1:
                if len(diff["NEW_HIRES"]) > 0:
                    st.dataframe(diff["NEW_HIRES"], use_container_width=True, hide_index=True)
                else:
                    st.info("无新增入职人员")
            with sub2:
                if len(diff["DEPARTED"]) > 0:
                    st.dataframe(diff["DEPARTED"], use_container_width=True, hide_index=True)
                else:
                    st.info("无离职人员")
            with sub3:
                missing_id = audited[audited["诊断状态"] == DataValidator.STATUS_MISSING_ID_CARD]
                if len(missing_id) > 0:
                    st.markdown(
                        "<div class='blocker-box'>🔴 以下人员身份证号缺失，"
                        "请前往【诊断与在线补齐】Tab 修正。</div>",
                        unsafe_allow_html=True)
                    show = [c for c in ["工号", "姓名", "身份证号", "诊断状态"] if c in missing_id.columns]
                    st.dataframe(missing_id[show], use_container_width=True, hide_index=True)
                else:
                    st.markdown("<div class='success-box'>✅ 身份证号均已就绪</div>",
                                unsafe_allow_html=True)
            with st.expander("🔄 异动明细（TRANSFERS）"):
                if len(diff["TRANSFERS"]) > 0:
                    st.dataframe(diff["TRANSFERS"], use_container_width=True, hide_index=True)
                else:
                    st.info("无薪资/部门异动")
        else:
            st.info("ℹ️ 本次为单表诊断模式，未进行异动计算。如需异动对比，请返回 Step 2 上传两张表。")

    # ===== Tab 2: 诊断与在线补齐 =====
    with tab2:
        st.markdown("#### ✏️ 数据完整性诊断与在线补齐")
        st.caption("双击单元格直接修改；点击下方按钮保存并重新诊断。")

        audited = st.session_state.audited_df
        status_counts = audited["诊断状态"].value_counts().to_dict()
        cols = st.columns(len(status_counts) if status_counts else 1)
        for i, (status, cnt) in enumerate(status_counts.items()):
            with cols[i]:
                st.metric(status, cnt)

        st.markdown("---")
        metrics = DataValidator.get_completeness_metrics(audited)
        if metrics["missing_field_breakdown"]:
            st.markdown("**字段缺失分布：**")
            bd = pd.DataFrame(sorted(metrics["missing_field_breakdown"].items(),
                                     key=lambda x: -x[1]),
                              columns=["字段", "缺失人数"])
            st.bar_chart(bd.set_index("字段"))

        st.markdown("---")
        st.markdown("##### 📋 员工大表（双击编辑）")
        edited = st.data_editor(
            st.session_state.edited_df,
            num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "缺失字段清单": st.column_config.ListColumn(
                    "缺失字段清单", disabled=True),
            },
            disabled=["诊断状态"], key="audit_editor")

        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            if st.button("💾 保存并重新诊断", type="primary", use_container_width=True):
                st.session_state.edited_df = edited.copy()
                biz = [c for c in edited.columns if c not in ("诊断状态", "缺失字段清单")]
                re_audited = DataValidator().audit_employee_profiles(edited[biz])
                st.session_state.audited_df = re_audited
                st.session_state.edited_df = re_audited.copy()
                m = DataValidator.get_completeness_metrics(re_audited)
                _log(f"补齐后重新诊断：合格 {m['qualified']}/{m['total']} ({m['completeness_pct']}%)")
                st.success(f"✅ 已保存，完整度 {m['completeness_pct']}%")
                st.rerun()
        with c2:
            if st.button("↩️ 还原", use_container_width=True):
                st.session_state.edited_df = st.session_state.audited_df.copy()
                st.rerun()
        with c3:
            if st.button("▶ 进入导出", type="primary", use_container_width=True):
                _set_stage("export")
                st.rerun()

        with st.expander("📜 日志"):
            st.code("\n".join(st.session_state.log_lines), language="text")


# ---------- Step 4: 模板渲染与导出 ----------
elif st.session_state.stage == "export":
    st.markdown("### 🚀 Step 4 · 模板渲染与导出中心")

    audited = st.session_state.audited_df
    validator = DataValidator()
    metrics = validator.get_completeness_metrics(audited)
    high_risk = metrics["high_risk_count"]

    if high_risk > 0:
        st.markdown(
            f"<div class='blocker-box'>🔴 仍有 <b>{high_risk}</b> 人核心项缺失，"
            f"导出已禁用。请返回 Step 3 补齐。</div>",
            unsafe_allow_html=True)
        blockers = audited[audited["诊断状态"].str.startswith("🔴")]
        show = [c for c in ["工号", "姓名", "诊断状态", "缺失字段清单"] if c in blockers.columns]
        st.dataframe(blockers[show], use_container_width=True, hide_index=True)
        can_export = False
    else:
        st.markdown("<div class='success-box'>✅ 核心项全部通过，可导出。</div>",
                    unsafe_allow_html=True)
        can_export = True

    st.markdown("---")

    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("← 返回诊断", use_container_width=True):
            _set_stage("work")
            st.rerun()
    with c2:
        if st.button("🔄 重新导入", use_container_width=True):
            _set_stage("import")
            st.rerun()

    st.markdown("---")

    if st.button("🚀 生成本月带公式薪资汇总表",
                 type="primary", use_container_width=True, disabled=not can_export):
        with st.spinner("Openpyxl 渲染中..."):
            try:
                # 重新加载上月主模板（保留原始公式样式）
                if st.session_state.last_month_raw is not None:
                    st.session_state.last_month_raw.seek(0)
                    wb = openpyxl.load_workbook(st.session_state.last_month_raw)
                else:
                    # 无原始模板时，用诊断后的主表生成新工作簿
                    wb = openpyxl.Workbook()
                    ws_new = wb.active
                    ws_new.append(list(st.session_state.master_df.columns))
                    for _, r in st.session_state.master_df.iterrows():
                        ws_new.append([r[c] if pd.notna(r[c]) else None
                                       for c in st.session_state.master_df.columns])
                    _log("无原始模板，按主表生成新工作簿")

                ws = wb.active
                _log(f"加载工作表：{ws.title}，{ws.max_row} 行 × {ws.max_column} 列")

                # 定位工号列（按表头匹配，回退第 1 列）
                id_col_idx = 1
                for c in range(1, ws.max_column + 1):
                    h = ws.cell(row=1, column=c).value
                    if h and str(h).strip() in _ID_ALIASES:
                        id_col_idx = c
                        break
                _log(f"工号列定位：第 {id_col_idx} 列")

                # 表头映射
                header_map = {}
                for c in range(1, ws.max_column + 1):
                    h = ws.cell(row=1, column=c).value
                    if h:
                        header_map[str(h).strip()] = c

                # 删除离职行
                diff = st.session_state.diff_result
                if diff and "工号" in diff["DEPARTED"].columns:
                    departed_keys = diff["DEPARTED"]["工号"].apply(DataNormalizer.clean_text).tolist()
                    deleted = ExcelFormulaEngine.delete_departed_rows(
                        ws, key_col_idx=id_col_idx, departed_keys=departed_keys)
                    _log(f"删除离职行：{deleted} 行")
                else:
                    _log("无离职名单或无工号列，跳过删行")

                # 插入新增入职行
                if diff and len(diff["NEW_HIRES"]) > 0:
                    hires = diff["NEW_HIRES"]
                    template_row = ws.max_row
                    inserted = 0
                    for _, row in hires.iterrows():
                        emp = {}
                        for col_name, col_idx in header_map.items():
                            if col_name in row.index and pd.notna(row[col_name]):
                                emp[col_idx] = row[col_name]
                        ExcelFormulaEngine.insert_employee_row(
                            ws, template_row_idx=template_row, employee_data=emp)
                        template_row += 1
                        inserted += 1
                    _log(f"插入新增入职行：{inserted} 行")

                # 写回补齐的身份证号
                edited = st.session_state.edited_df
                if "身份证号" in header_map and "身份证号" in edited.columns:
                    idc = header_map["身份证号"]
                    key_col_name = None
                    for alias in _ID_ALIASES:
                        if alias in edited.columns:
                            key_col_name = alias
                            break
                    if key_col_name:
                        id_map = dict(zip(
                            edited[key_col_name].apply(DataNormalizer.clean_text),
                            edited["身份证号"].apply(DataNormalizer.clean_text)))
                        written = 0
                        for r in range(2, ws.max_row + 1):
                            eid = DataNormalizer.clean_text(ws.cell(row=r, column=id_col_idx).value)
                            if eid in id_map and id_map[eid]:
                                ws.cell(row=r, column=idc).value = id_map[eid]
                                written += 1
                        _log(f"写回身份证号：{written} 条")

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.session_state.export_buffer = buf.getvalue()
                st.session_state.export_filename = f"薪资汇总表_{st.session_state.salary_month}.xlsx"
                _log("✅ 导出完成")
                st.success("🎉 薪资汇总表已生成！")
            except Exception as e:
                _log(f"❌ 导出失败：{e}")
                st.error(f"导出失败：{e}")
                with st.expander("堆栈"):
                    st.code(traceback.format_exc())

    if st.session_state.export_buffer is not None:
        st.markdown("---")
        st.download_button(
            "⬇️ 下载本月薪资汇总表 (.xlsx)",
            data=st.session_state.export_buffer,
            file_name=st.session_state.export_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    with st.expander("📜 处理日志"):
        st.code("\n".join(st.session_state.log_lines), language="text")
