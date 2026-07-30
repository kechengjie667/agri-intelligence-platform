"""
农产品市场情报 — 在线交互编辑平台
启动方式: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pathlib import Path
import shutil
from io import BytesIO

# ============================================================
# 常量配置
# ============================================================
EXCEL_PATH = Path(__file__).parent / "最终数据.xlsx"
BACKUP_DIR = Path(__file__).parent / "backups"

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="农产品市场情报平台",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 样式
# ============================================================
st.markdown("""
<style>
    .stat-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem; border-radius: 10px;
        color: white; text-align: center; margin-bottom: 0.5rem;
    }
    .stat-card.green { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
    .stat-card.orange { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
    .stat-value { font-size: 2rem; font-weight: bold; }
    .stat-label { font-size: 0.85rem; opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 数据层
# ============================================================

def load_data(sheet_name: str) -> pd.DataFrame:
    """读取指定 sheet 的数据，带容错"""
    if not EXCEL_PATH.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        except Exception:
            return pd.DataFrame()
    df = df.fillna("")
    return df


def save_data(df_to_save: pd.DataFrame, sheet_name: str) -> bool:
    """保存数据到指定 sheet，写入前自动备份"""
    try:
        # 1. 自动备份
        BACKUP_DIR.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BACKUP_DIR / f"备份_{timestamp}.xlsx"
        if EXCEL_PATH.exists():
            shutil.copy2(EXCEL_PATH, backup_path)

        # 2. 清理多余备份
        backups = sorted(BACKUP_DIR.glob("备份_*.xlsx"))
        if len(backups) > 20:
            for old in backups[:-20]:
                old.unlink()

        # 3. 完整覆写（比 mode="a" 更可靠，不会产生损坏文件）
        if EXCEL_PATH.exists():
            try:
                xls = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
                all_sheets = {}
                for s in xls.sheet_names:
                    if s != sheet_name:
                        all_sheets[s] = pd.read_excel(EXCEL_PATH, sheet_name=s, engine="openpyxl")
                all_sheets[sheet_name] = df_to_save
            except Exception:
                all_sheets = {sheet_name: df_to_save}
        else:
            all_sheets = {sheet_name: df_to_save}

        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            for s, sdf in all_sheets.items():
                sdf.fillna("").to_excel(writer, sheet_name=s, index=False)

        return True
    except Exception as e:
        st.sidebar.error(f"保存失败: {e}")
        return False


def get_available_sheets() -> list:
    """获取所有 sheet 名称"""
    if not EXCEL_PATH.exists():
        return ["Sheet1"]
    try:
        return pd.ExcelFile(EXCEL_PATH, engine="openpyxl").sheet_names
    except Exception:
        try:
            return pd.ExcelFile(EXCEL_PATH).sheet_names
        except Exception:
            return ["Sheet1"]


# ============================================================
# 日志管理
# ============================================================
if "log" not in st.session_state:
    st.session_state.log = []


def add_log(action: str, detail: str = ""):
    now = datetime.now().strftime("%H:%M:%S")
    st.session_state.log.insert(0, {"time": now, "action": action, "detail": detail})
    st.session_state.log = st.session_state.log[:50]


# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.title("🌾 农产品情报平台")

    available_sheets = get_available_sheets()
    if "selected_sheet" not in st.session_state:
        st.session_state.selected_sheet = available_sheets[0] if available_sheets else "Sheet1"

    selected_sheet = st.selectbox(
        "📄 选择数据表",
        available_sheets,
        index=available_sheets.index(st.session_state.selected_sheet)
        if st.session_state.selected_sheet in available_sheets else 0,
        key="sheet_selector",
    )
    st.session_state.selected_sheet = selected_sheet

    st.divider()

    df = load_data(selected_sheet)
    total_rows = len(df)
    total_cols = len(df.columns)
    unique_cats = df["类型"].nunique() if "类型" in df.columns and total_rows > 0 else 0

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'<div class="stat-card"><div class="stat-value">{total_rows}</div><div class="stat-label">总条目</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="stat-card green"><div class="stat-value">{total_cols}</div><div class="stat-label">字段数</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="stat-card orange"><div class="stat-value">{unique_cats}</div><div class="stat-label">类型数</div></div>', unsafe_allow_html=True)

    st.divider()

    st.subheader("📋 操作日志")
    if not st.session_state.log:
        st.caption("暂无操作记录")
    else:
        for entry in st.session_state.log[:10]:
            icon = {"新增": "➕", "修改": "✏️", "删除": "🗑️", "保存": "💾", "导出": "⬇️"}.get(entry["action"], "📌")
            st.caption(f'{icon} [{entry["time"]}] {entry["action"]} - {entry["detail"]}')
        if len(st.session_state.log) > 10:
            st.caption(f"... 共 {len(st.session_state.log)} 条记录")

    st.divider()

    st.subheader("🔒 备份管理")
    if BACKUP_DIR.exists():
        backups = sorted(BACKUP_DIR.glob("备份_*.xlsx"), reverse=True)
        st.caption(f"共 {len(backups)} 个备份（保留最近20个）")
        if backups:
            with st.expander("查看备份列表"):
                for bp in backups[:10]:
                    st.caption(f"📁 {bp.name}")
    else:
        st.caption("暂无备份")

# ============================================================
# 主区域
# ============================================================
st.title("📋 农产品市场情报 — 在线编辑平台")
st.caption(f"当前数据表: **{selected_sheet}** | 数据文件: `{EXCEL_PATH.name}`")

tab1, tab2 = st.tabs(["📋 数据编辑", "📊 数据可视化"])

# -------- Tab1: 数据编辑 --------
with tab1:
    if df.empty:
        st.info("该数据表为空，点击「新增行」开始添加数据。")
        edited_df = df
    else:
        # 核心：加一个「删除勾选」复选框列，勾选后点按钮真正删除
        if "___select___" not in df.columns:
            df = df.copy()
            df.insert(0, "___select___", False)

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "___select___": st.column_config.CheckboxColumn(
                    "删除勾选",
                    help="勾选要删除的行，然后点下方「删除勾选行」按钮",
                    width="small",
                ),
                "类型": st.column_config.TextColumn("类型", width="small"),
                "标题": st.column_config.TextColumn("标题", width="medium"),
                "简介": st.column_config.TextColumn("简介", width="large"),
            },
            height=400,
            key=f"editor_{selected_sheet}",
        )

    # 按钮行
    btn1, btn2, btn3, btn4 = st.columns([1, 1, 1, 5])

    with btn1:
        if st.button("➕ 新增行", use_container_width=True):
            current = edited_df if edited_df is not None else df
            new_row = {col: "" for col in current.columns if col != "___select___"}
            new_row["___select___"] = False
            current = pd.concat([current, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state["pending_data"] = current
            add_log("新增", "添加了1行")
            st.rerun()

    with btn2:
        delete_clicked = st.button("🗑️ 删除勾选行", use_container_width=True)

    with btn3:
        save_clicked = st.button("💾 保存修改", use_container_width=True, type="primary")

    # --- 删除勾选行 ---
    if delete_clicked:
        if edited_df is not None and "___select___" in edited_df.columns:
            before = len(edited_df)
            edited_df = edited_df[edited_df["___select___"] != True].copy()
            after = len(edited_df)
            deleted = before - after
            if deleted > 0:
                add_log("删除", f"删除了 {deleted} 行")
                st.session_state["pending_data"] = edited_df
                st.success(f"已删除 {deleted} 行, 剩余 {after} 行")
                st.rerun()
            else:
                st.warning("没有勾选任何行，请先在「删除勾选」列打勾")
        else:
            st.warning("没有可删除的数据")

    # --- 保存 ---
    if save_clicked:
        data_to_save = st.session_state.get("pending_data", edited_df)

        if data_to_save is not None:
            if "___select___" in data_to_save.columns:
                data_to_save = data_to_save.drop(columns=["___select___"])

            # 去掉完全空行
            data_to_save = data_to_save[data_to_save.astype(str).apply(
                lambda row: row.str.strip().str.len().sum() > 0, axis=1
            )]

            df_clean = df.drop(columns=["___select___"]) if "___select___" in df.columns else df

            if data_to_save.reset_index(drop=True).equals(df_clean.reset_index(drop=True)):
                st.info("数据没有变化，无需保存")
            else:
                if save_data(data_to_save, selected_sheet):
                    add_log("保存", f"保存到 {selected_sheet} 成功")
                    st.session_state.pop("pending_data", None)
                    st.success(f"已保存到 `{EXCEL_PATH.name}` → `{selected_sheet}`")
                    st.rerun()
                else:
                    st.error("保存失败，请检查文件是否被占用")
        else:
            st.info("没有数据可保存")

    # --- 导出 ---
    st.divider()
    exp1, exp2, _ = st.columns([2, 2, 8])
    export_df = st.session_state.get("pending_data", edited_df)
    if export_df is not None and "___select___" in export_df.columns:
        export_df = export_df.drop(columns=["___select___"])

    with exp1:
        if export_df is not None:
            csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ 导出 CSV", csv_data,
                f"市场情报_{selected_sheet}_{datetime.now():%Y%m%d}.csv",
                "text/csv", use_container_width=True,
            )
    with exp2:
        if export_df is not None:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                export_df.to_excel(writer, sheet_name=selected_sheet, index=False)
            st.download_button(
                "⬇️ 导出 Excel", buffer.getvalue(),
                f"市场情报_{selected_sheet}_{datetime.now():%Y%m%d}.xlsx",
                use_container_width=True,
            )

# -------- Tab2: 可视化 --------
with tab2:
    st.subheader("📊 数据可视化分析")

    if df.empty or len(df) == 0:
        st.info("当前数据表为空，请先添加数据。")
    else:
        df_viz = df.drop(columns=["___select___"]) if "___select___" in df.columns else df

        col1, col2 = st.columns(2)

        with col1:
            if "类型" in df_viz.columns:
                counts = df_viz["类型"].value_counts().reset_index()
                counts.columns = ["类型", "数量"]
                fig = px.pie(counts, names="类型", values="数量",
                             title="类型分布", hole=0.4,
                             color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_traces(textinfo="label+value+percent")
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)

        with col2:
            if "类型" in df_viz.columns:
                counts = df_viz["类型"].value_counts().reset_index()
                counts.columns = ["类型", "数量"]
                fig = px.bar(counts, x="类型", y="数量", title="各类型条目数",
                             color="类型", text="数量",
                             color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_traces(textposition="outside")
                fig.update_layout(height=400, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        st.divider()
        if "简介" in df_viz.columns:
            df_viz["简介字数"] = df_viz["简介"].apply(lambda x: len(str(x)))
            df_viz["标题字数"] = df_viz["标题"].apply(lambda x: len(str(x)))

            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(df_viz, x=df_viz.index, y="简介字数",
                             title="各条目简介字数", text="简介字数",
                             color="类型" if "类型" in df_viz.columns else None,
                             color_discrete_sequence=px.colors.qualitative.Set2,
                             labels={"index": "条目序号"})
                fig.update_traces(textposition="outside")
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(df_viz, x=df_viz.index, y="标题字数",
                             title="各条目标题字数", text="标题字数",
                             color="类型" if "类型" in df_viz.columns else None,
                             color_discrete_sequence=px.colors.qualitative.Set2,
                             labels={"index": "条目序号"})
                fig.update_traces(textposition="outside")
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)

# ============================================================
# 底部
# ============================================================
st.divider()
st.caption("💡 提示：双击单元格编辑 | 勾选「删除勾选」→ 点「删除勾选行」 | 修改后点「保存修改」 | 保存时自动备份")
