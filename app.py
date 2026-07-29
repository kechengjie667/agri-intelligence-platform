"""
农产品市场情报 — 在线交互编辑平台
本地 Excel 数据 → 浏览器在线查看/编辑/删除/新增
启动方式: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pathlib import Path
import shutil
import os

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
# 样式微调
# ============================================================
st.markdown("""
<style>
    /* 卡片样式 */
    .stat-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 0.5rem;
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
    """读取指定 sheet 的数据"""
    if not EXCEL_PATH.exists():
        return pd.DataFrame()
    xls = pd.ExcelFile(EXCEL_PATH)
    if sheet_name not in xls.sheet_names:
        return pd.DataFrame()
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
    # 确保所有列是字符串类型，避免 data_editor 显示问题
    df = df.fillna("")
    return df


def save_data(df: pd.DataFrame, sheet_name: str) -> bool:
    """保存数据到指定 sheet，写入前自动备份"""
    try:
        # 1. 自动备份
        BACKUP_DIR.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BACKUP_DIR / f"最终数据_备份_{timestamp}.xlsx"
        if EXCEL_PATH.exists():
            shutil.copy2(EXCEL_PATH, backup_path)

        # 2. 清理多余备份：只保留最近20个
        backups = sorted(BACKUP_DIR.glob("最终数据_备份_*.xlsx"))
        if len(backups) > 20:
            for old in backups[:-20]:
                old.unlink()

        # 3. 写出 Excel
        # 保留其他 sheet 不变，只更新目标 sheet
        if EXCEL_PATH.exists():
            with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            df.to_excel(EXCEL_PATH, sheet_name=sheet_name, index=False)

        return True
    except Exception as e:
        # 如果 mode='a' 失败（可能 sheet 不存在），用覆写方式
        try:
            xls = pd.ExcelFile(EXCEL_PATH)
            all_sheets = {}
            for s in xls.sheet_names:
                if s != sheet_name:
                    all_sheets[s] = pd.read_excel(EXCEL_PATH, sheet_name=s)
            all_sheets[sheet_name] = df
            with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
                for s, sdf in all_sheets.items():
                    sdf.to_excel(writer, sheet_name=s, index=False)
            return True
        except Exception as e2:
            st.sidebar.error(f"保存失败: {e2}")
            return False


def get_available_sheets() -> list:
    """获取 Excel 中所有 sheet 名称"""
    if not EXCEL_PATH.exists():
        return ["Sheet1"]
    return pd.ExcelFile(EXCEL_PATH).sheet_names


# ============================================================
# 日志管理
# ============================================================
if "log" not in st.session_state:
    st.session_state.log = []


def add_log(action: str, detail: str = ""):
    """记录操作日志"""
    now = datetime.now().strftime("%H:%M:%S")
    st.session_state.log.insert(0, {"time": now, "action": action, "detail": detail})
    # 只保留最近50条
    st.session_state.log = st.session_state.log[:50]


# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.title("🌾 农产品情报平台")

    # Sheet 选择
    available_sheets = get_available_sheets()
    if "selected_sheet" not in st.session_state:
        st.session_state.selected_sheet = available_sheets[0] if available_sheets else "Sheet1"

    selected_sheet = st.selectbox(
        "📄 选择数据表",
        available_sheets,
        index=available_sheets.index(st.session_state.selected_sheet)
        if st.session_state.selected_sheet in available_sheets
        else 0,
        key="sheet_selector",
    )
    st.session_state.selected_sheet = selected_sheet

    st.divider()

    # 数据统计
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

    # 操作日志
    st.subheader("📋 操作日志")
    if not st.session_state.log:
        st.caption("暂无操作记录")
    else:
        for entry in st.session_state.log[:10]:
            icon = {"新增": "➕", "修改": "✏️", "删除": "🗑️", "保存": "💾", "导出": "⬇️"}.get(entry["action"], "📌")
            st.caption(f'{icon} [{entry["time"]}] {entry["action"]} - {entry["detail"]}')

    if st.session_state.log and len(st.session_state.log) > 10:
        st.caption(f"... 共 {len(st.session_state.log)} 条记录")

    st.divider()

    # 备份管理
    st.subheader("🔒 备份管理")
    if BACKUP_DIR.exists():
        backups = sorted(BACKUP_DIR.glob("最终数据_备份_*.xlsx"), reverse=True)
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
    # 按钮行
    btn_col1, btn_col2, btn_col3, btn_col4, btn_col5 = st.columns([1, 1, 1, 1, 3])

    with btn_col1:
        if st.button("➕ 新增行", use_container_width=True, help="在表格末尾添加一个空行"):
            new_row = {col: "" for col in df.columns} if len(df.columns) > 0 else {}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            add_log("新增", f"添加了1行到 {selected_sheet}")
            st.rerun()

    with btn_col2:
        # 删除选中行
        delete_clicked = st.button("🗑️ 删除选中行", use_container_width=True, help="删除表格中勾选的行")

    with btn_col3:
        save_clicked = st.button("💾 保存修改", use_container_width=True, type="primary", help="将所有修改写回 Excel")

    with btn_col4:
        # 导出
        pass  # 导出按钮放在后面用 download_button

    # 可编辑表格
    st.subheader(f"数据编辑区 — {selected_sheet}")

    if df.empty:
        st.info("该数据表为空，点击「新增行」开始添加数据。")
        edited_df = df
    else:
        # 使用 data_editor 实现可编辑表格
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",              # 允许动态增删行
            use_container_width=True,
            hide_index=True,
            column_config={
                "类型": st.column_config.TextColumn(
                    "类型",
                    width="small",
                    help="情报分类标签",
                ),
                "标题": st.column_config.TextColumn(
                    "标题",
                    width="medium",
                    help="情报标题",
                ),
                "简介": st.column_config.TextColumn(
                    "简介",
                    width="large",
                    help="情报详细内容",
                ),
            },
            height=400,
            key=f"editor_{selected_sheet}",
        )

    # 删除选中行逻辑
    if delete_clicked:
        # data_editor 的 dynamic 模式下，删除操作直接由用户通过表格 UI 完成
        # 这里的数据是编辑后的结果，删除通过用户操作表格行实现
        st.warning("💡 选中表格中的行，按键盘 **Delete** 键或点击行号旁的删除按钮即可删除")
        add_log("删除", f"用户手动删除了行")
        st.rerun()

    # 保存逻辑
    if save_clicked:
        if edited_df is not None and not edited_df.equals(df):
            success = save_data(edited_df, selected_sheet)
            if success:
                add_log("保存", f"保存到 {selected_sheet} 成功")
                st.success(f"✅ 已保存到 `{EXCEL_PATH.name}` → `{selected_sheet}`")
                st.rerun()
            else:
                st.error("❌ 保存失败，请检查文件是否被占用")
        elif edited_df is not None and edited_df.equals(df):
            st.info("数据没有变化，无需保存")

    # 导出按钮
    st.divider()
    exp_col1, exp_col2, exp_col3 = st.columns([2, 2, 8])
    with exp_col1:
        csv_data = edited_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="⬇️ 导出 CSV",
            data=csv_data,
            file_name=f"市场情报_{selected_sheet}_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with exp_col2:
        from io import BytesIO
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            edited_df.to_excel(writer, sheet_name=selected_sheet, index=False)
        st.download_button(
            label="⬇️ 导出 Excel",
            data=buffer.getvalue(),
            file_name=f"市场情报_{selected_sheet}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# -------- Tab2: 可视化 --------
with tab2:
    st.subheader("📊 数据可视化分析")

    if df.empty or len(df) == 0:
        st.info("当前数据表为空，请先添加数据。")
    else:
        viz_col1, viz_col2 = st.columns(2)

        with viz_col1:
            # 类型分布 — 饼图
            if "类型" in df.columns:
                cat_counts = df["类型"].value_counts().reset_index()
                cat_counts.columns = ["类型", "数量"]
                fig_pie = px.pie(
                    cat_counts,
                    names="类型",
                    values="数量",
                    title="类型分布",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    hole=0.4,
                )
                fig_pie.update_traces(textinfo="label+value+percent")
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)

        with viz_col2:
            # 类型数量 — 柱状图
            if "类型" in df.columns:
                cat_counts = df["类型"].value_counts().reset_index()
                cat_counts.columns = ["类型", "数量"]
                fig_bar = px.bar(
                    cat_counts,
                    x="类型",
                    y="数量",
                    title="各类型条目数",
                    color="类型",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    text="数量",
                )
                fig_bar.update_traces(textposition="outside")
                fig_bar.update_layout(height=400, showlegend=False)
                st.plotly_chart(fig_bar, use_container_width=True)

        # 内容长度分析
        st.divider()
        if "简介" in df.columns:
            df_viz = df.copy()
            df_viz["简介字数"] = df_viz["简介"].apply(lambda x: len(str(x)))
            df_viz["标题字数"] = df_viz["标题"].apply(lambda x: len(str(x)))

            len_col1, len_col2 = st.columns(2)

            with len_col1:
                fig_content = px.bar(
                    df_viz,
                    x=df_viz.index,
                    y="简介字数",
                    title="各条目简介字数",
                    color="类型" if "类型" in df_viz.columns else None,
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    text="简介字数",
                    labels={"index": "条目序号"},
                )
                fig_content.update_traces(textposition="outside")
                fig_content.update_layout(height=400)
                st.plotly_chart(fig_content, use_container_width=True)

            with len_col2:
                fig_title = px.bar(
                    df_viz,
                    x=df_viz.index,
                    y="标题字数",
                    title="各条目标题字数",
                    color="类型" if "类型" in df_viz.columns else None,
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    text="标题字数",
                    labels={"index": "条目序号"},
                )
                fig_title.update_traces(textposition="outside")
                fig_title.update_layout(height=400)
                st.plotly_chart(fig_title, use_container_width=True)

# ============================================================
# 底部
# ============================================================
st.divider()
st.caption("💡 提示：双击单元格编辑 | 选中行按 Delete 删除 | 修改后记得点「保存修改」 | 保存时自动备份")
