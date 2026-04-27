import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import date
import hashlib

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="报销数据仪表板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义CSS美化 ====================
st.markdown("""
<style>
    /* 主标题样式 */
    .main-title {
        font-size: 2.5rem;
        font-weight: 600;
        color: #1E3A5F;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        text-align: center;
        color: #4A627A;
        margin-bottom: 2rem;
    }
    /* 卡片样式 */
    .css-1r6slb0, .stMetric {
        background-color: #FFFFFF;
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        transition: transform 0.2s;
    }
    .css-1r6slb0:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    /* 侧边栏样式 */
    .css-1d391kg {
        background-color: #F8FAFC;
    }
    /* 按钮样式 */
    .stButton > button {
        background-color: #1E3A5F;
        color: white;
        border-radius: 20px;
        border: none;
        padding: 0.5rem 1.2rem;
        font-weight: 500;
        transition: all 0.3s;
    }
    .stButton > button:hover {
        background-color: #2C5282;
        box-shadow: 0 2px 6px rgba(0,0,0,0.2);
    }
    /* 表格样式 */
    .dataframe {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid #E2E8F0;
    }
    /* 扩展器样式 */
    .streamlit-expanderHeader {
        background-color: #F1F5F9;
        border-radius: 12px;
        font-weight: 500;
    }
    /* 信息框样式 */
    .stAlert {
        border-radius: 12px;
        border-left-width: 6px;
    }
    /* 选中项目提示 */
    .project-highlight {
        background-color: #EFF6FF;
        border-left: 4px solid #1E3A5F;
        padding: 0.75rem 1rem;
        border-radius: 12px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown('<div class="main-title">📊 报销数据仪表板</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">费用监控 · 异常预警 · 会议分析</div>', unsafe_allow_html=True)
st.markdown("---")

# ==================== 常量 ====================
PROJECTS = ["Novartis", "Dove", "Gilead CSO", "Gilead", "Menarini", "Pfizer", "Hanson", "3M", "Servier", "药房"]
COST_COLS = ["火车/高铁", "打车", "公交", "其他", "酒店", "差补"]

# ==================== 缓存函数 ====================
@st.cache_data
def generate_template():
    return pd.DataFrame({
        "项目": [PROJECTS[0], PROJECTS[1]],
        "子项目": ["子项目1", "子项目2"],
        "申请人": ["张三", "李四"],
        "会议号": ["M001", "M002"],
        "费用日期": ["2026-04-01", "2026-04-15"],
        "出差地点": ["北京", "上海"],
        "费用发生地点": ["北京", "杭州"],
        "火车/高铁": [500, 0],
        "打车": [100, 250],
        "公交": [50, 30],
        "其他": [0, 50],
        "酒店": [600, 800],
        "差补": [200, 300],
        "总金额": [1450, 1430],
        "备注栏": ["正常报销", "打车超标"],
        "审批状态": ["已通过", "审批中"]
    })

@st.cache_data
def load_data(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file, engine='openpyxl')

@st.cache_data
def process_data(df, start_date, end_date, selected_project_filter, selected_applicant, selected_status):
    """数据处理和筛选（缓存）"""
    df_processed = df.copy()
    df_processed["费用日期"] = pd.to_datetime(df_processed["费用日期"], errors='coerce')
    df_processed = df_processed.dropna(subset=["费用日期"])
    for col in COST_COLS:
        df_processed[col] = pd.to_numeric(df_processed[col], errors='coerce').fillna(0)
    df_processed["总金额"] = df_processed[COST_COLS].sum(axis=1)
    if "备注栏" not in df_processed.columns:
        df_processed["备注栏"] = ""
    df_processed["项目"] = df_processed["项目"].apply(lambda x: x if x in PROJECTS else "其他")
    
    # 筛选
    df_processed = df_processed[(df_processed["费用日期"] >= pd.to_datetime(start_date)) &
                                (df_processed["费用日期"] <= pd.to_datetime(end_date))]
    if selected_project_filter != "全部":
        df_processed = df_processed[df_processed["项目"] == selected_project_filter]
    if selected_applicant != "全部":
        df_processed = df_processed[df_processed["申请人"] == selected_applicant]
    if selected_status != "全部":
        df_processed = df_processed[df_processed["审批状态"] == selected_status]
    return df_processed

@st.cache_data
def compute_meeting_stats(df):
    return df.groupby("会议号").agg(
        会议总金额=("总金额", "sum"),
        记录条数=("总金额", "count"),
        平均单次金额=("总金额", "mean")
    ).reset_index()

@st.cache_data
def compute_exception_flags(df):
    """计算异常标志（缓存）"""
    df_copy = df.copy()
    # 当日酒店
    df_copy["费用日期_仅日期"] = df_copy["费用日期"].dt.date
    daily_hotel = df_copy.groupby(["申请人", "费用日期_仅日期"])["酒店"].sum().reset_index()
    daily_hotel.rename(columns={"酒店": "当日酒店总额"}, inplace=True)
    df_copy = df_copy.merge(daily_hotel, on=["申请人", "费用日期_仅日期"], how="left")
    df_copy["当日酒店超标"] = df_copy["当日酒店总额"] > 400
    # 月酒店
    df_copy["年月"] = df_copy["费用日期"].dt.to_period("M").astype(str)
    monthly_hotel = df_copy.groupby(["申请人", "年月"])["酒店"].sum().reset_index()
    monthly_hotel.rename(columns={"酒店": "当月酒店总额"}, inplace=True)
    df_copy = df_copy.merge(monthly_hotel, on=["申请人", "年月"], how="left")
    df_copy["月酒店超标"] = df_copy["当月酒店总额"] > 13000
    df_copy["地点不一致"] = df_copy["出差地点"] != df_copy["费用发生地点"]
    df_copy["打车超标"] = df_copy["打车"] > 200
    
    def mark_exception(row):
        reasons = []
        if row["地点不一致"]: reasons.append("地点不一致")
        if row["当日酒店超标"]: reasons.append("当日酒店>400")
        if row["月酒店超标"]: reasons.append("月酒店>13000")
        return ", ".join(reasons) if reasons else "正常"
    df_copy["异常类型"] = df_copy.apply(mark_exception, axis=1)
    return df_copy

@st.cache_data
def create_pie_chart(project_totals):
    fig = px.pie(project_totals, names="项目", values="总金额", 
                 title="项目总费用占比", hole=0.3,
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_traces(textposition='inside', textinfo='percent+label', 
                      pull=[0.05]*len(project_totals))
    fig.update_layout(
        title_font_size=18,
        font=dict(size=12),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
    )
    return fig

@st.cache_data
def create_monthly_trend(monthly):
    fig = px.line(monthly, x="年月", y="总金额", markers=True, 
                  title="月度报销总额趋势", color_discrete_sequence=["#1E3A5F"])
    fig.update_traces(marker=dict(size=8), line=dict(width=2))
    fig.update_layout(title_font_size=18, xaxis_title="月份", yaxis_title="报销金额(元)")
    return fig

@st.cache_data
def create_cost_pie(cost_sum):
    fig = px.pie(values=cost_sum.values, names=cost_sum.index, 
                 title="费用类别构成", color_discrete_sequence=px.colors.sequential.Blues_r)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

# ==================== 导出函数 ====================
def export_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="报销数据")
    buffer.seek(0)
    return buffer

# ==================== 侧边栏 ====================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/receipt.png", width=60)
    st.markdown("## 📂 数据导入")
    template = generate_template()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        template.to_excel(writer, index=False, sheet_name="报销模板")
    st.download_button(
        label="📥 下载中文模板",
        data=buffer.getvalue(),
        file_name="报销模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("---")
    uploaded_file = st.file_uploader("上传报销数据 (Excel / CSV)", type=["xlsx", "csv"])
    
    st.markdown("---")
    st.markdown("## 🔍 全局筛选")
    if "df_original" in locals() and not df_original.empty:
        min_date = df_original["费用日期"].min().date()
        max_date = df_original["费用日期"].max().date()
    else:
        min_date = max_date = date.today()
    start_date = st.date_input("起始日期", value=min_date)
    end_date = st.date_input("结束日期", value=max_date)
    
    project_list = ["全部"]
    applicant_list = ["全部"]
    status_list = ["全部"]
    if "df_original" in locals():
        project_list += sorted([p for p in PROJECTS if p in df_original["项目"].unique()])
        applicant_list += sorted(df_original["申请人"].dropna().unique())
        status_list += sorted(df_original["审批状态"].dropna().unique())
    
    selected_project_filter = st.selectbox("🏷️ 项目", project_list)
    selected_applicant = st.selectbox("👤 申请人", applicant_list)
    selected_status = st.selectbox("📌 审批状态", status_list)

# ==================== 数据加载 ====================
if uploaded_file is not None:
    try:
        df_original = load_data(uploaded_file)
        st.success("✅ 数据加载成功！", icon="✅")
    except Exception as e:
        st.error(f"文件读取失败: {e}")
        st.stop()
else:
    st.info("👈 请下载模板填写并上传，当前为演示数据")
    df_original = generate_template()
    st.warning("⚠️ 演示数据仅作预览，请上传真实数据", icon="⚠️")

# 必要字段校验
required_cols = ["项目", "子项目", "申请人", "会议号", "费用日期", "出差地点", "费用发生地点",
                 "火车/高铁", "打车", "公交", "其他", "酒店", "差补", "审批状态"]
missing = [c for c in required_cols if c not in df_original.columns]
if missing:
    st.error(f"❌ 缺少必要字段：{missing}，请使用标准模板（含【会议号】）")
    st.stop()

# 应用处理与筛选（缓存）
df_filtered = process_data(df_original, start_date, end_date, 
                           selected_project_filter, selected_applicant, selected_status)
if df_filtered.empty:
    st.warning("⚠️ 当前筛选条件下无数据，请调整筛选条件")
    st.stop()

# 会议统计
meeting_stats = compute_meeting_stats(df_filtered)
overall_avg_meeting_amount = meeting_stats["平均单次金额"].mean() if not meeting_stats.empty else 0

# 异常计算
df_with_exception = compute_exception_flags(df_filtered)

# ==================== 饼图联动 ====================
if "selected_project" not in st.session_state:
    st.session_state.selected_project = None

project_totals = df_filtered.groupby("项目")["总金额"].sum().reset_index()
project_totals = project_totals[project_totals["项目"] != "其他"]
if not project_totals.empty:
    pie_chart = create_pie_chart(project_totals)
    selected_points = st.plotly_chart(pie_chart, selection_mode="points", on_select="rerun", key="pie", use_container_width=True)
    if selected_points and selected_points.selection and len(selected_points.selection.points) > 0:
        clicked_point = selected_points.selection.points[0]
        clicked_project = getattr(clicked_point, 'label', getattr(clicked_point, 'x', None))
        if clicked_project and clicked_project != st.session_state.selected_project:
            st.session_state.selected_project = clicked_project

col_clear, _ = st.columns([1, 5])
with col_clear:
    if st.button("🗑️ 清除项目联动"):
        st.session_state.selected_project = None
        st.rerun()

# 根据联动筛选数据
df_display = df_with_exception.copy()
if st.session_state.selected_project is not None:
    df_display = df_display[df_display["项目"] == st.session_state.selected_project]
    st.markdown(f'<div class="project-highlight">📌 当前仅显示项目：<strong>{st.session_state.selected_project}</strong> 的数据（点击饼图扇形联动）</div>', unsafe_allow_html=True)
    # 重新计算会议统计
    meeting_stats = compute_meeting_stats(df_display)
    overall_avg_meeting_amount = meeting_stats["平均单次金额"].mean() if not meeting_stats.empty else 0

# ==================== KPI卡片 ====================
st.markdown("## 📈 核心指标")
col1, col2, col3, col4, col5 = st.columns(5)
total_amt = df_display["总金额"].sum()
avg_amt = df_display["总金额"].mean()
abnormal_taxi = (df_display["打车"] > 200).sum()
approved_amt = df_display[df_display["审批状态"] == "已通过"]["总金额"].sum()
approved_ratio = approved_amt / total_amt * 100 if total_amt > 0 else 0

col1.metric("💰 报销总额", f"{total_amt:,.0f} 元")
col2.metric("📈 单均报销", f"{avg_amt:,.0f} 元")
col3.metric("🚕 打车>200", abnormal_taxi, delta="需关注" if abnormal_taxi>0 else None)
col4.metric("✅ 通过占比", f"{approved_ratio:.1f}%")
col5.metric("🎤 会议平均金额", f"{overall_avg_meeting_amount:,.0f} 元")
st.markdown("---")

# ==================== 会议分析 ====================
st.markdown("## 🎤 会议费用分析")
col_meeting1, col_meeting2 = st.columns(2)
with col_meeting1:
    st.subheader("🏆 会议总金额 TOP10")
    top_meetings = meeting_stats.nlargest(10, "会议总金额")[["会议号", "会议总金额", "记录条数", "平均单次金额"]]
    st.dataframe(top_meetings, use_container_width=True)
with col_meeting2:
    st.subheader("📈 会议平均金额分布")
    if not meeting_stats.empty:
        fig_meeting_avg = px.histogram(meeting_stats, x="平均单次金额", nbins=15, 
                                       title="各会议平均单次报销金额分布",
                                       color_discrete_sequence=["#2C5282"])
        fig_meeting_avg.update_layout(bargap=0.1)
        st.plotly_chart(fig_meeting_avg, use_container_width=True)
    else:
        st.info("无会议数据")
st.markdown("### 📋 会议明细汇总")
st.dataframe(meeting_stats.sort_values("会议总金额", ascending=False), use_container_width=True)
st.markdown("---")

# ==================== 月度趋势 ====================
st.markdown("## 📅 费用趋势")
df_display["年月"] = df_display["费用日期"].dt.to_period("M").astype(str)
monthly = df_display.groupby("年月")["总金额"].sum().reset_index()
if not monthly.empty:
    fig_trend = create_monthly_trend(monthly)
    st.plotly_chart(fig_trend, use_container_width=True)

# ==================== 费用构成 ====================
st.markdown("## 🧾 费用类别构成")
cost_sum = df_display[COST_COLS].sum()
if cost_sum.sum() > 0:
    fig_cost = create_cost_pie(cost_sum)
    st.plotly_chart(fig_cost, use_container_width=True)

# ==================== 项目对比 ====================
st.markdown("## 📊 项目费用对比")
if st.session_state.selected_project:
    compare_df = df_filtered.groupby("项目")["总金额"].sum().reset_index()
    fig_compare = px.bar(compare_df, x="项目", y="总金额", color="项目",
                         title=f"各项目总费用（当前选中：{st.session_state.selected_project}）",
                         color_discrete_map={st.session_state.selected_project: "#E53E3E"})
    fig_compare.update_layout(showlegend=False, xaxis_title="项目", yaxis_title="报销金额(元)")
    st.plotly_chart(fig_compare, use_container_width=True)
else:
    proj_sum = df_display.groupby("项目")["总金额"].sum().sort_values(ascending=False).reset_index()
    fig_proj = px.bar(proj_sum, x="项目", y="总金额", title="各项目报销总额", color="项目")
    fig_proj.update_layout(showlegend=False)
    st.plotly_chart(fig_proj, use_container_width=True)

# ==================== 异常监控（折叠） ====================
abnormal_count_total = (df_display["异常类型"] != "正常").sum()
st.info(f"📋 当前共有 **{abnormal_count_total}** 条异常单据（地点不一致/酒店超标），点击下方展开查看详情")
with st.expander("⚠️ 查看异常监控详情（点击展开）"):
    st.markdown("**异常规则说明：**")
    col_rule1, col_rule2, col_rule3 = st.columns(3)
    col_rule1.markdown("- 🔴 出差地点 ≠ 费用发生地点")
    col_rule2.markdown("- 🔴 同人同日酒店 > 400元")
    col_rule3.markdown("- 🔴 同人同月酒店 > 13000元")
    st.markdown("- 🟡 单笔打车 > 200元（黄色高亮提醒）")
    
    abnormal_df = df_display[df_display["异常类型"] != "正常"]
    if not abnormal_df.empty:
        st.subheader("📋 异常单据明细")
        show_cols = ["项目", "子项目", "申请人", "会议号", "费用日期", "出差地点", "费用发生地点",
                     "总金额", "酒店", "打车", "当日酒店总额", "当月酒店总额", "异常类型", "审批状态", "备注栏"]
        st.dataframe(abnormal_df[show_cols], use_container_width=True)
    else:
        st.success("✅ 未发现任何异常单据")
    
    st.markdown("### 🟡 单笔打车费用超过200元明细")
    taxi_over_df = df_display[df_display["打车超标"]]
    if not taxi_over_df.empty:
        st.dataframe(taxi_over_df[["申请人", "会议号", "费用日期", "出差地点", "打车", "审批状态", "备注栏"]], use_container_width=True)
    else:
        st.info("暂无打车超200元的单据")
st.markdown("---")

# ==================== 审批分析 ====================
st.markdown("## 📋 审批情况分析")
col_a, col_b = st.columns(2)
status_cnt = df_display["审批状态"].value_counts().reset_index()
status_cnt.columns = ["审批状态", "单据数"]
with col_a:
    fig_status = px.bar(status_cnt, x="审批状态", y="单据数", color="审批状态", title="各状态单据数量")
    fig_status.update_layout(showlegend=False)
    st.plotly_chart(fig_status, use_container_width=True)
status_amt = df_display.groupby("审批状态")["总金额"].sum().reset_index()
with col_b:
    fig_amt = px.pie(status_amt, values="总金额", names="审批状态", title="各状态金额占比")
    st.plotly_chart(fig_amt, use_container_width=True)

st.markdown("### 👥 申请人审批状态统计")
cross_tab = pd.crosstab(df_display["申请人"], df_display["审批状态"])
if not cross_tab.empty:
    st.dataframe(cross_tab, use_container_width=True)

# ==================== 员工排行 ====================
st.markdown("## 🏆 申请人报销TOP10")
top_applicants = df_display.groupby("申请人")["总金额"].sum().sort_values(ascending=False).head(10).reset_index()
if not top_applicants.empty:
    fig_top = px.bar(top_applicants, x="申请人", y="总金额", title="申请人报销总额 TOP10", color="申请人")
    fig_top.update_layout(showlegend=False)
    st.plotly_chart(fig_top, use_container_width=True)

# ==================== 明细导出 ====================
st.markdown("---")
with st.expander("📄 查看当前筛选全部明细数据"):
    display_cols = ["项目", "申请人", "会议号", "费用日期", "出差地点", "费用发生地点", 
                    "火车/高铁", "打车", "公交", "其他", "酒店", "差补", "总金额", "审批状态", "备注栏"]
    st.dataframe(df_display[display_cols].sort_values("费用日期", ascending=False), use_container_width=True)

export_buf = export_excel(df_display)
st.download_button(
    label="📥 导出当前筛选数据 Excel",
    data=export_buf,
    file_name="报销数据_导出.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("💡 提示：饼图支持点击项目联动；异常监控默认折叠；图表仅在数据变化时重新生成，性能优化。")
