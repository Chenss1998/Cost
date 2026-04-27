import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import date

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="报销数据仪表板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义CSS ====================
st.markdown("""
<style>
    /* 全局样式 */
    .stApp {
        background-color: #F0F2F6;
    }
    
    /* 顶部导航栏样式 */
    .top-nav {
        background: linear-gradient(135deg, #1E3A5F 0%, #2C5282 100%);
        padding: 1rem 2rem;
        border-radius: 0 0 20px 20px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }
    
    .logo-title {
        color: white;
        font-size: 1.5rem;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    /* 卡片样式 */
    .card {
        background: white;
        border-radius: 20px;
        padding: 1.2rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
        border: 1px solid rgba(0,0,0,0.05);
        height: 100%;
    }
    
    .card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.1);
    }
    
    .card-title {
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #64748B;
        margin-bottom: 0.5rem;
    }
    
    .card-value {
        font-size: 2rem;
        font-weight: 700;
        color: #1E3A5F;
        margin-bottom: 0.25rem;
    }
    
    .card-unit {
        font-size: 0.8rem;
        color: #94A3B8;
    }
    
    .card-trend-up {
        color: #10B981;
        font-size: 0.75rem;
        margin-top: 0.5rem;
    }
    
    .card-trend-down {
        color: #EF4444;
        font-size: 0.75rem;
        margin-top: 0.5rem;
    }
    
    /* 区块标题 */
    .section-title {
        font-size: 1.2rem;
        font-weight: 600;
        color: #1E293B;
        margin: 1.5rem 0 1rem;
        padding-left: 0.8rem;
        border-left: 4px solid #2C5282;
    }
    
    /* 侧边栏样式 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F172A 0%, #1E293B 100%);
    }
    
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stDateInput label {
        color: #E2E8F0 !important;
    }
    
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: #F1F5F9 !important;
    }
    
    /* 侧边栏按钮 */
    .stButton > button {
        background: linear-gradient(135deg, #FF6B35 0%, #F7931E 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.5rem 1rem;
        font-weight: 500;
        width: 100%;
        transition: all 0.3s;
    }
    
    .stButton > button:hover {
        opacity: 0.9;
        transform: scale(1.02);
    }
    
    /* 扩展器样式 */
    .streamlit-expanderHeader {
        background-color: #F8FAFC;
        border-radius: 12px;
        font-weight: 500;
        color: #1E3A5F;
    }
    
    /* 表格样式 */
    .dataframe {
        border-radius: 16px !important;
        overflow: hidden !important;
    }
    
    /* 信息框 */
    .stAlert {
        border-radius: 12px;
        border-left-width: 4px;
    }
    
    /* 联动高亮 */
    .project-highlight {
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        border-left: 4px solid #2C5282;
        padding: 0.75rem 1rem;
        border-radius: 12px;
        margin: 1rem 0;
    }
    
    /* 指标行 */
    .metrics-row {
        display: flex;
        gap: 1rem;
        margin-bottom: 2rem;
    }
    
    /* 图表容器 */
    .chart-container {
        background: white;
        border-radius: 20px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

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
    df_processed = df.copy()
    df_processed["费用日期"] = pd.to_datetime(df_processed["费用日期"], errors='coerce')
    df_processed = df_processed.dropna(subset=["费用日期"])
    for col in COST_COLS:
        df_processed[col] = pd.to_numeric(df_processed[col], errors='coerce').fillna(0)
    df_processed["总金额"] = df_processed[COST_COLS].sum(axis=1)
    if "备注栏" not in df_processed.columns:
        df_processed["备注栏"] = ""
    df_processed["项目"] = df_processed["项目"].apply(lambda x: x if x in PROJECTS else "其他")
    
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
    df_copy = df.copy()
    df_copy["费用日期_仅日期"] = df_copy["费用日期"].dt.date
    daily_hotel = df_copy.groupby(["申请人", "费用日期_仅日期"])["酒店"].sum().reset_index()
    daily_hotel.rename(columns={"酒店": "当日酒店总额"}, inplace=True)
    df_copy = df_copy.merge(daily_hotel, on=["申请人", "费用日期_仅日期"], how="left")
    df_copy["当日酒店超标"] = df_copy["当日酒店总额"] > 400
    
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

def export_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="报销数据")
    buffer.seek(0)
    return buffer

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### 📊 报销系统")
    st.markdown("---")
    
    st.markdown("#### 📂 数据导入")
    template = generate_template()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        template.to_excel(writer, index=False, sheet_name="报销模板")
    st.download_button(
        label="📥 下载模板",
        data=buffer.getvalue(),
        file_name="报销模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    uploaded_file = st.file_uploader("上传数据", type=["xlsx", "csv"], label_visibility="collapsed")
    
    st.markdown("---")
    st.markdown("#### 🔍 筛选条件")
    
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
    
    selected_project_filter = st.selectbox("项目", project_list)
    selected_applicant = st.selectbox("申请人", applicant_list)
    selected_status = st.selectbox("审批状态", status_list)

# ==================== 数据加载 ====================
if uploaded_file is not None:
    try:
        df_original = load_data(uploaded_file)
        st.success("✅ 数据加载成功", icon="✅")
    except Exception as e:
        st.error(f"文件读取失败: {e}")
        st.stop()
else:
    st.info("📌 请下载模板填写后上传，当前为演示数据")
    df_original = generate_template()
    st.warning("⚠️ 演示数据预览", icon="⚠️")

# 字段校验
required_cols = ["项目", "子项目", "申请人", "会议号", "费用日期", "出差地点", "费用发生地点",
                 "火车/高铁", "打车", "公交", "其他", "酒店", "差补", "审批状态"]
missing = [c for c in required_cols if c not in df_original.columns]
if missing:
    st.error(f"❌ 缺少必要字段：{missing}")
    st.stop()

# 数据处理
df_filtered = process_data(df_original, start_date, end_date, 
                           selected_project_filter, selected_applicant, selected_status)
if df_filtered.empty:
    st.warning("⚠️ 当前筛选条件下无数据，请调整筛选条件")
    st.stop()

# 会议统计
meeting_stats = compute_meeting_stats(df_filtered)
overall_avg_meeting = meeting_stats["平均单次金额"].mean() if not meeting_stats.empty else 0

# 异常计算
df_with_exception = compute_exception_flags(df_filtered)

# ==================== 饼图联动 ====================
if "selected_project" not in st.session_state:
    st.session_state.selected_project = None

project_totals = df_filtered.groupby("项目")["总金额"].sum().reset_index()
project_totals = project_totals[project_totals["项目"] != "其他"]

if not project_totals.empty:
    fig_pie = px.pie(project_totals, names="项目", values="总金额", 
                     title="项目费用占比", hole=0.4,
                     color_discrete_sequence=px.colors.qualitative.Set3)
    fig_pie.update_traces(textposition='inside', textinfo='percent+label', 
                          pull=[0.02]*len(project_totals),
                          hoverinfo='skip',
                          marker=dict(line=dict(color='white', width=2)))
    fig_pie.update_layout(
        title_font_size=16,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=400
    )
    selected_points = st.plotly_chart(fig_pie, selection_mode="points", on_select="rerun", key="pie", use_container_width=True)
    
    if selected_points and selected_points.selection and len(selected_points.selection.points) > 0:
        clicked_point = selected_points.selection.points[0]
        clicked_project = getattr(clicked_point, 'label', getattr(clicked_point, 'x', None))
        if clicked_project and clicked_project != st.session_state.selected_project:
            st.session_state.selected_project = clicked_project

col_clear, _ = st.columns([1, 5])
with col_clear:
    if st.button("🗑️ 清除联动"):
        st.session_state.selected_project = None
        st.rerun()

# 联动筛选
df_display = df_with_exception.copy()
if st.session_state.selected_project is not None:
    df_display = df_display[df_display["项目"] == st.session_state.selected_project]
    st.markdown(f'<div class="project-highlight">📌 当前显示项目：<strong>{st.session_state.selected_project}</strong>（点击饼图扇形联动）</div>', unsafe_allow_html=True)
    meeting_stats = compute_meeting_stats(df_display)
    overall_avg_meeting = meeting_stats["平均单次金额"].mean() if not meeting_stats.empty else 0

# ==================== KPI卡片 ====================
total_amt = df_display["总金额"].sum()
avg_amt = df_display["总金额"].mean()
abnormal_taxi = (df_display["打车"] > 200).sum()
approved_amt = df_display[df_display["审批状态"] == "已通过"]["总金额"].sum()
approved_ratio = approved_amt / total_amt * 100 if total_amt > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">💰 报销总额</div>
        <div class="card-value">{total_amt:,.0f}</div>
        <div class="card-unit">元</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">📈 单均报销</div>
        <div class="card-value">{avg_amt:,.0f}</div>
        <div class="card-unit">元/单</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    delta_html = f'<div class="card-trend-up">⚠️ {abnormal_taxi} 笔超标</div>' if abnormal_taxi > 0 else ''
    st.markdown(f"""
    <div class="card">
        <div class="card-title">🚕 打车超标</div>
        <div class="card-value">{abnormal_taxi}</div>
        <div class="card-unit">笔 (>200元)</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">✅ 通过占比</div>
        <div class="card-value">{approved_ratio:.1f}%</div>
        <div class="card-unit">已通过金额占比</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">🎤 会议平均</div>
        <div class="card-value">{overall_avg_meeting:,.0f}</div>
        <div class="card-unit">元/场</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ==================== 第一行图表：会议分析 ====================
st.markdown('<div class="section-title">🎤 会议费用分析</div>', unsafe_allow_html=True)
col_m1, col_m2 = st.columns(2)

with col_m1:
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.subheader("🏆 会议总金额 TOP10")
    top_meetings = meeting_stats.nlargest(10, "会议总金额")[["会议号", "会议总金额", "记录条数", "平均单次金额"]]
    st.dataframe(top_meetings, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col_m2:
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.subheader("📈 会议平均金额分布")
    if not meeting_stats.empty:
        fig_meeting = px.histogram(meeting_stats, x="平均单次金额", nbins=15,
                                   color_discrete_sequence=["#2C5282"],
                                   labels={"平均单次金额": "金额(元)", "count": "会议数量"})
        fig_meeting.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=350)
        st.plotly_chart(fig_meeting, use_container_width=True)
    else:
        st.info("暂无会议数据")
    st.markdown('</div>', unsafe_allow_html=True)

# ==================== 第二行图表：趋势 + 构成 ====================
st.markdown('<div class="section-title">📊 费用分析</div>', unsafe_allow_html=True)
col_t1, col_t2 = st.columns(2)

with col_t1:
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.subheader("📅 月度费用趋势")
    df_display["年月"] = df_display["费用日期"].dt.to_period("M").astype(str)
    monthly = df_display.groupby("年月")["总金额"].sum().reset_index()
    if not monthly.empty:
        fig_trend = px.line(monthly, x="年月", y="总金额", markers=True,
                            color_discrete_sequence=["#1E3A5F"])
        fig_trend.update_traces(marker=dict(size=8), line=dict(width=2))
        fig_trend.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=350)
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.info("暂无月度数据")
    st.markdown('</div>', unsafe_allow_html=True)

with col_t2:
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.subheader("🧾 费用类别构成")
    cost_sum = df_display[COST_COLS].sum()
    if cost_sum.sum() > 0:
        fig_cost = px.pie(values=cost_sum.values, names=cost_sum.index,
                          color_discrete_sequence=px.colors.sequential.Blues_r)
        fig_cost.update_traces(hoverinfo='skip', textposition='inside', textinfo='percent+label')
        fig_cost.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=350)
        st.plotly_chart(fig_cost, use_container_width=True)
    else:
        st.info("暂无费用数据")
    st.markdown('</div>', unsafe_allow_html=True)

# ==================== 第三行：项目对比 ====================
st.markdown('<div class="section-title">📊 项目费用对比</div>', unsafe_allow_html=True)
if st.session_state.selected_project:
    compare_df = df_filtered.groupby("项目")["总金额"].sum().reset_index()
    fig_compare = px.bar(compare_df, x="项目", y="总金额", color="项目",
                         title=f"各项目总费用（当前选中：{st.session_state.selected_project}）",
                         color_discrete_map={st.session_state.selected_project: "#E53E3E"})
    fig_compare.update_layout(showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=400)
    st.plotly_chart(fig_compare, use_container_width=True)
else:
    proj_sum = df_display.groupby("项目")["总金额"].sum().sort_values(ascending=False).reset_index()
    fig_proj = px.bar(proj_sum, x="项目", y="总金额", title="各项目报销总额",
                      color="项目", color_discrete_sequence=px.colors.qualitative.Set3)
    fig_proj.update_layout(showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=400)
    st.plotly_chart(fig_proj, use_container_width=True)

# ==================== 第四行：审批分析 ====================
st.markdown('<div class="section-title">📋 审批情况</div>', unsafe_allow_html=True)
col_a1, col_a2 = st.columns(2)

with col_a1:
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    status_cnt = df_display["审批状态"].value_counts().reset_index()
    status_cnt.columns = ["审批状态", "单据数"]
    fig_status = px.bar(status_cnt, x="审批状态", y="单据数", color="审批状态",
                        title="各状态单据数量")
    fig_status.update_layout(showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=350)
    st.plotly_chart(fig_status, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col_a2:
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    status_amt = df_display.groupby("审批状态")["总金额"].sum().reset_index()
    fig_amt = px.pie(status_amt, values="总金额", names="审批状态", title="各状态金额占比")
    fig_amt.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=350)
    st.plotly_chart(fig_amt, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("### 👥 申请人审批状态统计")
cross_tab = pd.crosstab(df_display["申请人"], df_display["审批状态"])
if not cross_tab.empty:
    styled_cross = cross_tab.style.background_gradient(cmap="Blues", axis=None)
    st.dataframe(styled_cross, use_container_width=True)

# ==================== 第五行：员工排行 ====================
st.markdown('<div class="section-title">🏆 员工报销排行</div>', unsafe_allow_html=True)
top_applicants = df_display.groupby("申请人")["总金额"].sum().sort_values(ascending=False).head(10).reset_index()
if not top_applicants.empty:
    fig_top = px.bar(top_applicants, x="申请人", y="总金额", title="申请人报销总额 TOP10",
                     color="申请人", color_discrete_sequence=px.colors.qualitative.Set3)
    fig_top.update_layout(showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=400)
    st.plotly_chart(fig_top, use_container_width=True)

# ==================== 异常监控 ====================
abnormal_count_total = (df_display["异常类型"] != "正常").sum()
st.info(f"📋 共有 **{abnormal_count_total}** 条异常单据，点击下方查看详情")

with st.expander("⚠️ 异常监控详情"):
    st.markdown("**异常规则：**")
    col_r1, col_r2, col_r3 = st.columns(3)
    col_r1.markdown("- 🔴 地点不一致")
    col_r2.markdown("- 🔴 当日酒店 > 400元")
    col_r3.markdown("- 🔴 月酒店 > 13000元")
    st.markdown("- 🟡 打车 > 200元（黄色高亮）")
    
    abnormal_df = df_display[df_display["异常类型"] != "正常"]
    if not abnormal_df.empty:
        show_cols = ["项目", "申请人", "会议号", "费用日期", "出差地点", "费用发生地点",
                     "总金额", "酒店", "打车", "异常类型", "审批状态"]
        st.dataframe(abnormal_df[show_cols], use_container_width=True)
    else:
        st.success("✅ 无异常单据")

# ==================== 明细导出 ====================
with st.expander("📄 明细数据"):
    display_cols = ["项目", "申请人", "会议号", "费用日期", "出差地点", "费用发生地点",
                    "火车/高铁", "打车", "公交", "其他", "酒店", "差补", "总金额", "审批状态", "备注栏"]
    st.dataframe(df_display[display_cols].sort_values("费用日期", ascending=False), use_container_width=True)

export_buf = export_excel(df_display)
st.download_button(
    label="📥 导出 Excel",
    data=export_buf,
    file_name="报销数据.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("💡 提示：点击饼图扇形可联动筛选；表格带渐变色背景；图表仅在数据变化时更新。")
