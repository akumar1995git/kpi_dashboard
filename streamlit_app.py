import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# ---------- SETUP ----------
st.set_page_config(
    page_title="Professional Employee KPI Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Color palette for charts
COLORS = px.colors.qualitative.Safe

employee_sheets = [
    'Role_vs_Reality_Analysis',
    'Hidden_Capacity_Burnout_Risk',
    'Work_Models_Effectiveness',
    'Digital_Collaboration_Overload',
    'Digital_Wellbeing_Index',
    'Data_Driven_Skill_Gap_Analysis',
    'High_Value_Work_Ratio',
    'Future_Skill_Readiness_Index',
    'Shadow_IT_Risk_Score'
]

# -------------- SIDEBAR ------------------
st.sidebar.title("KPI Filter")

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type='xlsx')
if uploaded_file:
    excel_file = uploaded_file
else:
    excel_file = "Updated_18_KPI_Dashboard.xlsx"

selected_kpi = st.sidebar.selectbox("Select Employee KPI", employee_sheets)
df = pd.read_excel(excel_file, sheet_name=selected_kpi)

emp_col = "Employee_ID"
has_emp = emp_col in df.columns
if has_emp:
    emp_list = sorted(df[emp_col].unique())
    selected_emp = st.sidebar.selectbox("Select Employee", ["All"] + list(emp_list))
    if selected_emp != "All":
        df = df[df[emp_col] == selected_emp]

# Auto-detect time and numeric columns
time_col = next((col for col in df.columns if 'report' in col.lower() or 'week' in col.lower() or 'quarter' in col.lower()), None)
numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
metric_cols = [c for c in numeric_cols if c != time_col]

# ---------- HEADER ----------
st.title("Professional Employee KPI Analytics Dashboard")
st.markdown("""
Current period: **April 2025 - September 2025**  
Explore detailed trends, distributions, and performance recommendations for core employee KPIs.
""")

# ---------- KEY METRICS TILES ----------
def kpi_color(val, reverse=False):
    # Color scale for green/yellow/red metric cards
    thresholds = np.nanpercentile(df[val], [30, 70])
    best = thresholds[1] if reverse else thresholds[0]
    worst = thresholds[0] if reverse else thresholds[1]
    avg = np.nanmean(df[val])
    if avg >= best:
        return "🟢"
    elif avg <= worst:
        return "🔴"
    else:
        return "🟡"

cols = st.columns(len(metric_cols[:4]))
for i, mc in enumerate(metric_cols[:4]):
    emoji = kpi_color(mc, reverse="risk" in mc.lower() or "overload" in mc.lower())
    val = df[mc].mean()
    st.session_state[mc] = val
    with cols[i]:
        st.metric(
            label=f"{mc.replace('_',' ').title()}",
            value=f"{val:,.2f}",
            delta=f"{df[mc].max()-df[mc].min():.2f} peak-to-low",
            delta_color="normal"
        )
        st.markdown(f"{emoji} <small>Avg. last 6 months</small>", unsafe_allow_html=True)

# Insights & Suggestions
def gen_insights(df, metric_col):
    vals = df[metric_col].dropna()
    if len(vals) < 3:
        return "Not enough data for insight."
    trend = np.polyfit(range(len(vals)), vals, 1)[0]
    recent = vals.iloc[-1]
    early = vals.iloc[0]
    pctchange = (recent - early) / abs(early) * 100 if early else np.nan
    is_good = ('risk' not in metric_col.lower() and 'overload' not in metric_col.lower())
    insight = ""
    if np.isnan(pctchange):
        return "No reliable trend found."
    if (pctchange > 0 and is_good) or (pctchange < 0 and not is_good):
        insight = "📈 **Positive trend**: Metric has improved over the period."
    else:
        insight = "⚠️ **Warning**: Metric has deteriorated over the period."
    suggestion = ""
    if 'risk' in metric_col.lower() or 'overload' in metric_col.lower():
        if recent > np.nanpercentile(vals, 70):
            suggestion = "🔺 Try workload balancing or stress reduction interventions."
        else:
            suggestion = "✅ Current risk levels are within optimal bounds."
    elif 'readiness' in metric_col.lower() or 'score' in metric_col.lower():
        if recent < np.nanpercentile(vals, 30):
            suggestion = "🔺 Enhance skill training programs and monitor future-skill engagement."
        else:
            suggestion = "✅ Readiness score is on track. Continue with ongoing efforts."
    elif 'wellbeing' in metric_col.lower():
        if recent < np.nanpercentile(vals, 40):
            suggestion = "🔺 Employee well-being is suboptimal; consider wellness programs."
        else:
            suggestion = "✅ Well-being levels are healthy compared to previous periods."
    elif 'collaboration' in metric_col.lower():
        if recent > np.nanpercentile(vals, 70):
            suggestion = "🔺 Meeting or collaboration overload detected, audit time allocation."
        else:
            suggestion = "✅ Collaboration level is balanced."
    else:
        if recent < np.nanpercentile(vals, 40):
            suggestion = "🔺 Monitor metric closely; recent dip detected."
        else:
            suggestion = "✅ No anomaly detected."
    return f"{insight}\n\n{suggestion}"

st.markdown("### 📊 Visualizations & Insights")

# -------- MAIN GRAPHIC AND INSIGHTS ---------
tab1, tab2, tab3, tab4 = st.tabs(["Trend", "Distribution", "Compare Employees", "Insights & Suggestions"])

with tab1:
    metric_to_plot = st.selectbox("Select Metric to Plot (Trend)", metric_cols, key="trend_metric")
    if time_col and metric_to_plot:
        fig = px.line(df, x=time_col, y=metric_to_plot, title=f"{metric_to_plot.replace('_',' ')} Trend",
                      color_discrete_sequence=[COLORS[0]])
        fig.update_traces(mode="lines+markers", line=dict(width=3))
        fig.update_layout(template="ggplot2", plot_bgcolor="white", title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Trend chart not available.")

with tab2:
    dist_metric = st.selectbox("Select Metric for Distribution", metric_cols, key="dist_metric")
    fig = px.histogram(df, x=dist_metric, nbins=15, title=f"{dist_metric.replace('_',' ')} Distribution", color_discrete_sequence=[COLORS[1]])
    fig.update_layout(template="simple_white", title_x=0.5)
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    if has_emp and metric_cols:
        box_metric = st.selectbox("Compare Employees by Metric", metric_cols, key="box_metric")
        fig = px.box(df, x=emp_col, y=box_metric, color=emp_col, points="all", color_discrete_sequence=COLORS, title=f"{box_metric.replace('_',' ')} Across Employees")
        fig.update_layout(template="plotly_white", showlegend=False, title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No employee comparison available.")

with tab4:
    st.subheader("Automated Insights")
    for metric in metric_cols[:4]:
        st.markdown(f"#### {metric.replace('_',' ').title()}")
        st.info(gen_insights(df, metric))

# ------- RAW DATA PREVIEW -------
with st.expander("🔽 See full raw KPI data", expanded=False):
    st.dataframe(df, use_container_width=True, height=300)

st.markdown("---")
st.caption("Dashboard best viewed on a wide screen. Powered by Plotly/Streamlit. Data: Updated_18_KPI_Dashboard.xlsx")

# ------------- PROFESSIONAL COLOR THEMES ------------
st.markdown(
    """
    <style>
    .stApp {background-color: #FAFBFF;}
    .stMetricLabel, .stMetricValue {color: #1f3a56 !important;}
    .css-1r6slb0 {background-color: #E9EFFB !important;}
    .stTabs [data-baseweb="tab-list"] button {background: #335C81 !important; color: #fff;}
    </style>
    """,
    unsafe_allow_html=True,
)
