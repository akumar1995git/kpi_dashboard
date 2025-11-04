import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# ---------------------------------------
# Setup page config and CSS for modes
# ---------------------------------------
st.set_page_config(page_title="Refined Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
    <style>
    .stApp, .main, .block-container {background-color: #f7f9fc; color:#101820;}
    .stMetricLabel, .stMetricValue {color: #101820 !important;}
    .stButton>button {background-color: #005b96 !important; color: white !important;}
    .stTextInput>div>input {background-color: #e7eef9 !important; color:#101820 !important;}
    .stSelectbox>div>div>div>select, .stSelectbox>div>div>div>option {background-color:#c4d1e8; color:#101820;}
    .css-1r6slb0 {background-color: #e6efff !important;}
    .stTabs [data-baseweb=\"tab-list\"] button {background:#005b96 !important;color:#fff;font-weight:bold;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------
# Employee KPI Sheets (as before)
# ---------------------------------------
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

# ---------------------------------------
# Load data, selection filters, basic info
# ---------------------------------------
uploaded_file = st.sidebar.file_uploader("Load Updated_18_KPI_Dashboard.xlsx", type=['xlsx'])
excel_path = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

selected_kpi = st.sidebar.selectbox("Select Employee KPI", employee_sheets)
df = pd.read_excel(excel_path, sheet_name=selected_kpi)

# Employee filter
emp_col = 'Employee_ID'
if emp_col in df.columns:
    employees = sorted(df[emp_col].unique())
    selected_emp = st.sidebar.selectbox("Filter by Employee", ['All'] + employees)
    if selected_emp != 'All':
        df = df[df[emp_col] == selected_emp]

# Time and metric columns
time_col = next((c for c in df.columns if 'report' in c.lower() or 'week' in c.lower() or 'quarter' in c.lower()), None)
metric_cols = [c for c in df.select_dtypes(include=np.number).columns if c != time_col]

# ---------------------------------------
# Layout - Filters pinned atop page for clean UI
# ---------------------------------------
st.title("Refined Employee KPI Dashboard")

select_container = st.container()
with select_container:
    col1, col2 = st.columns([3,1])
    with col1:
        st.markdown(f"### KPI: `{selected_kpi.replace('_',' ')}`")
    with col2:
        if emp_col in df.columns:
            emp_drop = st.selectbox("Employee", ['All'] + employees, index=0)
            if emp_drop != 'All':
                df = df[df[emp_col] == emp_drop]

# ---------------------------------------
# Summary Kpi Cards in a row
# ---------------------------------------
st.markdown("#### Key Metrics (Averages)")
cols = st.columns(min(4, len(metric_cols)))

def color_scale(v, good_high=True):
    if good_high:
        if v >= 70: return '🟢'
        elif v >= 50: return '🟡'
        else: return '🔴'
    else:
        if v <= 30: return '🟢'
        elif v <= 50: return '🟡'
        else: return '🔴'

for idx, metric in enumerate(metric_cols[:4]):
    val = df[metric].mean()
    # Metrics containing risk or overload - lower is better
    good_high = not ('risk' in metric.lower() or 'overload' in metric.lower() or 'burnout' in metric.lower())
    emoji = color_scale(val, good_high)

    with cols[idx]:
        st.metric(label=metric.replace('_',' ').title(), value=f"{val:.2f}", delta_color="normal")
        st.markdown(f"{emoji} <small>Average Value</small>", unsafe_allow_html=True)

st.markdown("---")

# ---------------------------------------
# Visualization with sparklines & gauges
# ---------------------------------------
st.markdown("### KPI Trend (Summary Sparkline)")

if time_col and metric_cols:
    spark_col = st.selectbox("Select Metric For Sparkline", metric_cols, key="sparkline_metric")

    # Prepare sparkline data
    spark_data = df[[time_col, spark_col]]
    spark_data_grouped = spark_data.groupby(time_col)[spark_col].mean().reset_index()
    spark_data_grouped = spark_data_grouped.sort_values(time_col)

    fig_spark = go.Figure(go.Scatter(y=spark_data_grouped[spark_col], x=spark_data_grouped[time_col], mode='lines+markers'))
    fig_spark.update_layout(
        height=150,
        margin=dict(l=20, r=20, t=30, b=20),
        xaxis_title=time_col.replace('_',' ').title(),
        yaxis_title=spark_col.replace('_',' ').title(),
        template='plotly_white'
    )

    st.plotly_chart(fig_spark, use_container_width=True)

else:
    st.info("No time-based data available for sparkline.")

st.markdown("---")

# ---------------------------------------
# Gauges for risk/score KPIs or selected metric
# ---------------------------------------
st.markdown("### Key KPI Gauges & Distribution")

risk_like_metrics = [c for c in metric_cols if any(word in c.lower() for word in ['risk', 'overload', 'burnout'])]

col_probs = st.columns(len(risk_like_metrics) if risk_like_metrics else 1)

if risk_like_metrics:
    for i, metric in enumerate(risk_like_metrics):
        avg_val = df[metric].mean()
        with col_probs[i]:
            gauge = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=avg_val,
                delta={'reference': np.percentile(df[metric], 50), 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
                gauge={'axis': {'range': [0, max(df[metric]) * 1.2]}, 'bar': {'color': "#ef553b"}},
                title={'text': metric.replace('_', ' ').title()}
            ))
            st.plotly_chart(gauge, use_container_width=True)
else:
    st.info("No risk-related metrics found for gauges.")

st.markdown("---")

# ---------------------------------------
# Insight Generation
# ---------------------------------------
def generate_insights(series):
    if len(series) < 3:
        return "Insufficient data for insights."
    trend = np.polyfit(range(len(series)), series, 1)[0]
    direction = "increasing" if trend > 0 else "decreasing"

    insight_msg = f"The trend indicates the metric is {direction} over time.\n"
    if trend > 0:
        insight_msg += "This is a positive sign. Maintain current strategies and monitor for consistency."
    else:
        insight_msg += "This suggests potential issues or risk. Investigate root causes and plan interventions."

    return insight_msg

st.markdown("### Automated Insights & Suggestions")

if metric_cols:
    for metric in metric_cols[:5]:
        vals = df[metric].dropna().values
        st.markdown(f"**{metric.replace('_',' ').title()}**")
        st.info(generate_insights(vals))
else:
    st.info("No numeric KPIs available for insight generation.")

# ---------------------------------------
# Raw data viewer
# ---------------------------------------
st.markdown("---")
with st.expander("View Raw Employee KPI Data"):
    st.dataframe(df)

st.markdown(
    "Streamlit dashboard built with care for clarity, professional looks, and user ease."
)
