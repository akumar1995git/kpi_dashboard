import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# ----------------------------------
# Page config and style
# ----------------------------------
st.set_page_config(
    page_title="Employee KPI Analytics Suite",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
[data-testid="stSidebar"] > div { background: #204051; color: #fff }
.stApp, .main, .block-container { background: #f7f9fc; }
.big-metric {font-size: 2em;font-weight:700; padding:18px 0 6px 0;}
.stMetricLabel, .stMetricValue {color: #101820 !important;}
</style>
""", unsafe_allow_html=True)

# ----------------------------------
# Data & Navigation
# ----------------------------------
employee_sheets = [
    'Role_vs_Reality_Analysis','Hidden_Capacity_Burnout_Risk','Work_Models_Effectiveness',
    'Digital_Collaboration_Overload','Digital_Wellbeing_Index','Data_Driven_Skill_Gap_Analysis',
    'High_Value_Work_Ratio','Future_Skill_Readiness_Index','Shadow_IT_Risk_Score'
]

# Sidebar Nav
menu = ["Summary", "Trends", "Compare Employees", "Raw Data"]
choice = st.sidebar.radio("Navigation", menu, index=0)

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type='xlsx')
source_file = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

kpi_sheet = st.sidebar.selectbox("Select KPI", employee_sheets)
df = pd.read_excel(source_file, sheet_name=kpi_sheet)

# Extract dimension columns
time_col = next((c for c in df.columns if any(x in c.lower() for x in ["report", "week", "date", "quarter"])), None)
emp_col = "Employee_ID" if "Employee_ID" in df.columns else None
metric_cols = [col for col in df.select_dtypes(include=np.number).columns if col != time_col]

# Top filters – period, employee (multi)
with st.container():
    st.title("Employee KPI Analytics Dashboard")
    filters = st.columns([2, 2, 4])
    period_vals = sorted(df[time_col].unique()) if time_col else []
    emp_vals = sorted(df[emp_col].unique()) if emp_col else []

    with filters[0]:
        if period_vals:
            period_sel = st.multiselect('Choose Period', period_vals, default=period_vals)
            df = df[df[time_col].isin(period_sel)]
    with filters[1]:
        if emp_col:
            emp_sel = st.multiselect('Choose Employees', emp_vals, default=emp_vals)
            df = df[df[emp_col].isin(emp_sel)]

# ----------------------------------
# SUMMARY DASHBOARD
# ----------------------------------
if choice == "Summary":
    st.subheader("Key Metrics Overview")
    metric_cards = st.columns(min(4, len(metric_cols)))
    for i, metric in enumerate(metric_cols[:4]):
        card_avg = df[metric].mean() if not df.empty else 0
        card_max = df[metric].max() if not df.empty else 0
        card_min = df[metric].min() if not df.empty else 0
        with metric_cards[i]:
            st.markdown(f"<div class='big-metric'>{card_avg:.2f}</div>", unsafe_allow_html=True)
            st.markdown(f"**{metric.replace('_',' ').title()}**")
            st.caption(f"Min: {card_min:.1f} | Max: {card_max:.1f}")

    st.markdown("---")
    st.subheader("Metric Distribution Snapshot")
    col_summ = st.columns(3)
    for i, metric in enumerate(metric_cols[:3]):
        with col_summ[i]:
            chart = px.histogram(df, x=metric, nbins=20, color_discrete_sequence=['#005b96'])
            chart.update_layout(height=200, margin=dict(l=10, r=10, t=30, b=30), template='simple_white', title=metric.replace('_',' '))
            st.plotly_chart(chart, use_container_width=True)

# ----------------------------------
# TRENDS
# ----------------------------------
elif choice == "Trends":
    st.subheader("Time-based Trends")
    if time_col:
        selected_metric = st.selectbox("Which metric?", metric_cols)
        fig = px.line(df, x=time_col, y=selected_metric, color=emp_col if emp_col else None,
            markers=True, line_shape="linear", color_discrete_sequence=px.colors.qualitative.Safe,
            title=f"{selected_metric.replace('_',' ')} Over Time")
        fig.update_layout(template="plotly_white", title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No time column detected in data.")

# ----------------------------------
# EMPLOYEE COMPARISON
# ----------------------------------
elif choice == "Compare Employees":
    st.subheader("Employee Metric Comparison")
    if emp_col:
        comp_metric = st.selectbox("Choose metric to compare", metric_cols)
        display_mode = st.radio("Bar or Box plot", ["Bar", "Box"], key="compare_mode")
        group_df = df.groupby(emp_col)[comp_metric].mean().reset_index() if display_mode == "Bar" else df
        if display_mode == "Bar":
            fig = px.bar(group_df, x=emp_col, y=comp_metric, title=f"Average {comp_metric.replace('_',' ')} per Employee",
                         color=comp_metric, color_continuous_scale="Magma")
        else:
            fig = px.box(group_df, x=emp_col, y=comp_metric, color=emp_col, title=f"{comp_metric.replace('_',' ')} Distribution Across Employees",
                         points="all", color_discrete_sequence=px.colors.qualitative.Bold)
        fig.update_layout(template="plotly_white", title_x=0.5, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Employee identifier column not found.")

    # Optional: Heatmap for multi-metric/employee comparison
    st.markdown("#### Multi-metric Correlation (employees × KPIs)")
    if len(metric_cols) > 1 and emp_col:
        pivot_df = df.pivot_table(index=emp_col, values=metric_cols, aggfunc='mean')
        fig = px.imshow(pivot_df, color_continuous_scale="YlGnBu", aspect="auto",
                        labels=dict(color="Value"))
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)

# ----------------------------------
# RAW DATA
# ----------------------------------
elif choice == "Raw Data":
    st.subheader("Raw Data View")
    st.dataframe(df, use_container_width=True)

    # Download button
    st.download_button("Download filtered data as CSV", df.to_csv(index=False), file_name=f"{kpi_sheet}_filtered.csv")

# ------------ Additional Files (Optional) ------------#
# If you want color configs, documentation, or config values for branding,
# add a config.py or markdown doc and import/load in your app as needed.

