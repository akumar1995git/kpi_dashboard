import streamlit as st
import pandas as pd
import plotly.express as px

# List only the employee-related KPI sheets
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

st.title("Employee KPI Dashboard")
st.write("Select and explore detailed KPIs for individual employees.")

# Load data from Excel - use your local file or repo link
excel_file = "Updated_18_KPI_Dashboard.xlsx"  # Put file path here or from GitHub repo

# Sidebar for KPI selection
selected_kpi = st.sidebar.selectbox("Select Employee KPI", employee_sheets)

# Load selected sheet
df = pd.read_excel(excel_file, sheet_name=selected_kpi)

# Show data preview
st.subheader(f"Data Preview: {selected_kpi.replace('_', ' ')}")
st.dataframe(df.head())

# Basic overview statistics
st.write("Summary Statistics")
st.write(df.describe())

# Employee dropdown
if "Employee_ID" in df.columns:
    emp_list = df["Employee_ID"].unique()
    selected_emp = st.selectbox("Select Employee", emp_list)
    emp_df = df[df["Employee_ID"] == selected_emp]
    st.write(f"Details for Employee: {selected_emp}")
    st.dataframe(emp_df)

    # Visualization
    # Detect columns to plot time trends (reporting period/date/quarter)
    time_col = next((col for col in df.columns if col.lower() in ["reporting_period", "week_ending_date", "quarter"]), None)
    metric_cols = [col for col in df.columns if col not in ["Employee_ID", time_col]]

    if time_col and metric_cols:
        metric_to_plot = st.selectbox("Select Metric to Plot", metric_cols)
        fig = px.line(emp_df, x=time_col, y=metric_to_plot, title=f"{metric_to_plot} over time for {selected_emp}")
        st.plotly_chart(fig)
else:
    st.write("No Employee_ID column found in this sheet.")

st.write("---")
st.markdown("Data source: [Updated_18_KPI_Dashboard.xlsx](link to repo or upload)")

