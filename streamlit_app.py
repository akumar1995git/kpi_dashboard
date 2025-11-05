import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

employee_kpi_sheets = [
    'Role_vs_Reality_Analysis',
    'Hidden_Capacity_Burnout_Risk',
    'Work_Models_Effectiveness',
    'Digital_Collaboration_Overload',
    'Digital_Wellbeing_Index',
    'High_Value_Work_Ratio',
    'Future_Skill_Readiness_Index'
]
TIME_COLS = ['Reporting_Period','Week_Ending_Date','Quarter','Date']

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type='xlsx')
source_file = uploaded_file if uploaded_file else "Updated_18_KPI_Dashboard.xlsx"

dfs = []
for sheet in employee_kpi_sheets:
    df = pd.read_excel(source_file, sheet_name=sheet)
    found_time = next((c for c in TIME_COLS if c in df.columns), None)
    if found_time:
        if np.issubdtype(df[found_time].dtype, np.datetime64):
            df['Month'] = pd.to_datetime(df[found_time]).dt.to_period('M').astype(str)
        else:
            df['Month'] = df[found_time].astype(str).str[:7]
    else:
        df['Month'] = pd.NA
    df['KPI_Sheet'] = sheet
    if 'Employee_ID' not in df:
        df['Employee_ID'] = pd.NA
    dfs.append(df)

full_df = pd.concat(dfs, ignore_index=True)
full_df['Month'] = full_df['Month'].fillna('Unknown')
full_df['Employee_ID'] = full_df['Employee_ID'].fillna('Unknown')

periods = sorted([p for p in full_df['Month'].unique() if p and p != 'Unknown'])
employees = sorted([e for e in full_df['Employee_ID'].unique() if e and e != 'Unknown'])
metric_cols = [col for col in full_df.select_dtypes(include=np.number).columns if col != 'Employee_ID']

# ---- UI HEADER (mimic SWAST look) ----
st.markdown("""
<div style='text-align:center;padding-top:10px;padding-bottom:5px'>
    <img src='https://upload.wikimedia.org/wikipedia/commons/f/f5/Emblem_of_India.svg' width='60' style='vertical-align:middle;'>
    <span style='font-size:2.3em;font-weight:700;padding-left:15px;'>Your Organization - Employee KPI Dashboard</span>
</div>
<div style='text-align:center;font-size:1.1em;color:#666;padding-bottom:20px'>
    Data period: <b>{0}</b> to <b>{1}</b>
</div>
""".format(periods[0] if periods else '', periods[-1] if periods else ''), unsafe_allow_html=True)

# ---- Filter bar ----
st.markdown("""
<div style='background:#e4e8ee;border-radius:8px;padding:14px 18px 4px 18px;margin-bottom:12px'>
""", unsafe_allow_html=True)
flt_cols = st.columns([2,2])
with flt_cols[0]:
    selected_periods = st.multiselect('Select Month(s)', periods, default=periods[-3:] if len(periods)>3 else periods)
with flt_cols[1]:
    selected_emps = st.multiselect('Select Employees', employees, default=employees)
st.markdown("</div>", unsafe_allow_html=True)

filtered_df = full_df[
    (full_df['Month'].isin(selected_periods)) & 
    (full_df['Employee_ID'].isin(selected_emps))
]

if filtered_df.empty:
    st.warning("No matching data. Adjust your period or employee selection.")
    st.stop()

# ---- Top Summary Cards ----
st.markdown("<hr />", unsafe_allow_html=True)
metric_columns = metric_cols[:4]
card_cols = st.columns(len(metric_columns))
last_month, prev_month = None, None
if len(selected_periods) >= 2:
    last_month, prev_month = selected_periods[-1], selected_periods[-2]

for i, metric in enumerate(metric_columns):
    # Current and previous month means
    v, v_prev = np.nan, np.nan
    if last_month:
        v = filtered_df[filtered_df['Month'] == last_month][metric].mean()
    if prev_month:
        v_prev = filtered_df[filtered_df['Month'] == prev_month][metric].mean()
    delta = None if pd.isna(v) or pd.isna(v_prev) else v - v_prev
    color = "#13a813" if delta is not None and delta < 0 else "#d62728"
    icon = "↓" if delta is not None and delta < 0 else "↑"
    delta_text = f"{icon} {abs(delta):.1f} vs prev" if delta is not None else "N/A"
    # Insight sentence
    if delta is None:
        insight = "Not enough data to show change."
    elif delta > 0:
        insight = "Metric has increased since previous month."
    elif delta < 0:
        insight = "Metric has decreased since previous month."
    else:
        insight = "No month-on-month change."
    with card_cols[i]:
        st.markdown(f"""
        <div class='kpi-card' style='padding:5px 15px 7px 0'>
            <span style='font-size:2em;font-weight:700'>{v:.1f}</span>
            <span style='font-size:1em;font-weight:500'>{metric.replace('_',' ').title()}</span><br>
            <span style='color:{color}'><b>{delta_text}</b></span>
            <div style='font-size:0.96em;margin-top:4px;color:#444'>{insight}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<hr />", unsafe_allow_html=True)

# ---- Trends and Distributions ----
for metric in metric_columns:
    left, right = st.columns([3,2])
    with left:
        st.markdown(f"#### {metric.replace('_',' ').title()} by Month, per Employee")
        chart_df = filtered_df.groupby(['Month','Employee_ID'])[metric].mean().reset_index()
        fig = px.line(chart_df, x='Month', y=metric, color='Employee_ID',
                      markers=True, line_shape='linear', template='plotly_white', height=280)
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.markdown("##### Employee Distribution (last selected month)")
        box_df = filtered_df[filtered_df['Month']==last_month] if last_month else filtered_df
        if box_df.empty:
            st.info("No data for this month.")
        else:
            fig2 = px.box(box_df, x='Employee_ID', y=metric, color='Employee_ID', points='outliers', template='plotly_white', height=280)
            st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")
st.subheader("Raw KPI Data")
st.dataframe(filtered_df, use_container_width=True)
st.download_button("Download Filtered Data as CSV", filtered_df.to_csv(index=False), "filtered_employee_kpis.csv")
