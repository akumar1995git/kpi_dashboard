import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# --- Page Setup ---
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
TIME_COLS = ['Reporting_Period', 'Week_Ending_Date', 'Quarter', 'Date']

uploaded_file = st.sidebar.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type="xlsx")
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
    if 'Employee_ID' not in df:
        df['Employee_ID'] = pd.NA
    df['KPI_Sheet'] = sheet
    dfs.append(df)

full_df = pd.concat(dfs, ignore_index=True)
full_df['Month'] = full_df['Month'].fillna('Unknown')
full_df['Employee_ID'] = full_df['Employee_ID'].fillna('Unknown')

periods = sorted([p for p in full_df['Month'].unique() if p and str(p).lower() not in ['nat', 'unknown', 'nan']])
employees = sorted([e for e in full_df['Employee_ID'].unique() if e and str(e).lower() not in ['unknown', 'nan']])
metric_cols = [col for col in full_df.select_dtypes(np.number).columns if col != 'Employee_ID']

# UI/Brand Header Example
st.markdown("""
<div style='display:flex;align-items:center;padding-top:8px'>
    <img src='https://upload.wikimedia.org/wikipedia/commons/f/f5/Emblem_of_India.svg' width='55'>
    <span style='font-size:2.4em;font-weight:700;padding:10px 0 0 14px'>Employee KPI Management Dashboard</span>
</div>
""", unsafe_allow_html=True)

flt_area = st.container()
with flt_area:
    col1, col2 = st.columns([2,2])
    sel_months = col1.multiselect("Select Month(s)", periods, default=periods[-3:] if len(periods)>3 else periods)
    sel_emps = col2.multiselect("Select Employees", employees, default=employees)

filtered_df = full_df[(full_df['Month'].isin(sel_months)) & (full_df['Employee_ID'].isin(sel_emps))]
if filtered_df.empty:
    st.warning("No matching data. Adjust your month/employee selection.")
    st.stop()

# --- Top KPI Cards with Deltas ---
main_metrics = metric_cols[:3]
delta_cols = st.columns(len(main_metrics))
sorted_months = sorted(filtered_df['Month'].unique())
last_month, prev_month = sorted_months[-1], sorted_months[-2] if len(sorted_months)>1 else None

for i, metric in enumerate(main_metrics):
    v_now = filtered_df[filtered_df['Month'] == last_month][metric].mean() if last_month else np.nan
    v_prev = filtered_df[filtered_df['Month'] == prev_month][metric].mean() if prev_month else np.nan
    delta_txt, badge_color = "N/A", "#aaa"
    if pd.notna(v_now) and pd.notna(v_prev):
        delta_val = v_now - v_prev
        if delta_val > 0:
            delta_txt, badge_color = f"↑ {delta_val:.1f} vs prev", "#d62728"
        elif delta_val < 0:
            delta_txt, badge_color = f"↓ {abs(delta_val):.1f} vs prev", "#13a813"
        else:
            delta_txt, badge_color = "No change", "#888"
    delta_cols[i].markdown(f"""
        <div style='font-size:2em;font-weight:700'>{v_now:.1f}</div>
        <div style='font-weight:400;font-size:1.2em'>{metric.replace('_',' ').title()}</div>
        <div style='background:{badge_color};padding:4px 12px;border-radius:8px;color:#fff;display:inline-block;margin-top:4px;'>{delta_txt}</div>
    """, unsafe_allow_html=True)

st.markdown("<hr />", unsafe_allow_html=True)

# --- Chart area: clean, grouped bar, color-mapped bar, summary table ---
chart_col1, chart_col2, chart_col3 = st.columns([1,1,1])

with chart_col1:
    # 1. Number of "Completed" or "Active" entries per Month-Employee as bar chart
    comp_col = main_metrics[0]
    agg = filtered_df.groupby(['Month','Employee_ID'])[comp_col].mean().reset_index()
    fig = px.bar(agg, x='Month', y=comp_col, color='Employee_ID', barmode='group', height=320,
                 title=f"{comp_col.replace('_',' ')} by Month, grouped by Employee")
    st.plotly_chart(fig, use_container_width=True)

with chart_col2:
    # 2. Distribution bar: per month (like predicted arrivals, color by month)
    pred_col = main_metrics[1] if len(main_metrics) > 1 else comp_col
    agg2 = filtered_df.groupby(['Month'])[pred_col].mean().reset_index()
    fig2 = px.bar(agg2, x='Month', y=pred_col, color=pred_col, color_continuous_scale='Blues',
                  title=f"Monthly Avg {pred_col.replace('_',' ')}", height=320)
    st.plotly_chart(fig2, use_container_width=True)

with chart_col3:
    # 3. Gradient bar (average duration or other target, with color mapping, target line)
    dur_col = main_metrics[2] if len(main_metrics) > 2 else comp_col
    agg3 = filtered_df.groupby(['Month'])[dur_col].mean().reset_index()
    fig3 = px.bar(agg3, x='Month', y=dur_col, color=dur_col, color_continuous_scale='sunset',
                  title=f"Avg {dur_col.replace('_',' ')} with Target", height=320)
    # Add reference/target line (example target = avg of all months or set value)
    target_val = filtered_df[dur_col].mean()
    fig3.add_shape(type="line",
        x0=0, x1=len(agg3)-1, y0=target_val, y1=target_val,
        line=dict(color="black", width=2), xref="x", yref="y"
    )
    fig3.add_trace(px.line(agg3, x='Month', y=[target_val]*len(agg3)).data[0])
    st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# --- Summary Table ("Current status") ---
st.subheader("Employee KPI Monthly Summary")
summary_table = filtered_df.groupby(['Month','Employee_ID'])[main_metrics].mean()
st.dataframe(summary_table.round(2), use_container_width=True)
st.download_button("Download as CSV", summary_table.to_csv(), "kpi_monthly_summary.csv")

# --- Optional: expand for other charts or raw data table below ---

