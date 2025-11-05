import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# --- Page and style ---
st.set_page_config(page_title="Employee KPI Analytics", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
    .big-title {font-size:2.5em;font-weight:700;margin-bottom:-12px;}
    .kpi-card {padding:8px 15px 0 0;}
    .delta-green {color: #13a813;}
    .delta-red {color: #d62728;}
    hr {margin:2em 0;}
    </style>
    """, unsafe_allow_html=True
)

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

# --- Load, standardize, concat ---
dfs = []
for sheet in employee_kpi_sheets:
    df = pd.read_excel(source_file, sheet_name=sheet)
    # Pick preferred time col, rename as 'Time_Period'
    found = next((c for c in TIME_COLS if c in df.columns), None)
    if found:
        df = df.rename(columns={found:'Time_Period'})
    else:
        df['Time_Period'] = np.nan
    df['KPI_Sheet'] = sheet
    dfs.append(df)
full_df = pd.concat(dfs, ignore_index=True)

emp_col = 'Employee_ID'
periods = sorted(full_df['Time_Period'].dropna().unique())
employees = sorted(full_df[emp_col].dropna().unique())
metrics = [c for c in full_df.select_dtypes(np.number).columns if c not in ['Time_Period',emp_col]]

# ----------------
# HEADER
# ----------------
logo_url = "https://upload.wikimedia.org/wikipedia/commons/f/f5/Emblem_of_India.svg"  # Example: Replace with your branding/logo
cols = st.columns([1,8])
with cols[0]: st.image(logo_url, width=60)
with cols[1]:
    st.markdown('<div class="big-title">Your Organization - Employee KPI Report</div>', unsafe_allow_html=True)
    st.caption("Contact: helpdesk@example.com |  Data refreshed automatically | Powered by Streamlit")

# ----------------
# FILTERS
# ----------------
with st.container():
    flt = st.columns([2,2,6])
    periods_sel = flt[0].multiselect("Period", periods, default=periods[-3:] if len(periods)>2 else periods)
    emps_sel = flt[1].multiselect("Employee", employees, default=employees)
    # Optionally filter by KPI
    # kpi_sel = flt[2].multiselect("KPI", employee_kpi_sheets, default=employee_kpi_sheets)
    # full_df = full_df[full_df['KPI_Sheet'].isin(kpi_sel)]
filtered = full_df[full_df['Time_Period'].isin(periods_sel) & full_df[emp_col].isin(emps_sel)]

# ----------------
# TOP SUMMARY Cards (pick most actionable metrics)
# ----------------
show_metrics = metrics[:3]  # Change for your target metrics
curr, prev = filtered[metrics].mean(), filtered[metrics].shift(1).mean()
summary = st.columns(len(show_metrics))
for i,m in enumerate(show_metrics):
    v, v_prev = curr.get(m,0), prev.get(m,0)
    delta = v-v_prev if not pd.isna(v_prev) else 0
    delta_arrow = "↓" if delta < 0 else "↑"
    delta_style = "delta-green" if delta < 0 else "delta-red"
    summary[i].markdown(f"""
    <div class="kpi-card">
    <span style="font-size:2em;font-weight:600">{v:.1f}</span><br>
    <span style="font-size:1em;font-weight:500">{m.replace('_',' ').title()}</span><br>
    <span class="{delta_style}">{delta_arrow} {abs(delta):.1f} compared to previous</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<hr />", unsafe_allow_html=True)

# ----------------
# INSIGHTFUL CHARTS (no pie, no clutter)
# ----------------
for sheet in employee_kpi_sheets:
    dfi = filtered[filtered['KPI_Sheet']==sheet]
    st.header(sheet.replace('_',' '))
    # Trend over time
    timec = 'Time_Period'
    kc = [c for c in dfi.select_dtypes(np.number).columns if c not in ['Time_Period',emp_col]]
    if not dfi.empty and timec in dfi.columns and kc:
        c = kc[0]
        fig = px.bar(dfi, x=timec, y=c, color=emp_col, barmode='group', height=250,
                    title=f"{c.replace('_',' ')} by Period and Employee", color_discrete_sequence=px.colors.qualitative.Safe)
        st.plotly_chart(fig, use_container_width=True)
    # Per-employee distribution
    if not dfi.empty and kc:
        c = kc[0]
        fig2 = px.box(dfi, x=emp_col, y=c, color=emp_col, title=f"{c.replace('_',' ')} Distribution by Employee",
                      points='outliers', color_discrete_sequence=px.colors.qualitative.Plotly)
        st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")
st.subheader("Raw Data")
st.dataframe(filtered, use_container_width=True, height=350)
st.download_button("Download current table as CSV", filtered.to_csv(index=False), "filtered_kpis.csv")
