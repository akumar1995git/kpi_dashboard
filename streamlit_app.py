import streamlit as st
import pandas as pd
import plotly.express as px
import altair as alt

# 1. Page configuration
st.set_page_config(
    page_title="Organizational Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)
alt.themes.enable("dark")

# 2. Load data
# Replace with your actual file name
data = pd.read_excel('Updated_18_KPI_Dashboard.xlsx')

# 3. Sidebar filters
st.sidebar.title("Filters")
departments = data['Department'].unique()
selected_dept = st.sidebar.multiselect("Select Department(s)", departments, default=list(departments))
months = data['Month'].unique()
selected_months = st.sidebar.multiselect("Select Month(s)", months, default=list(months))

filtered = data[
    data['Department'].isin(selected_dept) &
    data['Month'].isin(selected_months)
]

st.title("Organizational Analytics Dashboard")

# 4. KPI Metric Cards
st.subheader("Quick KPIs")
col1, col2, col3 = st.columns(3)
col1.metric("Avg Burnout Score", f"{filtered['Burnout Score'].mean():.2f}")
col2.metric("Total Automation Savings (₹)", f"{filtered['Automation Savings'].sum():,.0f}")
col3.metric("High Value Work (%)", f"{filtered['HighValueWork'].mean():.1f}")

# 5. Burnout Distribution
st.subheader("Burnout Risk Distribution")
fig_burnout = px.histogram(filtered, x="Burnout Score", color="Department", nbins=10)
st.plotly_chart(fig_burnout, use_container_width=True)

# 6. Time Drain Breakdown
st.subheader("Time Drain Breakdown")
fig_time = alt.Chart(filtered).mark_bar().encode(
    x='LowValueTasks:Q',
    y='Department:N',
    color='Department:N'
)
st.altair_chart(fig_time, use_container_width=True)

# 7. Collaboration and Meetings Chart
st.subheader("Collaboration Load")
fig_meetings = px.box(filtered, x="Department", y="MeetingLoad")
st.plotly_chart(fig_meetings, use_container_width=True)

# 8. Automation Impact Over Time
st.subheader("Automation Savings Over Time")
fig_auto = px.line(filtered, x="Month", y="Automation Savings", color="Department", markers=True)
st.plotly_chart(fig_auto, use_container_width=True)

# 9. Agility (Radar Chart)
st.subheader("Departmental Agility Comparison")
agility_table = filtered.groupby('Department')['AgilityScore'].mean().reset_index()
fig_agility = px.line_polar(agility_table, r='AgilityScore', theta='Department', line_close=True)
st.plotly_chart(fig_agility, use_container_width=True)

# 10. Skill Gap Pie Chart
st.subheader("Skill Gap Analysis")
skill_gaps = filtered['SkillGapCategory'].value_counts().reset_index()
fig_gap = px.pie(skill_gaps, names='index', values='SkillGapCategory')
st.plotly_chart(fig_gap, use_container_width=True)

# 11. Risk Alerts Table
st.subheader("Risk Alerts")
critical_alerts = filtered[filtered['Burnout Score'] > 4.5]
st.dataframe(critical_alerts[['Employee Name', 'Department', 'Burnout Score']], height=250)

# 12. Download Options
st.sidebar.download_button(
    label="Download Filtered Data",
    data=filtered.to_csv(index=False),
    file_name="filtered_KPI_dashboard.csv",
    mime="text/csv"
)

st.sidebar.markdown("Made with Streamlit and Python 📈")

# End of script
