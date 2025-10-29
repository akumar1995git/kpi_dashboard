import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="Comprehensive Organizational Analytics Dashboard")

@st.cache_data
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    data = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    return data

uploaded_file = st.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type=["xlsx"])

if uploaded_file:
    data = load_data(uploaded_file)

    st.title("Comprehensive Organizational Analytics Dashboard")
    
    # Executive Summary & Key Insights
    st.header("Executive Summary & Key Insights")

    # 1. Burnout Risk Summary
    if "Hidden_Capacity_Burnout_Risk" in data:
        df = data["Hidden_Capacity_Burnout_Risk"]
        avg_burnout = df['Burnout_Risk_Score'].mean()
        high_burnout_pct = (df['Burnout_Risk_Score'] > 4.5).mean() * 100
        st.markdown(f"### Burnout Risk\nAverage Score: **{avg_burnout:.2f}**. High Burnout (>4.5): **{high_burnout_pct:.1f}%**.")
        fig = px.histogram(df, x='Burnout_Risk_Score', nbins=20, title="Burnout Risk Score Distribution")
        st.plotly_chart(fig, use_container_width=True)

    # 2. High Value Work Summary
    if "High_Value_Work_Ratio" in data:
        df = data["High_Value_Work_Ratio"]
        avg_hv = df['High_Value_Work_Percentage'].mean()*100
        st.markdown(f"### High Value Work\nAverage %: **{avg_hv:.1f}%**.")
        agg = df.groupby('Reporting_Period')['High_Value_Work_Percentage'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='High_Value_Work_Percentage',
                      title="High Value Work Over Time",
                      labels={"High_Value_Work_Percentage": "High Value Work %"})
        st.plotly_chart(fig, use_container_width=True)

    # 3. Digital Wellbeing
    if "Digital_Wellbeing_Index" in data:
        df = data["Digital_Wellbeing_Index"]
        avg_dw = df['Digital_Wellbeing_Score'].mean()
        st.markdown(f"### Digital Wellbeing\nAverage Score: **{avg_dw:.2f}**.")
        agg = df.groupby('Reporting_Period')['Digital_Wellbeing_Score'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='Digital_Wellbeing_Score',
                      title="Digital Wellbeing Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # 4. Organizational Agility
    if "Organizational_Agility_Score" in data:
        df = data["Organizational_Agility_Score"]
        avg_agility = df['Agility_Score'].mean()
        st.markdown(f"### Organizational Agility\nAverage Score: **{avg_agility:.2f}**.")
        agg = df.groupby('Quarter')['Agility_Score'].mean().reset_index()
        fig = px.bar(agg, x='Quarter', y='Agility_Score', title="Organizational Agility by Quarter")
        st.plotly_chart(fig, use_container_width=True)

    # 5. Skill Gap Analysis
    if "Data_Driven_Skill_Gap_Analysis" in data:
        df = data["Data_Driven_Skill_Gap_Analysis"]
        agg = df.groupby('Skill_Category')['Skill_Gap_Score'].mean().reset_index()
        highest_gap = agg.loc[agg['Skill_Gap_Score'].idxmax()]
        st.markdown(f"### Skill Gaps\nLargest average gap in **{highest_gap.Skill_Category}**: {highest_gap.Skill_Gap_Score:.2f}")
        fig = px.bar(agg, x='Skill_Category', y='Skill_Gap_Score', title="Average Skill Gap by Category")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # Detailed Charts Section
    st.header("In-Depth Analytics")

    # Role vs Reality - Time Low Value Tasks by Reporting Period and Employee
    if "Role_vs_Reality_Analysis" in data:
        df = data["Role_vs_Reality_Analysis"]
        fig = px.bar(df, x='Reporting_Period', y='Time_Low_Value_Tasks_Hours', color='EmployeeID',
                     title="Low Value Task Hours by Employee Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Capacity Utilization Trend by Employee
    if "Hidden_Capacity_Burnout_Risk" in data:
        df = data["Hidden_Capacity_Burnout_Risk"]
        fig = px.line(df, x='Week_Ending_Date', y='Capacity_Utilization_Percentage', color='EmployeeID',
                      title="Capacity Utilization Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Work Model Productivity Index
    if "Work_Models_Effectiveness" in data:
        df = data["Work_Models_Effectiveness"]
        fig = px.box(df, x='WorkModel', y='ProductivityIndex', title="Productivity Index by Work Model")
        st.plotly_chart(fig, use_container_width=True)

    # Meeting Load Hours Distribution
    if "Digital_Wellbeing_Index" in data:
        df = data["Digital_Wellbeing_Index"]
        fig = px.histogram(df, x='MeetingLoadHours', nbins=20, title="Meeting Load Hours Distribution")
        st.plotly_chart(fig, use_container_width=True)

    # Collaboration Overload Percentage Over Time
    if "Digital_Collaboration_Overload" in data:
        df = data["Digital_Collaboration_Overload"]
        fig = px.line(df, x='WeekEndingDate', y='CollaborationOverloadPercentage', color='EmployeeID',
                      title="Collaboration Overload Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Shadow IT Risk Score Distribution
    if "Shadow_IT_Risk_Score" in data:
        df = data["Shadow_IT_Risk_Score"]
        fig = px.histogram(df, x='Risk_Score', nbins=20, title="Shadow IT Risk Score Distribution")
        st.plotly_chart(fig, use_container_width=True)

    # Process Brittleness Over Time
    if "Process_Brittleness_Index" in data:
        df = data["Process_Brittleness_Index"]
        agg = df.groupby(['ProcessID', 'ReportingPeriod']).agg(Brittleness=('BrittlenessPercentage', 'mean')).reset_index()
        fig = px.line(agg, x='ReportingPeriod', y='Brittleness', color='ProcessID',
                      title="Process Brittleness Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Automation Velocity Impact
    if "Automation_Velocity_Impact" in data:
        df = data["Automation_Velocity_Impact"]
        agg = df.groupby(['AutomationProjectID', 'ReportingPeriod']).agg(Savings=('ProcessAutomationSavingsHours', 'sum')).reset_index()
        fig = px.bar(agg, x='ReportingPeriod', y='Savings', color='AutomationProjectID',
                     title="Automation Savings Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Order to Delivery Processing Time
    if "OrderToDelivery" in data:
        df = data["OrderToDelivery"]
        fig = px.line(df, x='WeekEndingDate', y='ProcessingTimeDays', title="Order to Delivery Processing Time Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Issue to Resolution Processing Time
    if "IssueToResolution" in data:
        df = data["IssueToResolution"]
        fig = px.line(df, x='WeekEndingDate', y='ProcessingTimeDays', title="Issue to Resolution Processing Time Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Cross Functional Process Latency by Process
    if "CrossFunctionalProcessLatency" in data:
        df = data["CrossFunctionalProcessLatency"]
        fig = px.line(df, x='WeekEndingDate', y='ProcessingTimeDays', color='ProcessName', title="Cross Functional Process Latency")
        st.plotly_chart(fig, use_container_width=True)

    # Vulnerability Hotspots by Task
    if "VulnerabilityHotspots" in data:
        df = data["VulnerabilityHotspots"]
        fig = px.bar(df, x='CriticalTaskID', y='IndividualVulnerabilityPercentage', color='IsVulnerable',
                     title="Individual Vulnerability Percentage by Task")
        st.plotly_chart(fig, use_container_width=True)

   
else:
    st.info("Please upload the Excel file to generate the comprehensive dashboard.")
