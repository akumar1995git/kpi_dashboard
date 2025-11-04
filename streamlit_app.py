import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="Organizational Analytics Dashboard")

@st.cache_data
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    data = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    return data

uploaded_file = st.file_uploader("Upload Updated_18_KPI_Dashboard.xlsx", type=["xlsx"])

if uploaded_file:
    data = load_data(uploaded_file)

    st.title("Organizational Analytics Dashboard")

    st.header("Executive Summary & Key Insights")

   

    # High Value Work
    if "High_Value_Work_Ratio" in data:
        df = data["High_Value_Work_Ratio"]
        avg_hv = df['High_Value_Work_Percentage'].mean() * 100
        st.markdown(f"**High Value Work:** Average percentage **{avg_hv:.1f}%**")
        agg = df.groupby('Reporting_Period')['High_Value_Work_Percentage'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='High_Value_Work_Percentage',
                      title="High Value Work Percentage Over Time",
                      labels={"High_Value_Work_Percentage": "High Value Work %"})
        st.plotly_chart(fig, use_container_width=True)

    # Digital Wellbeing
    if "Digital_Wellbeing_Index" in data:
        df = data["Digital_Wellbeing_Index"]
        avg_dw = df['Digital_Wellbeing_Score'].mean()
        st.markdown(f"**Digital Wellbeing:** Average score {avg_dw:.2f}")
        agg = df.groupby('Reporting_Period')['Digital_Wellbeing_Score'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='Digital_Wellbeing_Score',
                      title="Digital Wellbeing Score Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Organizational Agility
    if "Organizational_Agility_Score" in data:
        df = data["Organizational_Agility_Score"]
        avg_ag = df['Agility_Score'].mean()
        st.markdown(f"**Organizational Agility:** Average score {avg_ag:.2f}")
        agg = df.groupby('Quarter')['Agility_Score'].mean().reset_index()
        fig = px.bar(agg, x='Quarter', y='Agility_Score', title="Organizational Agility Score by Quarter")
        st.plotly_chart(fig, use_container_width=True)

    # Skill Gap Analysis
    if "Data_Driven_Skill_Gap_Analysis" in data:
        df = data["Data_Driven_Skill_Gap_Analysis"]
        agg = df.groupby('Skill_Category')['Skill_Gap_Score'].mean().reset_index()
        highest_gap = agg.loc[agg['Skill_Gap_Score'].idxmax()]
        st.markdown(f"**Skill Gaps:** Largest gap in **{highest_gap.Skill_Category}**: {highest_gap.Skill_Gap_Score:.2f}")
        fig = px.bar(agg, x='Skill_Category', y='Skill_Gap_Score', title="Skill Gap Score by Category")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.header("Detailed Interactive Visualizations")

    # Role vs Reality - Low Value Tasks over Time by Employee_ID
    if "Role_vs_Reality_Analysis" in data:
        df = data["Role_vs_Reality_Analysis"]
        fig = px.bar(df, x='Reporting_Period', y='Time_Low_Value_Tasks_Hours', color='Employee_ID',
                     title="Low Value Task Hours by Employee and Period")
        st.plotly_chart(fig, use_container_width=True)

    # Capacity Utilization trend
    if "Hidden_Capacity_Burnout_Risk" in data:
        df = data["Hidden_Capacity_Burnout_Risk"]
        fig = px.line(df, x='Week_Ending_Date', y='Capacity_Utilization_Percentage', color='Employee_ID',
                      title="Weekly Capacity Utilization")
        st.plotly_chart(fig, use_container_width=True)

    # Work Model Productivity
    if "Work_Models_Effectiveness" in data:
        df = data["Work_Models_Effectiveness"]
        fig = px.box(df, x='Work_Model', y='Productivity_Index', title="Productivity Index by Work Model")
        st.plotly_chart(fig, use_container_width=True)

    # Meeting Load Hours Distribution
    if "Digital_Wellbeing_Index" in data:
        df = data["Digital_Wellbeing_Index"]
        fig = px.histogram(df, x='Meeting_Load_Hours', nbins=20, title="Meeting Load Hours Distribution")
        st.plotly_chart(fig, use_container_width=True)

    # Collaboration Overload percentage over time
    if "Digital_Collaboration_Overload" in data:
        df = data["Digital_Collaboration_Overload"]
        fig = px.line(df, x='Week_Ending_Date', y='Collaboration_Overload_Percentage', color='Employee_ID',
                      title="Collaboration Overload Over Weeks")
        st.plotly_chart(fig, use_container_width=True)

    # Shadow IT Risk Score distribution
    if "Shadow_IT_Risk_Score" in data:
        df = data["Shadow_IT_Risk_Score"]
        fig = px.histogram(df, x='Risk_Score', nbins=20, title="Shadow IT Risk Score Distribution")
        st.plotly_chart(fig, use_container_width=True)

    # Process Brittleness over time
    if "Process_Brittleness_Index" in data:
        df = data["Process_Brittleness_Index"]
        agg = df.groupby(['Process_ID', 'Reporting_Period']).agg(Brittleness=('Brittleness_Percentage', 'mean')).reset_index()
        fig = px.line(agg, x='Reporting_Period', y='Brittleness', color='Process_ID', title="Process Brittleness")
        st.plotly_chart(fig, use_container_width=True)

    # Automation Savings over time
    if "Automation_Velocity_Impact" in data:
        df = data["Automation_Velocity_Impact"]
        agg = df.groupby(['Automation_Project_ID', 'Reporting_Period']).agg(Savings=('Process_Automation_Savings_Hours', 'sum')).reset_index()
        fig = px.bar(agg, x='Reporting_Period', y='Savings', color='Automation_Project_ID', title="Automation Savings")
        st.plotly_chart(fig, use_container_width=True)

    # Order to Delivery Processing Time
    if "OrderToDelivery" in data:
        df = data["OrderToDelivery"]
        fig = px.line(df, x='Week_Ending_Date', y='Processing_Time_Days', title="Order to Delivery Processing Time")
        st.plotly_chart(fig, use_container_width=True)

    # Issue to Resolution Processing Time
    if "IssueToResolution" in data:
        df = data["IssueToResolution"]
        fig = px.line(df, x='Week_Ending_Date', y='Processing_Time_Days', title="Issue to Resolution Processing Time")
        st.plotly_chart(fig, use_container_width=True)

    # Cross-Functional Process Latency
    if "CrossFunctionalProcessLatency" in data:
        df = data["CrossFunctionalProcessLatency"]
        fig = px.line(df, x='Week_Ending_Date', y='Processing_Time_Days', color='Process_Name', title="Cross-Functional Process Latency")
        st.plotly_chart(fig, use_container_width=True)

    # Vulnerability Hotspots by Critical Task
    if "VulnerabilityHotspots" in data:
        df = data["VulnerabilityHotspots"]
        fig = px.bar(df, x='Critical_Task_ID', y='Individual_Vulnerability_Percentage', color='Is_Vulnerable', title="Vulnerability Hotspots")
        st.plotly_chart(fig, use_container_width=True)

    
else:
    st.info("Please upload the Excel file to begin analysis and visualization.")
