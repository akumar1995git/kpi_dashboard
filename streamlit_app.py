import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="Organizational Analytics Dashboard")

# Load all sheets into a dictionary of DataFrames
@st.cache_data
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    df_sheets = {sheet: xls.parse(sheet) for sheet in sheets}
    return df_sheets

uploaded_file = st.file_uploader("Upload the Updated_18_KPI_Dashboard Excel file", type=["xlsx"])

if uploaded_file:
    data = load_data(uploaded_file)
    
    st.title("Organizational Analytics Dashboard")
    
    # Display Executive Summary
    st.header("Executive Summary")
    
    # Sample summaries (customize to your needs)
    try:
        # Burnout Risk overview
        burnout_df = data.get("Hidden_Capacity_Burnout_Risk")
        if burnout_df is not None:
            avg_burnout = burnout_df['Burnout_Risk_Score'].mean()
            max_burnout = burnout_df['Burnout_Risk_Score'].max()
            st.markdown(f"- Average Burnout Risk Score: **{avg_burnout:.2f}**")
            st.markdown(f"- Maximum Burnout Risk Score observed: **{max_burnout:.2f}**")
        
        # High Value Work Overview
        high_value_df = data.get("High_Value_Work_Ratio")
        if high_value_df is not None:
            avg_high_value_pct = high_value_df['High_Value_Work_Percentage'].mean() * 100
            st.markdown(f"- Average High Value Work Percentage: **{avg_high_value_pct:.1f}%**")
        
        # Organizational Agility Score
        agility_df = data.get("Organizational_Agility_Score")
        if agility_df is not None:
            avg_agility = agility_df['Agility_Score'].mean()
            st.markdown(f"- Average Organizational Agility Score: **{avg_agility:.2f}**")

    except Exception as e:
        st.warning("Error calculating Executive Summary: " + str(e))

    st.markdown("---")
    
    # Show sheet-wise charts in expandable sections

    # 1. Role vs Reality Analysis
    with st.expander("Role vs Reality Analysis"):
        df = data.get("Role_vs_Reality_Analysis")
        if df is not None:
            st.subheader("Time Spent on Low-Value Tasks")
            fig1 = px.bar(df, x='Reporting_Period', y='Time_Low_Value_Tasks_Hours', 
                          color='Employee_ID',
                          labels={'Reporting_Period': 'Month', 'Time_Low_Value_Tasks_Hours': 'Hours on Low Value Tasks'},
                          title='Time Spent on Low-Value Tasks by Employee Over Months')
            st.plotly_chart(fig1, use_container_width=True)

            st.subheader("Percentage of Low Value Work")
            fig2 = px.line(df, x='Reporting_Period', y='Low_Value_Work_Percentage',
                           color='Employee_ID',
                           labels={'Reporting_Period': 'Month', 'Low_Value_Work_Percentage': 'Low Value Work %'},
                           title='Low Value Work Percentage Over Months')
            st.plotly_chart(fig2, use_container_width=True)
    
    # 2. Burnout Risk and Capacity Utilization
    with st.expander("Burnout Risk and Capacity Utilization"):
        df = data.get("Hidden_Capacity_Burnout_Risk")
        if df is not None:
            st.subheader("Capacity Utilization Over Time")
            fig = px.line(df, x='Week_Ending_Date', y='Capacity_Utilization_Percentage',
                          color='Employee_ID',
                          labels={'Week_Ending_Date': 'Week Ending Date', 'Capacity_Utilization_Percentage': 'Capacity Utilization'},
                          title='Weekly Capacity Utilization per Employee')
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Burnout Risk Score Distribution")
            fig = px.histogram(df, x='Burnout_Risk_Score',
                               nbins=20,
                               title='Burnout Risk Score Distribution')
            st.plotly_chart(fig, use_container_width=True)

    # 3. Work Models Effectiveness
    with st.expander("Work Models Effectiveness"):
        df = data.get("Work_Models_Effectiveness")
        if df is not None:
            st.subheader("Productivity Index by Work Model")
            fig = px.box(df, x='Work_Model', y='Productivity_Index', 
                         title='Productivity Index Distribution by Work Model')
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Output Units Completed over Time")
            fig = px.line(df, x='Reporting_Period', y='Output_Units_Completed',
                          color='Work_Model',
                          labels={'Reporting_Period': 'Month'},
                          title='Output Units Completed by Work Model Over Months')
            st.plotly_chart(fig, use_container_width=True)

    # 4. Digital Wellbeing Index
    with st.expander("Digital Wellbeing Index"):
        df = data.get("Digital_Wellbeing_Index")
        if df is not None:
            st.subheader("Digital Wellbeing Score over Reporting Period")
            fig = px.line(df, x='Reporting_Period', y='Digital_Wellbeing_Score',
                          color='Employee_ID',
                          labels={'Reporting_Period': 'Date'},
                          title='Digital Wellbeing Score Trends')
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Meeting Load Hours Histogram")
            fig = px.histogram(df, x='Meeting_Load_Hours', nbins=20,
                               title='Distribution of Meeting Load Hours')
            st.plotly_chart(fig, use_container_width=True)

    # 5. Organizational Agility Score
    with st.expander("Organizational Agility Score"):
        df = data.get("Organizational_Agility_Score")
        if df is not None:
            st.subheader("Agility Score by Department and Quarter")
            fig = px.bar(df, x='Department_ID', y='Agility_Score', color='Quarter',
                         barmode='group',
                         title="Organizational Agility Score per Department")
            st.plotly_chart(fig, use_container_width=True)
    
    # 6. Skill Gap Analysis
    with st.expander("Data Driven Skill Gap Analysis"):
        df = data.get("Data_Driven_Skill_Gap_Analysis")
        if df is not None:
            st.subheader("Average Skill Gap Score by Skill Category")
            agg = df.groupby('Skill_Category')['Skill_Gap_Score'].mean().reset_index()
            fig = px.bar(agg, x='Skill_Category', y='Skill_Gap_Score', 
                         title="Average Skill Gap Score by Skill Category")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Performance vs Benchmark Score Scatter")
            fig = px.scatter(df, x='Benchmark_Score', y='Performance_Score',
                             color='Skill_Category',
                             title="Performance vs Benchmark by Skill Category")
            st.plotly_chart(fig, use_container_width=True)

    # 7. Process Brittleness Index
    with st.expander("Process Brittleness Index"):
        df = data.get("Process_Brittleness_Index")
        if df is not None:
            st.subheader("Brittleness Percentage by Process Over Months")
            fig = px.line(df, x='Reporting_Period', y='Brittleness_Percentage', color='Process_ID', 
                          title="Process Brittleness Over Time")
            st.plotly_chart(fig, use_container_width=True)

    # 8. Automation Velocity Impact
    with st.expander("Automation Velocity Impact"):
        df = data.get("Automation_Velocity_Impact")
        if df is not None:
            st.subheader("Process Automation Savings Hours Over Time")
            fig = px.line(df, x='Reporting_Period', y='Process_Automation_Savings_Hours', color='Automation_Project_ID',
                          title="Automation Time Savings by Project Over Months")
            st.plotly_chart(fig, use_container_width=True)

    # 9. Process Adherence Rate
    with st.expander("Process Adherence Rate"):
        df = data.get("Process_Adherence_Rate")
        if df is not None:
            st.subheader("Process Adherence Rate Over Time")
            fig = px.line(df, x='Date', y='Adherence_Rate_Percentage', color='Process_ID',
                          title="Process Adherence Rate by Process Over Days")
            st.plotly_chart(fig, use_container_width=True)

    # 10. Digital Collaboration Overload
    with st.expander("Digital Collaboration Overload"):
        df = data.get("Digital_Collaboration_Overload")
        if df is not None:
            st.subheader("Collaboration Overload Percentage Over Time")
            fig = px.line(df, x='Week_Ending_Date', y='Collaboration_Overload_Percentage', color='Employee_ID',
                          title="Collaboration Overload per Employee Over Weeks")
            st.plotly_chart(fig, use_container_width=True)

    # 11. Shadow IT Risk Score
    with st.expander("Shadow IT Risk Score"):
        df = data.get("Shadow_IT_Risk_Score")
        if df is not None:
            st.subheader("Risk Score Distribution")
            fig = px.histogram(df, x='Risk_Score', nbins=20,
                               title="Distribution of Shadow IT Risk Scores")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Unauthorized Apps Count Over Time")
            fig = px.line(df, x='Week_Ending_Date', y='Unauthorized_Apps_Count', color='Employee_ID',
                          title="Unauthorized Apps Count per Employee Over Time")
            st.plotly_chart(fig, use_container_width=True)

    # 12. Future Skill Readiness Index
    with st.expander("Future Skill Readiness Index"):
        df = data.get("Future_Skill_Readiness_Index")
        if df is not None:
            st.subheader("Readiness Score by Quarter")
            fig = px.bar(df, x='Quarter', y='Readiness_Score',
                         title="Readiness Score Over Quarters")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Skills Alignment Percentage Distribution")
            fig = px.histogram(df, x='Skills_Alignment_Percentage',
                               title="Distribution of Skills Alignment Percentage")
            st.plotly_chart(fig, use_container_width=True)

    # Add more sheets and visualizations similarly...

else:
    st.info("Please upload the Excel file to load data and display dashboards.")
