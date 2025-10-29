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
    
    # Executive Summary with Insightful Commentary
    st.header("Executive Summary & Key Insights")

    # Burnout Risk Summary and Insight
    if "Hidden_Capacity_Burnout_Risk" in data:
        df_burnout = data["Hidden_Capacity_Burnout_Risk"]
        avg_burnout = df_burnout['Burnout_Risk_Score'].mean()
        high_burnout_pct = (df_burnout['Burnout_Risk_Score'] > 4.5).mean() * 100
        st.markdown(f"**Burnout Risk:** Average score is **{avg_burnout:.2f}**. "
                    f"Critically, **{high_burnout_pct:.1f}%** of employees show high burnout risk (>4.5).")
        fig = px.histogram(df_burnout, x='Burnout_Risk_Score', nbins=20,
                           title="Distribution of Burnout Risk Scores")
        st.plotly_chart(fig, use_container_width=True)

    # High Value Work Summary and Insight
    if "High_Value_Work_Ratio" in data:
        df_hv = data["High_Value_Work_Ratio"]
        avg_hv_pct = df_hv['High_Value_Work_Percentage'].mean()*100
        st.markdown(f"**High Value Work:** Employees spend on average **{avg_hv_pct:.1f}%** of their time on high-value activities. "
                    "Improving this ratio is key to organizational efficiency.")
        agg = df_hv.groupby('Reporting_Period')['High_Value_Work_Percentage'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='High_Value_Work_Percentage',
                      labels={"High_Value_Work_Percentage": "High Value Work %", "Reporting_Period": "Reporting Period"},
                      title="Average High Value Work Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Digital Wellbeing Insight
    if "Digital_Wellbeing_Index" in data:
        df_dw = data["Digital_Wellbeing_Index"]
        avg_dw_score = df_dw['Digital_Wellbeing_Score'].mean()
        st.markdown(f"**Digital Wellbeing:** Average wellbeing score is **{avg_dw_score:.2f}** out of 10. "
                    "Sustained attention to reducing overload can improve this metric.")
        agg = df_dw.groupby('Reporting_Period')['Digital_Wellbeing_Score'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='Digital_Wellbeing_Score',
                      title="Digital Wellbeing Score Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Organizational Agility Summary
    if "Organizational_Agility_Score" in data:
        df_ag = data["Organizational_Agility_Score"]
        avg_agility = df_ag['Agility_Score'].mean()
        st.markdown(f"**Organizational Agility:** The average agility score across departments is **{avg_agility:.2f}**, "
                    "highlighting areas of strength and potential improvement.")
        agg = df_ag.groupby('Quarter')['Agility_Score'].mean().reset_index()
        fig = px.bar(agg, x='Quarter', y='Agility_Score',
                     labels={"Agility_Score": "Average Agility Score"},
                     title="Organizational Agility Score by Quarter")
        st.plotly_chart(fig, use_container_width=True)

    # Skill Gap Analysis Summary
    if "Data_Driven_Skill_Gap_Analysis" in data:
        df_skill = data["Data_Driven_Skill_Gap_Analysis"]
        agg = df_skill.groupby('Skill_Category')['Skill_Gap_Score'].mean().reset_index()
        highest_gap = agg.loc[agg['Skill_Gap_Score'].idxmax()]
        st.markdown(f"**Skill Gaps:** The area with largest average gap is **{highest_gap.Skill_Category}** "
                    f"with a score of **{highest_gap.Skill_Gap_Score:.2f}**, signaling training priorities.")
        fig = px.bar(agg, x='Skill_Category', y='Skill_Gap_Score', title="Average Skill Gap by Category")
        st.plotly_chart(fig, use_container_width=True)

    # Process Brittleness Summary
    if "Process_Brittleness_Index" in data:
        df_brittle = data["Process_Brittleness_Index"]
        avg_brittleness = df_brittle['Brittleness_Percentage'].mean()
        st.markdown(f"**Process Brittleness:** The average brittleness across processes is **{avg_brittleness:.2f}%**, "
                    "indicating potential fragility in workflows requiring attention.")
        agg = df_brittle.groupby('Reporting_Period')['Brittleness_Percentage'].mean().reset_index()
        fig = px.line(agg, x='Reporting_Period', y='Brittleness_Percentage',
                      title="Process Brittleness Over Time")
        st.plotly_chart(fig, use_container_width=True)

    # Automation Impact Summary
    if "Automation_Velocity_Impact" in data:
        df_auto = data["Automation_Velocity_Impact"]
        total_savings = df_auto['Process_Automation_Savings_Hours'].sum()
        st.markdown(f"**Automation Impact:** Total time saved due to process automation projects is **{total_savings:.0f} hours**, "
                    "boosting overall efficiency.")
        agg = df_auto.groupby('Reporting_Period')['Process_Automation_Savings_Hours'].sum().reset_index()
        fig = px.bar(agg, x='Reporting_Period', y='Process_Automation_Savings_Hours',
                     title="Automation Savings Over Time")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.info("This dashboard presents summarized KPIs and key insights to help you identify organizational strengths and areas requiring strategic focus. Use these insights to drive targeted interventions.")

else:
    st.info("Please upload the Excel file to begin analysis and visualization.")
