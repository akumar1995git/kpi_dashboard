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

    # Display sheet names for reference
    st.sidebar.header("Sheets in Excel")
    st.sidebar.write(list(data.keys()))

    # Sidebar filters for Department and Reporting Period
    dept_options = []
    period_options = []

    # Try to infer Department options from Role_vs_Reality_Analysis sheet
    if "Role_vs_Reality_Analysis" in data:
        df_role = data["Role_vs_Reality_Analysis"]
        if 'Department_ID' in df_role.columns:
            dept_options = df_role['Department_ID'].dropna().unique().tolist()
        if 'Reporting_Period' in df_role.columns:
            period_options = sorted(df_role['Reporting_Period'].dropna().unique())

    selected_depts = st.sidebar.multiselect("Select Departments", options=dept_options, default=dept_options)
    selected_periods = st.sidebar.multiselect("Select Reporting Periods", options=period_options, default=period_options)

    st.title("Organizational Analytics Dashboard")

    # Executive Summary - Key Aggregates
    st.header("Executive Summary")

    try:
        # Burnout Risk Summary
        if "Hidden_Capacity_Burnout_Risk" in data:
            df_burnout = data["Hidden_Capacity_Burnout_Risk"]
            if 'Department_ID' in df_burnout.columns:
                df_burnout = df_burnout[df_burnout['Department_ID'].isin(selected_depts)]
            avg_burnout = df_burnout['Burnout_Risk_Score'].mean()
            st.markdown(f"- Average Burnout Risk Score: **{avg_burnout:.2f}**")

        # High Value Work Summary
        if "High_Value_Work_Ratio" in data:
            df_hv = data["High_Value_Work_Ratio"]
            if 'Department_ID' in df_hv.columns:
                df_hv = df_hv[df_hv['Department_ID'].isin(selected_depts)]
            avg_hv = df_hv['High_Value_Work_Percentage'].mean() * 100
            st.markdown(f"- Average High Value Work Percentage: **{avg_hv:.1f}%**")

        # Organizational Agility Summary
        if "Organizational_Agility_Score" in data:
            df_agility = data["Organizational_Agility_Score"]
            if 'Department_ID' in df_agility.columns:
                df_agility = df_agility[df_agility['Department_ID'].isin(selected_depts)]
            avg_agility = df_agility['Agility_Score'].mean()
            st.markdown(f"- Average Organizational Agility Score: **{avg_agility:.2f}**")
    except Exception as e:
        st.warning("Executive summary calculation error: " + str(e))

    st.markdown("---")

    # Role vs Reality Analysis - Average Low Value Work %
    if "Role_vs_Reality_Analysis" in data:
        df = data["Role_vs_Reality_Analysis"]
        df = df[(df['Department_ID'].isin(selected_depts)) & (df['Reporting_Period'].isin(selected_periods))]
        agg = df.groupby(['Department_ID', 'Reporting_Period']).agg(
            Avg_Low_Value_Hours=('Time_Low_Value_Tasks_Hours', 'mean'),
            Avg_Low_Value_Pct=('Low_Value_Work_Percentage', 'mean')
        ).reset_index()

        st.subheader("Low Value Work - Average Percentage by Department and Period")
        fig = px.bar(agg, x='Reporting_Period', y='Avg_Low_Value_Pct', color='Department_ID', barmode='group',
                     labels={'Avg_Low_Value_Pct': 'Avg % Low Value Work', 'Reporting_Period': 'Period'})
        st.plotly_chart(fig, use_container_width=True)

    # Burnout Risk Score trend by Department and Quarter
    if "Hidden_Capacity_Burnout_Risk" in data:
        df = data["Hidden_Capacity_Burnout_Risk"]
        df = df[(df['Department_ID'].isin(selected_depts))]
        agg = df.groupby(['Department_ID', 'Quarter']).agg(Avg_Burnout=('Burnout_Risk_Score', 'mean')).reset_index()
        st.subheader("Burnout Risk by Department and Quarter")
        fig = px.line(agg, x='Quarter', y='Avg_Burnout', color='Department_ID', markers=True)
        st.plotly_chart(fig, use_container_width=True)

    # High Value Work Percentage trend by Department and Reporting Period
    if "High_Value_Work_Ratio" in data:
        df = data["High_Value_Work_Ratio"]
        df = df[(df['Department_ID'].isin(selected_depts)) & (df['Reporting_Period'].isin(selected_periods))]
        agg = df.groupby(['Department_ID', 'Reporting_Period']).agg(Avg_HV_Work=('High_Value_Work_Percentage', 'mean')).reset_index()
        st.subheader("High Value Work Percentage by Department and Period")
        fig = px.line(agg, x='Reporting_Period', y='Avg_HV_Work', color='Department_ID', markers=True,
                      labels={'Avg_HV_Work': 'Avg High Value Work %'})
        st.plotly_chart(fig, use_container_width=True)

    # Digital Wellbeing Score trend
    if "Digital_Wellbeing_Index" in data:
        df = data["Digital_Wellbeing_Index"]
        df = df[(df['Department_ID'].isin(selected_depts)) & (df['Reporting_Period'].isin(selected_periods))]
        agg = df.groupby(['Department_ID', 'Reporting_Period']).agg(Avg_Wellbeing=('Digital_Wellbeing_Score', 'mean')).reset_index()
        st.subheader("Digital Wellbeing Score by Department and Reporting Period")
        fig = px.line(agg, x='Reporting_Period', y='Avg_Wellbeing', color='Department_ID', markers=True)
        st.plotly_chart(fig, use_container_width=True)

    # Organizational Agility Score by Department and Quarter
    if "Organizational_Agility_Score" in data:
        df = data["Organizational_Agility_Score"]
        df = df[(df['Department_ID'].isin(selected_depts))]
        agg = df.groupby(['Department_ID', 'Quarter']).agg(Avg_Agility=('Agility_Score', 'mean')).reset_index()
        st.subheader("Organizational Agility Score by Department and Quarter")
        fig = px.bar(agg, x='Quarter', y='Avg_Agility', color='Department_ID', barmode='group')
        st.plotly_chart(fig, use_container_width=True)

    # Data Driven Skill Gap Analysis by Skill Category
    if "Data_Driven_Skill_Gap_Analysis" in data:
        df = data["Data_Driven_Skill_Gap_Analysis"]
        agg = df.groupby('Skill_Category').agg(Avg_Skill_Gap=('Skill_Gap_Score', 'mean')).reset_index()
        st.subheader("Average Skill Gap by Skill Category")
        fig = px.bar(agg, x='Skill_Category', y='Avg_Skill_Gap')
        st.plotly_chart(fig, use_container_width=True)

    # Process Brittleness Index by Process and Reporting Period
    if "Process_Brittleness_Index" in data:
        df = data["Process_Brittleness_Index"]
        agg = df.groupby(['Process_ID', 'Reporting_Period']).agg(Avg_Brittleness=('Brittleness_Percentage', 'mean')).reset_index()
        st.subheader("Process Brittleness Over Time")
        fig = px.line(agg, x='Reporting_Period', y='Avg_Brittleness', color='Process_ID', markers=True)
        st.plotly_chart(fig, use_container_width=True)

    # Automation Velocity Impact over time by Project
    if "Automation_Velocity_Impact" in data:
        df = data["Automation_Velocity_Impact"]
        agg = df.groupby(['Automation_Project_ID', 'Reporting_Period']).agg(Avg_Savings=('Process_Automation_Savings_Hours', 'mean')).reset_index()
        st.subheader("Process Automation Savings Over Time")
        fig = px.line(agg, x='Reporting_Period', y='Avg_Savings', color='Automation_Project_ID', markers=True)
        st.plotly_chart(fig, use_container_width=True)

    # Additional aggregated charts can be added similarly...

else:
    st.info("Please upload the Excel file to begin.")
