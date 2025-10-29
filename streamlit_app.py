import streamlit as st
import pandas as pd
import plotly.express as px

# Set page config
st.set_page_config(layout="wide", page_title="Summary Organizational Analytics Dashboard")

@st.cache_data
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    return {sheet: xls.parse(sheet) for sheet in sheets}

uploaded_file = st.file_uploader("Upload Excel file (Updated_18_KPI_Dashboard.xlsx)", type=["xlsx"])

if uploaded_file:
    data = load_data(uploaded_file)
    
    # Sidebar filters
    st.sidebar.header("Filters")

    # Department filter (if available)
    department_options = None
    if 'Role_vs_Reality_Analysis' in data:
        role_reality_df = data['Role_vs_Reality_Analysis']
        department_options = role_reality_df['Department'].unique().tolist() if 'Department' in role_reality_df.columns else None
        selected_depts = st.sidebar.multiselect("Select Departments", options=department_options, default=department_options)
    else:
        selected_depts = None

    # Reporting Period filter (month/quarter/year)
    all_periods = None
    if 'Role_vs_Reality_Analysis' in data:
        all_periods = sorted(role_reality_df['Reporting_Period'].dropna().unique())
        selected_periods = st.sidebar.multiselect("Select Reporting Periods", options=all_periods, default=all_periods)
    else:
        selected_periods = None

    st.title("Summary Organizational Analytics Dashboard")

    st.header("Executive Summary")
    try:
        # Burnout Risk summary (aggregated)
        burnout_df = data.get("Hidden_Capacity_Burnout_Risk")
        if burnout_df is not None:
            if 'Department' in burnout_df.columns and selected_depts:
                burnout_df = burnout_df[burnout_df['Department'].isin(selected_depts)]
            avg_burnout = burnout_df['Burnout_Risk_Score'].mean()
            st.markdown(f"- Average Burnout Risk Score across data: **{avg_burnout:.2f}**")

        # High Value Work Ratio summary
        hv_work_df = data.get("High_Value_Work_Ratio")
        if hv_work_df is not None:
            if 'Department' in hv_work_df.columns and selected_depts:
                hv_work_df = hv_work_df[hv_work_df['Department'].isin(selected_depts)]
            avg_high_value_pct = hv_work_df['High_Value_Work_Percentage'].mean() * 100
            st.markdown(f"- Average High Value Work Percentage: **{avg_high_value_pct:.1f}%**")

    except Exception as e:
        st.warning("Error in Executive Summary calculation: " + str(e))

    st.markdown("---")

    # Role vs Reality: Summarize average low-value work by Department and Reporting Period
    if 'Role_vs_Reality_Analysis' in data:
        df = data['Role_vs_Reality_Analysis']
        if 'Department' in df.columns and selected_depts:
            df = df[df['Department'].isin(selected_depts)]
        if 'Reporting_Period' in df.columns and selected_periods:
            df = df[df['Reporting_Period'].isin(selected_periods)]

        agg_df = df.groupby(['Department', 'Reporting_Period']).agg(
            Avg_Low_Value_Work_Hours=('Time_Low_Value_Tasks_Hours', 'mean'),
            Avg_Low_Value_Work_Pct=('Low_Value_Work_Percentage', 'mean')
        ).reset_index()

        st.subheader("Average Low Value Work by Department and Period")
        fig = px.bar(agg_df, x='Reporting_Period', y='Avg_Low_Value_Work_Pct', color='Department',
                     labels={'Avg_Low_Value_Work_Pct': 'Avg % Low Value Work', 'Reporting_Period': 'Period'},
                     barmode='group')
        st.plotly_chart(fig, use_container_width=True)

    # Burnout Risk by Department and Quarter
    if 'Hidden_Capacity_Burnout_Risk' in data:
        df = data['Hidden_Capacity_Burnout_Risk']
        if 'Department' in df.columns and selected_depts:
            df = df[df['Department'].isin(selected_depts)]
        if 'Quarter' in df.columns:
            agg_df = df.groupby(['Department', 'Quarter']).agg(
                Avg_Burnout_Risk=('Burnout_Risk_Score', 'mean'),
                Max_Burnout_Risk=('Burnout_Risk_Score', 'max')
            ).reset_index()
            st.subheader("Average Burnout Risk by Department and Quarter")
            fig = px.line(agg_df, x='Quarter', y='Avg_Burnout_Risk', color='Department',
                          labels={"Avg_Burnout_Risk": "Avg Burnout Risk Score"},
                          markers=True)
            st.plotly_chart(fig, use_container_width=True)

    # High Value Work Percentage over time (Department level)
    if 'High_Value_Work_Ratio' in data:
        df = data['High_Value_Work_Ratio']
        if 'Department' in df.columns and selected_depts:
            df = df[df['Department'].isin(selected_depts)]
        if 'Reporting_Period' in df.columns and selected_periods:
            df = df[df['Reporting_Period'].isin(selected_periods)]
        agg_df = df.groupby(['Department', 'Reporting_Period']).agg(
            Avg_HV_Work_Pct = ('High_Value_Work_Percentage', 'mean')
        ).reset_index()
        st.subheader("Average High Value Work Percentage by Department and Period")
        fig = px.line(agg_df, x='Reporting_Period', y='Avg_HV_Work_Pct', color='Department', markers=True,
                      labels={"Avg_HV_Work_Pct": "Avg High Value Work %"})

        st.plotly_chart(fig, use_container_width=True)
    
    # Digital Wellbeing Score by Department and Period
    if 'Digital_Wellbeing_Index' in data:
        df = data['Digital_Wellbeing_Index']
        if 'Department' in df.columns and selected_depts:
            df = df[df['Department'].isin(selected_depts)]
        if 'Reporting_Period' in df.columns and selected_periods:
            df = df[df['Reporting_Period'].isin(selected_periods)]
        agg_df = df.groupby(['Department', 'Reporting_Period']).agg(
            Avg_DW_Score=('Digital_Wellbeing_Score', 'mean')
        ).reset_index()
        st.subheader("Digital Wellbeing Score by Department and Period")
        fig = px.line(agg_df, x='Reporting_Period', y='Avg_DW_Score', color='Department', markers=True)
        st.plotly_chart(fig, use_container_width=True)

    # Organizational Agility by Department and Quarter
    if 'Organizational_Agility_Score' in data:
        df = data['Organizational_Agility_Score']
        if 'Department' in df.columns and selected_depts:
            df = df[df['Department'].isin(selected_depts)]
        if 'Quarter' in df.columns:
            agg_df = df.groupby(['Department', 'Quarter']).agg(
                Avg_Agility_Score=('Agility_Score', 'mean')
            ).reset_index()
            st.subheader("Organizational Agility Score by Department and Quarter")
            fig = px.bar(agg_df, x='Quarter', y='Avg_Agility_Score', color='Department', barmode="group")
            st.plotly_chart(fig, use_container_width=True)

    # Skill Gap Average by Skill Category
    if 'Data_Driven_Skill_Gap_Analysis' in data:
        df = data['Data_Driven_Skill_Gap_Analysis']
        if 'Skill_Category' in df.columns:
            agg_df = df.groupby('Skill_Category').agg(
                Avg_Skill_Gap_Score=('Skill_Gap_Score', 'mean')
            ).reset_index()
            st.subheader("Average Skill Gap Score by Skill Category")
            fig = px.bar(agg_df, x='Skill_Category', y='Avg_Skill_Gap_Score')
            st.plotly_chart(fig, use_container_width=True)
    
    # Process Brittleness Over Time
    if 'Process_Brittleness_Index' in data:
        df = data['Process_Brittleness_Index']
        if 'Process_ID' in df.columns and 'Reporting_Period' in df.columns:
            agg_df = df.groupby(['Process_ID', 'Reporting_Period']).agg(
                Avg_Brittleness_Percentage = ('Brittleness_Percentage', 'mean')
            ).reset_index()
            st.subheader("Average Process Brittleness Over Time")
            fig = px.line(agg_df, x='Reporting_Period', y='Avg_Brittleness_Percentage', color='Process_ID',
                          markers=True)
            st.plotly_chart(fig, use_container_width=True)

    # Automation Impact Over Time
    if 'Automation_Velocity_Impact' in data:
        df = data['Automation_Velocity_Impact']
        if 'Automation_Project_ID' in df.columns and 'Reporting_Period' in df.columns:
            agg_df = df.groupby(['Automation_Project_ID', 'Reporting_Period']).agg(
                Avg_Savings_Hours = ('Process_Automation_Savings_Hours', 'mean')
            ).reset_index()
            st.subheader("Average Automation Savings Hours Over Time")
            fig = px.line(agg_df, x='Reporting_Period', y='Avg_Savings_Hours', color='Automation_Project_ID',
                          markers=True)
            st.plotly_chart(fig, use_container_width=True)
    
    # Additional aggregated charts can be added similarly...

else:
    st.info("Please upload the Excel file to begin.")

