import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.main {background-color: #f8f9fa;}
.stApp {background-color: #f8f9fa;}
.narrative-box {
    background-color: #ffffff;
    padding: 20px;
    border-radius: 10px;
    border-left: 4px solid #2E86AB;
    margin: 20px 0;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
.story-title {
    color: #2E86AB;
    font-size: 1.8em;
    font-weight: 700;
    margin: 30px 0 15px 0;
}
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    file_path = "Enhanced_25_Employee_KPI_Dashboard.xlsx"
    sheets = [
        "Role_vs_Reality_Analysis",
        "Hidden_Capacity_Burnout_Risk",
        "Work_Models_Effectiveness",
        "Digital_Collaboration_Overload",
        "Digital_Wellbeing_Index",
        "Data_Driven_Skill_Gap_Analysis",
        "High_Value_Work_Ratio",
        "Future_Skill_Readiness_Index",
        "Shadow_IT_Risk_Score"
    ]
    
    all_data = {}
    for sheet in sheets:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            all_data[sheet] = df
        except Exception as e:
            st.error(f"Could not load {sheet}: {e}")
    return all_data

data = load_data()

st.title("Employee KPI Dashboard")
st.markdown("**Workforce Analytics** | April - September 2025")
st.markdown("---")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Summary", 
    "💼 Productivity", 
    "🧘 Wellbeing", 
    "📚 Skills", 
    "🔒 Security"
])

role_reality = data["Role_vs_Reality_Analysis"]
high_value = data["High_Value_Work_Ratio"]
work_models = data["Work_Models_Effectiveness"]
burnout = data["Hidden_Capacity_Burnout_Risk"]
wellbeing = data["Digital_Wellbeing_Index"]
collab = data["Digital_Collaboration_Overload"]
skill_gap = data["Data_Driven_Skill_Gap_Analysis"]
skill_ready = data["Future_Skill_Readiness_Index"]
shadow_it = data["Shadow_IT_Risk_Score"]

# MONOCHROME COLOR PALETTES
mono_greys = ['#2c3e50', '#34495e', '#7f8c8d', '#95a5a6', '#bdc3c7', '#ecf0f1']
mono_blues = ['#0f1f3f', '#1a3a52', '#2d5a6d', '#5a7f94', '#8fa9be', '#c5d9e8']

with tab1:
    st.markdown('<div class="story-title">Executive Summary</div>', unsafe_allow_html=True)
    
    # Calculate metrics with scales
    productivity = work_models['Productivity_Index'].mean()
    burnout_score = burnout['Burnout_Risk_Score'].mean()
    skill_readiness = skill_ready['Readiness_Score'].mean()
    security_risk = shadow_it['Risk_Score'].mean()
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Productivity Index", f"{productivity:.1f}%")
    
    with col2:
        st.metric("Burnout Risk", f"{burnout_score:.1f}/10")
    
    with col3:
        st.metric("Skill Readiness", f"{skill_readiness:.2f}/10")
    
    with col4:
        st.metric("Security Risk", f"{security_risk:.1f}%")
        
       
    st.markdown("---")
    st.subheader("Organizational Health Overview")
    
    # Create closed loop for radar chart so all connecting lines are visible
categories = ["Productivity", "Security", "Wellbeing", "Skills"]
current = [
    (productivity),
    ((100 - security_risk) / 100) * 100,
    wellbeing["Digital_Wellbeing_Score"].mean() * 100,
    (skill_readiness / 10) * 100
]
target = [100, 85, 80, 70]

# Close the polygon by appending the first value and first label to the end
categories_closed = categories + [categories[0]]
current_closed = current + [current[0]]
target_closed = target + [target[0]]

fig = go.Figure()

fig.add_trace(go.Scatterpolar(
    r=current_closed,
    theta=categories_closed,
    fill='toself',
    name='Current',
    line=dict(color=mono_blues[0], width=2),
    fillcolor=mono_blues[4]
))
fig.add_trace(go.Scatterpolar(
    r=target_closed,
    theta=categories_closed,
    fill='toself',
    name='Target',
    line=dict(color=mono_blues[2], width=2),
    fillcolor=mono_greys[5],
    opacity=0.5
))
fig.update_layout(
    polar=dict(
        radialaxis=dict(range=[0, 120], tickfont=dict(size=10)),
        angularaxis=dict(tickfont=dict(size=11))
    ),
    showlegend=True,
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color=mono_greys[0]),
    height=500
)

st.plotly_chart(fig, use_container_width=True)


with tab2:
    st.markdown('<div class="story-title">Productivity Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Work Time Distribution")
        avg_low = role_reality["Low_Value_Work_Percentage"].mean() * 100
        avg_high = high_value["High_Value_Work_Percentage"].mean() * 100
        fig = go.Figure()
        fig.add_trace(go.Bar(x=["Low-Value", "High-Value"], y=[avg_low, avg_high], 
                            marker_color=[mono_greys[2], mono_blues[0]]))
        fig.update_layout(yaxis_title="% of Work Time", showlegend=False,
                         paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("Productivity by Work Model (Dynamic Scale)")
        model_data = work_models.groupby("Work_Model")["Productivity_Index"].mean()
        min_val = model_data.min()
        max_val = model_data.max()
        
        fig = px.bar(x=model_data.index, y=model_data.values)
        fig.update_traces(marker_color=mono_blues[0])
        fig.update_yaxes(range=[min_val * 0.95, max_val * 1.05])
        fig.update_layout(yaxis_title="Productivity Index", showlegend=False,
                         paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
    
    # Quartile selection for productivity
    st.subheader("Employee Performance Quartiles")
    
    emp_prod = work_models.groupby("Employee_ID")["Productivity_Index"].mean()
    q1, q2, q3 = emp_prod.quantile([0.25, 0.5, 0.75])
    
    quartile_options = ["Top Quartile (Q4)", "Second Quartile (Q3)", "Third Quartile (Q2)", "Bottom Quartile (Q1)"]
    selected_quartile = st.selectbox("Select Performance Group:", quartile_options)
    
    if selected_quartile == "Top Quartile (Q4)":
        selected_emps = emp_prod[emp_prod >= q3].index.tolist()
        title_text = f"Top Performers (n={len(selected_emps)})"
    elif selected_quartile == "Second Quartile (Q3)":
        selected_emps = emp_prod[(emp_prod >= q2) & (emp_prod < q3)].index.tolist()
        title_text = f"Second Quartile (n={len(selected_emps)})"
    elif selected_quartile == "Third Quartile (Q2)":
        selected_emps = emp_prod[(emp_prod >= q1) & (emp_prod < q2)].index.tolist()
        title_text = f"Third Quartile (n={len(selected_emps)})"
    elif selected_quartile == "Bottom Quartile (Q1)":
        selected_emps = emp_prod[emp_prod < q1].index.tolist()
        title_text = f"Bottom Quartile (n={len(selected_emps)})"
    else:
        selected_emps = emp_prod.index.tolist()
        title_text = "All Employees"
    
    if selected_emps:
        quartile_data = emp_prod[emp_prod.index.isin(selected_emps)].sort_values(ascending=False)
        fig = px.bar(y=quartile_data.index, x=quartile_data.values, orientation='h', title=title_text)
        fig.update_traces(marker_color=mono_blues[2])
        fig.update_layout(yaxis_title="Employee", xaxis_title="Productivity Index",
                         paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("Low-Value Work Trend")
    role_reality["Month"] = pd.to_datetime(role_reality["Reporting_Period"]).dt.strftime('%Y-%m')
    monthly = role_reality.groupby("Month")["Low_Value_Work_Percentage"].mean() * 100
    
    fig = px.line(x=monthly.index, y=monthly.values, markers=True, title="Low-Value Work Trend")
    fig.update_traces(line=dict(color=mono_greys[2], width=3), marker=dict(size=8))
    fig.update_layout(yaxis_title="% of Work Time", xaxis_title="Month",
                     paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("View Top Performers"):
        top_emp = work_models.groupby("Employee_ID")["Productivity_Index"].mean().nlargest(10)
        for idx, (emp, score) in enumerate(top_emp.items(), 1):
            st.write(f"{idx}. {emp}: {score:.2f}")
    
    with st.expander("View Bottom Performers"):
        bottom_emp = work_models.groupby("Employee_ID")["Productivity_Index"].mean().nsmallest(10)
        for idx, (emp, score) in enumerate(bottom_emp.items(), 1):
            st.write(f"{idx}. {emp}: {score:.2f}")

with tab3:
    st.markdown('<div class="story-title">Wellbeing Assessment</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        cap = burnout["Capacity_Utilization_Percentage"].mean() * 100
        st.metric("Capacity Utilization", f"{cap:.1f}%")
    with col2:
        burn_avg = burnout["Burnout_Risk_Score"].mean()
        st.metric("Burnout Risk", f"{burn_avg:.1f}/10")
    with col3:
        coll_avg = collab["Collaboration_Overload_Percentage"].mean() * 100
        st.metric("Collaboration Time", f"{coll_avg:.1f}%")
    
    st.markdown("---")
    
    # Month selector for burnout distribution
    burnout["Month"] = pd.to_datetime(burnout["Week_Ending_Date"]).dt.strftime('%Y-%m')
    available_months = sorted(burnout["Month"].unique())
    selected_month = st.selectbox("Select Month for Distribution:", available_months, index=len(available_months)-1, key="burnout_month")
    
    month_data = burnout[burnout["Month"] == selected_month]["Burnout_Risk_Score"]
    
    st.subheader(f"Burnout Risk Distribution - {selected_month}")
    fig = px.histogram(month_data, nbins=15, title=f"Distribution ({selected_month})")
    fig.update_traces(marker_color=mono_greys[2])
    fig.update_layout(xaxis_title="Burnout Risk Score", yaxis_title="Frequency",
                     showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    fig.add_vline(x=5, line_dash="dash", line_color=mono_blues[0], annotation_text="High Risk Threshold")
    st.plotly_chart(fig, use_container_width=True)
    
    # Burnout trend
    st.subheader("Average Burnout Risk Trend")
    burnout_trend = burnout.groupby("Month")["Burnout_Risk_Score"].mean()
    
    fig = px.line(x=burnout_trend.index, y=burnout_trend.values, markers=True, title="Monthly Average Burnout Risk")
    fig.update_traces(line=dict(color=mono_greys[2], width=3), marker=dict(size=8))
    fig.update_layout(yaxis_title="Burnout Risk Score", xaxis_title="Month",
                     paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("View At-Risk Employees"):
        at_risk = burnout[burnout["Burnout_Risk_Score"] > 5].groupby("Employee_ID")["Burnout_Risk_Score"].mean().sort_values(ascending=False).head(10)
        for emp, score in at_risk.items():
            st.warning(f"⚠️ {emp}: {score:.2f}/10")

with tab4:
    st.markdown('<div class="story-title">Skills Development</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Readiness Index", f"{skill_ready['Readiness_Score'].mean():.2f}/10")
    with col2:
        st.metric("Training Completion", f"{skill_ready['Training_Completion_Percentage'].mean():.1f}%")
    with col3:
        st.metric("Avg Skill Gap", f"{skill_gap['Skill_Gap_Score'].mean():.2f}")
    
    st.markdown("---")
    
    st.subheader("Skills vs Benchmark")
    skill_comp = skill_gap.groupby("Skill_Category")[["Performance_Score", "Benchmark_Score"]].mean()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Current", x=skill_comp.index, y=skill_comp["Performance_Score"], marker_color=mono_blues[1]))
    fig.add_trace(go.Bar(name="Target", x=skill_comp.index, y=skill_comp["Benchmark_Score"], marker_color=mono_greys[2]))
    fig.update_layout(barmode="group", yaxis_title="Score", 
                     paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)
    
    # Quarter selector for readiness distribution
    skill_ready_quarters = skill_ready['Quarter'].unique()
    selected_quarter = st.selectbox("Select Quarter for Distribution:", sorted(skill_ready_quarters), key="readiness_quarter")
    
    quarter_data = skill_ready[skill_ready["Quarter"] == selected_quarter]["Readiness_Score"]
    
    st.subheader(f"Readiness Score Distribution - {selected_quarter}")
    fig = px.histogram(quarter_data, nbins=10, title=f"Distribution ({selected_quarter})")
    fig.update_traces(marker_color=mono_blues[2])
    fig.update_layout(xaxis_title="Readiness Score", yaxis_title="Frequency",
                     showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    fig.add_vline(x=5, line_dash="dash", line_color=mono_greys[0], annotation_text="Target")
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("Priority Development"):
        low = skill_ready[skill_ready["Readiness_Score"] < 5].groupby("Employee_ID")["Readiness_Score"].mean().sort_values().head(10)
        for emp, score in low.items():
            st.warning(f"{emp}: {score:.2f}/10")

with tab5:
    st.markdown('<div class="story-title">Security Assessment</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Avg Risk", f"{shadow_it['Risk_Score'].mean():.1f}%")
    with col2:
        high = len(shadow_it[shadow_it["Risk_Score"] > 60])
        st.metric("High-Risk Cases", f"{high} ({high/len(shadow_it)*100:.1f}%)")
    with col3:
        st.metric("Unauthorized Apps", f"{shadow_it['Unauthorized_Apps_Count'].mean():.1f}")
    
    st.markdown("---")
    
    # Month selector for risk distribution
    shadow_it["Month"] = pd.to_datetime(shadow_it["Week_Ending_Date"]).dt.strftime('%Y-%m')
    available_months_risk = sorted(shadow_it["Month"].unique())
    selected_month_risk = st.selectbox("Select Month for Distribution:", available_months_risk, index=len(available_months_risk)-1, key="risk_month")
    
    month_risk_data = shadow_it[shadow_it["Month"] == selected_month_risk]["Risk_Score"]
    
    st.subheader(f"Risk Distribution - {selected_month_risk}")
    fig = px.histogram(month_risk_data, nbins=15, title=f"Distribution ({selected_month_risk})")
    fig.update_traces(marker_color=mono_greys[2])
    fig.update_layout(xaxis_title="Risk Score", yaxis_title="Frequency",
                     showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    fig.add_vline(x=60, line_dash="dash", line_color=mono_blues[0], annotation_text="High Risk Threshold")
    st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("Risk by Data Sensitivity")
    risk_sens = shadow_it.groupby("Data_Sensitivity_Level")["Risk_Score"].mean()
    fig = px.bar(x=risk_sens.index, y=risk_sens.values)
    fig.update_traces(marker_color=mono_greys[2])
    fig.update_layout(yaxis_title="Avg Risk Score", xaxis_title="Data Sensitivity Level",
                     paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)
    
    shadow_it_trend = shadow_it.groupby("Month")["Unauthorized_Apps_Count"].mean()
    
    st.subheader("Unauthorized Apps Trend")
    fig = px.line(x=shadow_it_trend.index, y=shadow_it_trend.values, markers=True, title="Monthly Average Unauthorized Apps")
    fig.update_traces(line=dict(color=mono_greys[2], width=3), marker=dict(size=8))
    fig.update_layout(yaxis_title="Avg Count", xaxis_title="Month",
                     paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("High-Risk Employees"):
        high_risk_emp = shadow_it[shadow_it["Risk_Score"] > 60].groupby("Employee_ID")["Risk_Score"].mean().sort_values(ascending=False).head(10)
        for emp, score in high_risk_emp.items():
            st.error(f"🚨 {emp}: {score:.1f}")

st.markdown("---")
st.markdown("<div style='text-align:center;color:#666;padding:20px'><strong>Employee KPI Dashboard</strong><br>Workforce Analytics | April-September 2025</div>", unsafe_allow_html=True)
