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
.metric-clickable {
    cursor: pointer;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 15px;
    border-radius: 10px;
    color: white;
    text-align: center;
    transition: transform 0.2s;
}
.metric-clickable:hover {
    transform: scale(1.05);
}
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    file_path = "Updated_18_KPI_Dashboard.xlsx"
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

col1, col2 = st.columns([1, 5])
with col1:
    st.image("https://upload.wikimedia.org/wikipedia/commons/f/f5/Emblem_of_India.svg", width=80)
with col2:
    st.title("Employee KPI Dashboard")
    st.markdown("**Interactive Workforce Analytics** | April - September 2025")

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

with tab1:
    st.markdown('<div class="story-title">Executive Summary</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Productivity", f"{work_models['Productivity_Index'].mean():.1f}")
        if st.button("View Details", key="prod_detail"):
            st.session_state.selected = "productivity"
    
    with col2:
        st.metric("Burnout Risk", f"{burnout['Burnout_Risk_Score'].mean():.1f}")
        if st.button("View Details", key="burnout_detail"):
            st.session_state.selected = "burnout"
    
    with col3:
        st.metric("Skill Readiness", f"{skill_ready['Readiness_Score'].mean():.2f}/10")
        if st.button("View Details", key="skills_detail"):
            st.session_state.selected = "skills"
    
    with col4:
        st.metric("Security Risk", f"{shadow_it['Risk_Score'].mean():.1f}")
        if st.button("View Details", key="security_detail"):
            st.session_state.selected = "security"
    
    st.markdown("---")
    
    if "selected" not in st.session_state:
        st.session_state.selected = None
    
    if st.session_state.selected == "productivity":
        st.subheader("Productivity Deep Dive")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Average Productivity", f"{work_models['Productivity_Index'].mean():.2f}")
            st.metric("Min", f"{work_models['Productivity_Index'].min():.2f}")
            st.metric("Max", f"{work_models['Productivity_Index'].max():.2f}")
        with col2:
            work_model_stats = work_models.groupby("Work_Model")["Productivity_Index"].mean()
            st.bar_chart(work_model_stats)
    
    elif st.session_state.selected == "burnout":
        st.subheader("Burnout Risk Details")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Average Risk", f"{burnout['Burnout_Risk_Score'].mean():.2f}")
            at_risk = len(burnout[burnout['Burnout_Risk_Score'] > 5])
            st.metric("At-Risk Employees", at_risk)
        with col2:
            fig = px.histogram(burnout, x="Burnout_Risk_Score", nbins=20, color_discrete_sequence=["#e63946"])
            fig.add_vline(x=5, line_dash="dash")
            st.plotly_chart(fig, use_container_width=True)
    
    elif st.session_state.selected == "skills":
        st.subheader("Skills Development Details")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Readiness Index", f"{skill_ready['Readiness_Score'].mean():.2f}/10")
            st.metric("Training Completion", f"{skill_ready['Training_Completion_Percentage'].mean():.1f}%")
        with col2:
            skill_analysis = skill_gap.groupby("Skill_Category")["Skill_Gap_Score"].mean()
            st.bar_chart(skill_analysis)
    
    elif st.session_state.selected == "security":
        st.subheader("Security Risk Details")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Average Risk", f"{shadow_it['Risk_Score'].mean():.1f}")
            high_risk = len(shadow_it[shadow_it['Risk_Score'] > 60])
            st.metric("High-Risk Cases", f"{high_risk} ({high_risk/len(shadow_it)*100:.1f}%)")
        with col2:
            risk_sens = shadow_it.groupby("Data_Sensitivity_Level")["Risk_Score"].mean()
            st.bar_chart(risk_sens)
    
    if st.session_state.selected is None:
        st.subheader("Organizational Health")
        fig = go.Figure()
        categories = ["Productivity", "Wellbeing", "Skills", "Security"]
        current = [(work_models["Productivity_Index"].mean()/110)*100, wellbeing["Digital_Wellbeing_Score"].mean()*100, (skill_ready["Readiness_Score"].mean()/10)*100, ((100-shadow_it["Risk_Score"].mean())/100)*100]
        target = [85, 80, 70, 85]
        
        fig.add_trace(go.Scatterpolar(r=current, theta=categories, fill="toself", name="Current", line_color="#667eea"))
        fig.add_trace(go.Scatterpolar(r=target, theta=categories, fill="toself", name="Target", line_color="#f093fb", opacity=0.6))
        fig.update_layout(polar=dict(radialaxis=dict(range=[0,100])), showlegend=True)
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    st.subheader("Interactive Employee Analysis")
    
    selected_emp = st.selectbox("Select Employee", role_reality["Employee_ID"].unique(), key="emp_select_1")
    
    if selected_emp:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            low_val = role_reality[role_reality["Employee_ID"]==selected_emp]["Low_Value_Work_Percentage"].mean()*100
            st.metric(f"{selected_emp} - Low-Value %", f"{low_val:.1f}%")
        
        with col2:
            high_val = high_value[high_value["Employee_ID"]==selected_emp]["High_Value_Work_Percentage"].mean()*100
            st.metric(f"{selected_emp} - High-Value %", f"{high_val:.1f}%")
        
        with col3:
            prod = work_models[work_models["Employee_ID"]==selected_emp]["Productivity_Index"].mean()
            st.metric(f"{selected_emp} - Productivity", f"{prod:.1f}")
        
        emp_tabs = st.tabs(["Wellbeing", "Skills", "Security"])
        
        with emp_tabs[0]:
            burn = burnout[burnout["Employee_ID"]==selected_emp]["Burnout_Risk_Score"].mean()
            well = wellbeing[wellbeing["Employee_ID"]==selected_emp]["Digital_Wellbeing_Score"].mean()*100
            coll = collab[collab["Employee_ID"]==selected_emp]["Collaboration_Overload_Percentage"].mean()*100
            
            st.write(f"**Burnout Risk:** {burn:.2f}")
            st.write(f"**Wellbeing Score:** {well:.1f}%")
            st.write(f"**Collaboration Time:** {coll:.1f}%")
        
        with emp_tabs[1]:
            emp_skills = skill_gap[skill_gap["Employee_ID"]==selected_emp]
            if not emp_skills.empty:
                st.dataframe(emp_skills[["Skill_Category", "Performance_Score", "Benchmark_Score"]])
        
        with emp_tabs[2]:
            risk = shadow_it[shadow_it["Employee_ID"]==selected_emp]["Risk_Score"].mean()
            apps = shadow_it[shadow_it["Employee_ID"]==selected_emp]["Unauthorized_Apps_Count"].mean()
            st.write(f"**Risk Score:** {risk:.1f}")
            st.write(f"**Unauthorized Apps:** {apps:.1f}")

with tab2:
    st.markdown('<div class="story-title">Productivity Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Work Time Distribution")
        avg_low = role_reality["Low_Value_Work_Percentage"].mean() * 100
        avg_high = high_value["High_Value_Work_Percentage"].mean() * 100
        fig = go.Figure()
        fig.add_trace(go.Bar(x=["Low-Value", "High-Value"], y=[avg_low, avg_high], marker_color=["#e63946", "#06d6a0"]))
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("By Work Model")
        model_data = work_models.groupby("Work_Model")["Productivity_Index"].mean()
        fig = px.bar(x=model_data.index, y=model_data.values, color=model_data.values, color_continuous_scale="Viridis")
        st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("Trends Over Time")
    role_reality["Month"] = pd.to_datetime(role_reality["Reporting_Period"]).dt.to_period("M").astype(str)
    monthly = role_reality.groupby("Month")["Low_Value_Work_Percentage"].mean() * 100
    
    fig = px.line(x=monthly.index, y=monthly.values, markers=True, title="Low-Value Work Trend")
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("View Top Performers"):
        top_emp = work_models.groupby("Employee_ID")["Productivity_Index"].mean().nlargest(10)
        for idx, (emp, score) in enumerate(top_emp.items(), 1):
            st.write(f"{idx}. {emp}: {score:.2f}")
    
    with st.expander("Download Data"):
        csv = role_reality.to_csv(index=False)
        st.download_button("Download", csv, "productivity.csv")

with tab3:
    st.markdown('<div class="story-title">Wellbeing Assessment</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        cap = burnout["Capacity_Utilization_Percentage"].mean() * 100
        st.metric("Capacity Utilization", f"{cap:.1f}%")
    with col2:
        burn_avg = burnout["Burnout_Risk_Score"].mean()
        st.metric("Burnout Risk", f"{burn_avg:.1f}")
    with col3:
        coll_avg = collab["Collaboration_Overload_Percentage"].mean() * 100
        st.metric("Collaboration Time", f"{coll_avg:.1f}%")
    
    st.subheader("Risk Distribution")
    fig = px.histogram(burnout, x="Burnout_Risk_Score", nbins=30, color_discrete_sequence=["#e63946"])
    fig.add_vline(x=5, line_dash="dash", line_color="orange")
    st.plotly_chart(fig, use_container_width=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Wellbeing Trend")
        wellbeing["Month"] = pd.to_datetime(wellbeing["Reporting_Period"]).dt.to_period("M").astype(str)
        month_well = wellbeing.groupby("Month")["Digital_Wellbeing_Score"].mean() * 100
        fig = px.line(x=month_well.index, y=month_well.values, markers=True)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("Collaboration Distribution")
        coll_data = collab.groupby("Employee_ID")["Collaboration_Overload_Percentage"].mean() * 100
        fig = px.box(y=coll_data.values)
        st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("View At-Risk Employees"):
        at_risk = burnout[burnout["Burnout_Risk_Score"] > 5].groupby("Employee_ID")["Burnout_Risk_Score"].mean().sort_values(ascending=False).head(10)
        for emp, score in at_risk.items():
            st.warning(f"⚠️ {emp}: {score:.2f}")

with tab4:
    st.markdown('<div class="story-title">Skills Development</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Readiness Index", f"{skill_ready['Readiness_Score'].mean():.2f}/10")
    with col2:
        st.metric("Training Completion", f"{skill_ready['Training_Completion_Percentage'].mean():.1f}%")
    with col3:
        st.metric("Avg Skill Gap", f"{skill_gap['Skill_Gap_Score'].mean():.2f}")
    
    st.subheader("Skills vs Benchmark")
    skill_comp = skill_gap.groupby("Skill_Category")[["Performance_Score", "Benchmark_Score"]].mean()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Current", x=skill_comp.index, y=skill_comp["Performance_Score"], marker_color="#457b9d"))
    fig.add_trace(go.Bar(name="Target", x=skill_comp.index, y=skill_comp["Benchmark_Score"], marker_color="#06d6a0"))
    fig.update_layout(barmode="group")
    st.plotly_chart(fig, use_container_width=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Training by Quarter")
        train_q = skill_ready.groupby("Quarter")["Training_Completion_Percentage"].mean()
        fig = px.bar(x=train_q.index, y=train_q.values, color=train_q.values, color_continuous_scale="Blues")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("Readiness Distribution")
        fig = px.histogram(skill_ready, x="Readiness_Score", nbins=20, color_discrete_sequence=["#e63946"])
        st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("Priority Development"):
        low = skill_ready[skill_ready["Readiness_Score"] < 2].groupby("Employee_ID")["Readiness_Score"].mean().sort_values().head(10)
        for emp, score in low.items():
            st.warning(f"{emp}: {score:.2f}/10")

with tab5:
    st.markdown('<div class="story-title">Security Assessment</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Avg Risk", f"{shadow_it['Risk_Score'].mean():.1f}")
    with col2:
        high = len(shadow_it[shadow_it["Risk_Score"] > 60])
        st.metric("High-Risk", f"{high} ({high/len(shadow_it)*100:.1f}%)")
    with col3:
        st.metric("Unauthorized Apps", f"{shadow_it['Unauthorized_Apps_Count'].mean():.1f}")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Risk Distribution")
        fig = px.histogram(shadow_it, x="Risk_Score", nbins=30, color_discrete_sequence=["#e63946"])
        fig.add_vline(x=60, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("Risk by Data Sensitivity")
        risk_sens = shadow_it.groupby("Data_Sensitivity_Level")["Risk_Score"].mean()
        fig = px.bar(x=risk_sens.index, y=risk_sens.values, color=risk_sens.values, color_continuous_scale="Reds")
        st.plotly_chart(fig, use_container_width=True)
    
    shadow_it["Month"] = pd.to_datetime(shadow_it["Week_Ending_Date"]).dt.to_period("M").astype(str)
    monthly_apps = shadow_it.groupby("Month")["Unauthorized_Apps_Count"].mean()
    
    fig = px.line(x=monthly_apps.index, y=monthly_apps.values, markers=True, title="Unauthorized Apps Trend")
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("High-Risk Employees"):
        high_risk_emp = shadow_it[shadow_it["Risk_Score"] > 60].groupby("Employee_ID")["Risk_Score"].mean().sort_values(ascending=False).head(10)
        for emp, score in high_risk_emp.items():
            st.error(f"🚨 {emp}: {score:.1f}")

st.markdown("---")
st.markdown("<div style='text-align:center;color:#666;padding:20px'><strong>Employee KPI Dashboard</strong><br>Workforce Analytics | April-September 2025</div>", unsafe_allow_html=True)
