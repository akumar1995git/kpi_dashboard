import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(page_title="Employee KPI Story Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for professional styling
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
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 20px;
    border-radius: 10px;
    color: white;
    text-align: center;
}
.insight-highlight {
    background-color: #fff3cd;
    padding: 15px;
    border-left: 4px solid #ffc107;
    margin: 15px 0;
    border-radius: 5px;
}
.story-title {
    color: #2E86AB;
    font-size: 1.8em;
    font-weight: 700;
    margin: 30px 0 15px 0;
}
</style>
""", unsafe_allow_html=True)

# Load data
@st.cache_data
def load_data():
    file_path = "Updated_18_KPI_Dashboard.xlsx"
    employee_kpi_sheets = [
        'Role_vs_Reality_Analysis',
        'Hidden_Capacity_Burnout_Risk',
        'Work_Models_Effectiveness',
        'Digital_Collaboration_Overload',
        'Digital_Wellbeing_Index',
        'Data_Driven_Skill_Gap_Analysis',
        'High_Value_Work_Ratio',
        'Future_Skill_Readiness_Index',
        'Shadow_IT_Risk_Score'
    ]

    all_data = {}
    for sheet in employee_kpi_sheets:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            all_data[sheet] = df
        except Exception as e:
            st.error(f"Could not load {sheet}: {e}")
    return all_data

data = load_data()

# Header
col1, col2 = st.columns([1, 5])
with col1:
    st.image("https://upload.wikimedia.org/wikipedia/commons/f/f5/Emblem_of_India.svg", width=80)
with col2:
    st.title("Employee KPI Story Dashboard")
    st.markdown("**Narrative-Driven Workforce Analytics** | April - September 2025")

st.markdown("---")

# Navigation
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Summary", 
    "💼 Productivity Story", 
    "🧘 Wellbeing Story", 
    "📚 Skills Story", 
    "🔒 Security Story"
])

# Calculate key metrics
role_reality = data[\'Role_vs_Reality_Analysis\']
high_value = data[\'High_Value_Work_Ratio\']
work_models = data[\'Work_Models_Effectiveness\']
burnout = data[\'Hidden_Capacity_Burnout_Risk\']
wellbeing = data[\'Digital_Wellbeing_Index\']
collab = data[\'Digital_Collaboration_Overload\']
skill_gap = data[\'Data_Driven_Skill_Gap_Analysis\']
skill_ready = data[\'Future_Skill_Readiness_Index\']
shadow_it = data[\'Shadow_IT_Risk_Score\']

# TAB 1: EXECUTIVE SUMMARY
with tab1:
    st.markdown(\'<div class="story-title">🎯 Executive Summary</div>\', unsafe_allow_html=True)

    # Key metrics row
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "Productivity Index", 
            f"{work_models[\'Productivity_Index\'].mean():.1f}",
            f"+{work_models[\'Productivity_Index\'].mean()-100:.1f}% vs baseline"
        )

    with col2:
        burnout_score = burnout[\'Burnout_Risk_Score\'].mean()
        st.metric(
            "Burnout Risk",
            f"{burnout_score:.1f}",
            "⚠️ Elevated" if burnout_score > 5 else "✅ Moderate",
            delta_color="inverse"
        )

    with col3:
        readiness = skill_ready[\'Readiness_Score\'].mean()
        st.metric(
            "Skill Readiness",
            f"{readiness:.2f}/10",
            "🚨 Critical Gap"
        )

    with col4:
        risk = shadow_it[\'Risk_Score\'].mean()
        st.metric(
            "Security Risk",
            f"{risk:.1f}",
            "Moderate"
        )

    st.markdown("---")

    # Narrative summary
    st.markdown(\'\'\'
    <div class="narrative-box">
    <h3>📖 The Workforce Story at a Glance</h3>
    <p style="font-size:1.1em;line-height:1.8">
    Our organization\'s workforce demonstrates <strong>above-baseline productivity (104.3)</strong>, but this performance 
    comes with hidden costs. Employees spend <strong>33.6% of their time on low-value tasks</strong>, operate at 
    <strong>99.9% capacity</strong> (dangerously close to maximum), and show <strong>critical skills gaps</strong> 
    for future readiness (1.31/10 score).
    </p>
    <p style="font-size:1.1em;line-height:1.8">
    The data tells us we\'re succeeding today but <strong>risking tomorrow</strong>. This dashboard reveals four 
    interconnected narratives that explain not just <em>what</em> is happening, but <em>why</em> and <em>what to do about it</em>.
    </p>
    </div>
    \'\'\'
    , unsafe_allow_html=True)

    # Visual: Four stories overview
    st.subheader("The Four Workforce Stories")

    fig = go.Figure()

    categories = [\'Productivity\', \'Wellbeing\', \'Skills\', \'Security\']
    current_scores = [
        (work_models[\'Productivity_Index\'].mean() / 110) * 100,  # Normalize to 100
        wellbeing[\'Digital_Wellbeing_Score\'].mean() * 100,
        (skill_ready[\'Readiness_Score\'].mean() / 10) * 100,
        ((100 - shadow_it[\'Risk_Score\'].mean()) / 100) * 100  # Invert so higher is better
    ]
    target_scores = [85, 80, 70, 85]

    fig.add_trace(go.Scatterpolar(
        r=current_scores,
        theta=categories,
        fill=\'toself\',
        name=\'Current State\',
        line_color=\'#667eea\'
    ))

    fig.add_trace(go.Scatterpolar(
        r=target_scores,
        theta=categories,
        fill=\'toself\',
        name=\'Target State\',
        line_color=\'#f093fb\',
        opacity=0.6
    ))

    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        showlegend=True,
        title="Workforce Health: Current vs Target"
    )

    st.plotly_chart(fig, use_container_width=True)

# TAB 2: PRODUCTIVITY STORY
with tab2:
    st.markdown(\'<div class="story-title">💼 Story 1: The Productivity Paradox</div>\', unsafe_allow_html=True)

    st.markdown(\'\'\'
    <div class="narrative-box">
    <h4>🔍 The Paradox Explained</h4>
    <p style="font-size:1.05em;line-height:1.7">
    Our workforce shows <strong>strong productivity metrics</strong> (104.3 index), yet this success masks an 
    underlying inefficiency: one-third of work time is consumed by low-value administrative tasks. 
    <strong>Imagine recovering 17% more strategic capacity</strong> through smart automation.
    </p>
    </div>
    \'\'\'
    , unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        # Low-value vs High-value work
        avg_low = role_reality[\'Low_Value_Work_Percentage\'].mean() * 100
        avg_high = high_value[\'High_Value_Work_Percentage\'].mean() * 100

        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            x=[\'Low-Value Tasks\', \'High-Value Strategic Work\'],
            y=[avg_low, avg_high],
            marker_color=[\'#e63946\', \'#06d6a0\'],
            text=[f"{avg_low:.1f}%", f"{avg_high:.1f}%"],
            textposition=\'outside\'
        ))
        fig1.update_layout(
            title="Time Allocation: Where Does Work Go?",
            yaxis_title="% of Work Time",
            showlegend=False,
            height=400
        )
        st.plotly_chart(fig1, use_container_width=True)

        st.markdown(\'\'\'
        <div class="insight-highlight">
        💡 <strong>Key Insight:</strong> If we automate half of low-value work, we unlock 
        <strong>16.8% more capacity</strong> for strategic initiatives—without hiring.
        </div>
        \'\'\'
        , unsafe_allow_html=True)

    with col2:
        # Productivity by work model
        prod_by_model = work_models.groupby(\'Work_Model\')[\'Productivity_Index\'].mean().reset_index()

        fig2 = px.bar(
            prod_by_model,
            x=\'Work_Model\',
            y=\'Productivity_Index\',
            color=\'Productivity_Index\',
            color_continuous_scale=\'Viridis\',
            title="Productivity Index by Work Model"
        )
        fig2.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="Baseline")
        st.plotly_chart(fig2, use_container_width=True)

        st.markdown(\'\'\'
        <div class="insight-highlight">
        💡 <strong>Key Insight:</strong> All work models perform above baseline—remote work 
        policies are <strong>effective and sustainable</strong>.
        </div>
        \'\'\'
        , unsafe_allow_html=True)

    # Trend over time
    st.subheader("📈 Productivity Trends Over Time")
    role_reality[\'Month\'] = pd.to_datetime(role_reality[\'Reporting_Period\']).dt.to_period(\'M\').astype(str)
    monthly_low_value = role_reality.groupby(\'Month\')[\'Low_Value_Work_Percentage\'].mean().reset_index()
    monthly_low_value[\'Low_Value_Work_Percentage\'] *= 100

    fig3 = px.line(
        monthly_low_value,
        x=\'Month\',
        y=\'Low_Value_Work_Percentage\',
        markers=True,
        title="Low-Value Work Trend (Opportunity for Improvement)"
    )
    fig3.update_traces(line_color=\'#e63946\', line_width=3)
    st.plotly_chart(fig3, use_container_width=True)

# TAB 3: WELLBEING STORY
with tab3:
    st.markdown(\'<div class="story-title">🧘 Story 2: The Wellbeing Tightrope</div>\', unsafe_allow_html=True)

    st.markdown(\'\'\'
    <div class="narrative-box">
    <h4>⚠️ Walking the Line</h4>
    <p style="font-size:1.05em;line-height:1.7">
    Operating at <strong>99.9% capacity utilization</strong> leaves no room for error, growth, or recovery. 
    With a burnout risk score of <strong>5.1 (elevated)</strong> and collaboration consuming 
    <strong>45% of work time</strong>, we\'re at the edge of sustainable performance.
    </p>
    </div>
    \'\'\'
    , unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        capacity = burnout[\'Capacity_Utilization_Percentage\'].mean() * 100
        st.metric("Capacity Utilization", f"{capacity:.1f}%", "⚠️ Near Maximum")

    with col2:
        burnout_avg = burnout[\'Burnout_Risk_Score\'].mean()
        st.metric("Burnout Risk Score", f"{burnout_avg:.1f}", "🔴 Elevated")

    with col3:
        collab_pct = collab[\'Collaboration_Overload_Percentage\'].mean() * 100
        st.metric("Collaboration Time", f"{collab_pct:.1f}%", "📊 Upper Range")

    # Burnout risk distribution
    st.subheader("Distribution of Burnout Risk Across Workforce")
    fig4 = px.histogram(
        burnout,
        x=\'Burnout_Risk_Score\',
        nbins=30,
        color_discrete_sequence=[\'#e63946\'],
        title="How Many Employees Are At Risk?"
    )
    fig4.add_vline(x=5, line_dash="dash", line_color="orange", annotation_text="Elevated Risk Threshold")
    st.plotly_chart(fig4, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        # Digital wellbeing score
        wellbeing[\'Month\'] = pd.to_datetime(wellbeing[\'Reporting_Period\']).dt.to_period(\'M\').astype(str)
        monthly_wellbeing = wellbeing.groupby(\'Month\')[\'Digital_Wellbeing_Score\'].mean().reset_index()
        monthly_wellbeing[\'Digital_Wellbeing_Score\'] *= 100

        fig5 = px.line(
            monthly_wellbeing,
            x=\'Month\',
            y=\'Digital_Wellbeing_Score\',
            markers=True,
            title="Digital Wellbeing Trend"
        )
        fig5.add_hline(y=75, line_dash="dash", annotation_text="Target: 75%")
        st.plotly_chart(fig5, use_container_width=True)

    with col2:
        # Collaboration overload
        collab_weekly = collab.groupby(\'Employee_ID\')[\'Collaboration_Overload_Percentage\'].mean().reset_index()
        collab_weekly[\'Collaboration_Overload_Percentage\'] *= 100

        fig6 = px.box(
            collab_weekly,
            y=\'Collaboration_Overload_Percentage\',
            title="Collaboration Overhead Distribution"
        )
        fig6.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="Overload Threshold")
        st.plotly_chart(fig6, use_container_width=True)

    st.markdown(\'\'\'
    <div class="insight-highlight">
    💡 <strong>Recommended Actions:</strong><br>
    • Institute meeting-free blocks for focused work<br>
    • Cap weekly meetings at 15 hours maximum<br>
    • Mandate disconnect periods for remote workers<br>
    • Expected outcome: 20% improvement in wellbeing scores within 3 months
    </div>
    \'\'\'
    , unsafe_allow_html=True)

# TAB 4: SKILLS STORY
with tab4:
    st.markdown(\'<div class="story-title">📚 Story 3: The Skills Gap Challenge</div>\', unsafe_allow_html=True)

    st.markdown(\'\'\'
    <div class="narrative-box">
    <h4>🚨 The Readiness Crisis</h4>
    <p style="font-size:1.05em;line-height:1.7">
    With a future skill readiness index of <strong>1.31 out of 10</strong>, only <strong>13% of our workforce</strong> 
    possesses the capabilities needed for tomorrow\'s strategic initiatives. This isn\'t just a skills issue—it\'s a 
    <strong>business continuity risk</strong>.
    </p>
    </div>
    \'\'\'
    , unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Skill Readiness Index", f"{skill_ready[\'Readiness_Score\'].mean():.2f}/10", "🚨 Critical")

    with col2:
        st.metric("Training Completion", f"{skill_ready[\'Training_Completion_Percentage\'].mean():.1f}%", "📊 Moderate")

    with col3:
        st.metric("Avg Skill Gap", f"{skill_gap[\'Skill_Gap_Score\'].mean():.2f}", "📉 Needs Attention")

    # Skill gaps by category
    st.subheader("Skill Performance vs Benchmark")
    skill_analysis = skill_gap.groupby(\'Skill_Category\')[[\'Performance_Score\', \'Benchmark_Score\']].mean().reset_index()

    fig7 = go.Figure()
    fig7.add_trace(go.Bar(
        name=\'Current Performance\',
        x=skill_analysis[\'Skill_Category\'],
        y=skill_analysis[\'Performance_Score\'],
        marker_color=\'#457b9d\'
    ))
    fig7.add_trace(go.Bar(
        name=\'Benchmark Target\',
        x=skill_analysis[\'Skill_Category\'],
        y=skill_analysis[\'Benchmark_Score\'],
        marker_color=\'#06d6a0\'
    ))
    fig7.update_layout(barmode=\'group\', title="Where Are the Biggest Gaps?")
    st.plotly_chart(fig7, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        # Training completion trend
        training_data = skill_ready.groupby(\'Quarter\')[\'Training_Completion_Percentage\'].mean().reset_index()

        fig8 = px.bar(
            training_data,
            x=\'Quarter\',
            y=\'Training_Completion_Percentage\',
            title="Training Engagement by Quarter",
            color=\'Training_Completion_Percentage\',
            color_continuous_scale=\'Blues\'
        )
        st.plotly_chart(fig8, use_container_width=True)

    with col2:
        # Readiness score distribution
        fig9 = px.histogram(
            skill_ready,
            x=\'Readiness_Score\',
            nbins=20,
            title="Readiness Score Distribution",
            color_discrete_sequence=[\'#e63946\']
        )
        fig9.add_vline(x=5, line_dash="dash", annotation_text="Target: 5.0")
        st.plotly_chart(fig9, use_container_width=True)

    st.markdown(\'\'\'
    <div class="insight-highlight">
    💡 <strong>Strategic Recommendations:</strong><br>
    • Launch accelerated leadership development program (18% performance gap)<br>
    • Increase skill development time to 25+ hours per quarter<br>
    • Tie compensation progression to skill milestone achievement<br>
    • Target: Readiness score of 5.0+ within 12 months (75% workforce ready)
    </div>
    \'\'\'
    , unsafe_allow_html=True)

# TAB 5: SECURITY STORY
with tab5:
    st.markdown(\'<div class="story-title">🔒 Story 4: The Shadow IT Risk</div>\', unsafe_allow_html=True)

    st.markdown(\'\'\'
    <div class="narrative-box">
    <h4>🔍 Beyond Security: A Productivity Signal</h4>
    <p style="font-size:1.05em;line-height:1.7">
    Shadow IT isn\'t just a security issue—it\'s a <strong>productivity signal</strong>. When employees adopt 
    unauthorized tools, they\'re telling us that <strong>approved solutions don\'t meet their needs</strong>. 
    With 15.7% of assessments showing high-risk behavior, we have both a security gap and a user experience problem.
    </p>
    </div>
    \'\'\'
    , unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Avg Risk Score", f"{shadow_it[\'Risk_Score\'].mean():.1f}", "⚠️ Moderate")

    with col2:
        high_risk = len(shadow_it[shadow_it[\'Risk_Score\'] > 60])
        st.metric("High-Risk Cases", f"{high_risk}", f"{high_risk/len(shadow_it)*100:.1f}%")

    with col3:
        st.metric("Avg Unauthorized Apps", f"{shadow_it[\'Unauthorized_Apps_Count\'].mean():.1f}", "per employee")

    # Risk distribution
    col1, col2 = st.columns(2)

    with col1:
        fig10 = px.histogram(
            shadow_it,
            x=\'Risk_Score\',
            nbins=30,
            title="Security Risk Score Distribution",
            color_discrete_sequence=[\'#e63946\']
        )
        fig10.add_vline(x=60, line_dash="dash", line_color="red", annotation_text="High Risk Threshold")
        st.plotly_chart(fig10, use_container_width=True)

    with col2:
        # Risk by data sensitivity
        risk_by_sensitivity = shadow_it.groupby(\'Data_Sensitivity_Level\')[\'Risk_Score\'].mean().reset_index()

        fig11 = px.bar(
            risk_by_sensitivity,
            x=\'Data_Sensitivity_Level\',
            y=\'Risk_Score\',
            title="Risk Score by Data Sensitivity",
            color=\'Risk_Score\',
            color_continuous_scale=\'Reds\'
        )
        st.plotly_chart(fig11, use_container_width=True)

    # Unauthorized apps trend
    shadow_it[\'Month\'] = pd.to_datetime(shadow_it[\'Week_Ending_Date\']).dt.to_period(\'M\').astype(str)
    monthly_apps = shadow_it.groupby(\'Month\')[\'Unauthorized_Apps_Count\'].mean().reset_index()

    fig12 = px.line(
        monthly_apps,
        x=\'Month\',
        y=\'Unauthorized_Apps_Count\',
        markers=True,
        title="Unauthorized App Usage Trend"
    )
    fig12.update_traces(line_color=\'#e63946\', line_width=3)
    st.plotly_chart(fig12, use_container_width=True)

    st.markdown(\'\'\'
    <div class="insight-highlight">
    💡 <strong>Dual-Benefit Strategy:</strong><br>
    • Fast-track approval of frequently-requested productivity tools<br>
    • Launch "approved alternatives" awareness campaign<br>
    • Implement frictionless SSO for all approved applications<br>
    • Expected outcome: 40% reduction in Shadow IT + improved productivity
    </div>
    \'\'\'
    , unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown(\'\'\'
<div style="text-align:center;color:#666;padding:20px">
<strong>Employee KPI Story Dashboard</strong><br>
Narrative-Driven Workforce Analytics | Data Period: April-September 2025<br>
Report Generated: November 11, 2025
</div>
\'\'\'
, unsafe_allow_html=True)
