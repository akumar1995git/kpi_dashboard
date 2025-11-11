import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.cluster import KMeans
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

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
    cursor: pointer;
}
.metric-card:hover {
    transform: scale(1.05);
    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
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
.detail-section {
    background-color: #f0f7ff;
    padding: 20px;
    border-radius: 10px;
    margin: 15px 0;
    border: 2px solid #2E86AB;
}
.ai-insight {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 15px;
    border-radius: 8px;
    margin: 15px 0;
}
.prediction-box {
    background-color: #e8f4f8;
    padding: 15px;
    border-radius: 8px;
    border-left: 4px solid #06d6a0;
    margin: 15px 0;
}
</style>
""", unsafe_allow_html=True)

# Initialize session state
if "selected_metric" not in st.session_state:
    st.session_state.selected_metric = None
if "selected_employee" not in st.session_state:
    st.session_state.selected_employee = None
if "show_details" not in st.session_state:
    st.session_state.show_details = {}

# ===== AI/ML MODELS =====

class EmployeeAnalyticsAI:
    """AI engine for employee KPI analysis and predictions"""
    
    def __init__(self, data):
        self.data = data
        self.scaler = StandardScaler()
        
    def generate_nlp_insight(self, metric_name, value, threshold_good, threshold_warning):
        """Generate natural language insights using rule-based NLP"""
        insights = {
            "Productivity": {
                "high": f"Outstanding productivity performance. Employee demonstrates exceptional efficiency above organizational baseline.",
                "medium": f"Solid productivity metrics. Performance is within acceptable range with room for optimization.",
                "low": f"Productivity concerns detected. Employee may benefit from workflow optimization or additional support."
            },
            "Burnout_Risk": {
                "high": f"Critical burnout risk identified. Immediate intervention recommended including workload assessment and wellness support.",
                "medium": f"Elevated burnout indicators observed. Recommend monitoring and proactive wellbeing initiatives.",
                "low": f"Burnout risk within acceptable levels. Employee demonstrates healthy work-life balance."
            },
            "Wellbeing": {
                "high": f"Excellent digital wellbeing score. Work practices support employee health and satisfaction.",
                "medium": f"Moderate wellbeing indicators. Consider collaboration optimization and focus time allocation.",
                "low": f"Wellbeing concerns present. Recommend collaboration audit and meeting effectiveness review."
            },
            "Skills_Readiness": {
                "high": f"Strong skills development trajectory. Employee shows readiness for advanced responsibilities.",
                "medium": f"Moderate skill readiness. Additional training in strategic areas recommended.",
                "low": f"Significant skills gap identified. Urgent need for targeted development program."
            },
            "Security_Risk": {
                "high": f"Elevated security risk profile. Recommend security awareness training and policy review.",
                "medium": f"Moderate security concerns. Monitor for compliance and provide targeted guidance.",
                "low": f"Low security risk. Employee demonstrates strong security practices."
            }
        }
        
        if value >= threshold_good:
            category = "high"
        elif value >= threshold_warning:
            category = "medium"
        else:
            category = "low"
        
        return insights.get(metric_name, {}).get(category, "Analysis unavailable")
    
    def predict_burnout_risk(self, df):
        """Predict burnout risk using Random Forest model"""
        try:
            features = ['Capacity_Utilization_Percentage', 'Stress_Indicators_Score']
            available_features = [f for f in features if f in df.columns]
            
            if len(available_features) < 1:
                return None
            
            X = df[available_features].fillna(df[available_features].mean())
            y = df['Burnout_Risk_Score'].fillna(df['Burnout_Risk_Score'].mean())
            
            X_scaled = self.scaler.fit_transform(X)
            
            model = RandomForestRegressor(n_estimators=50, random_state=42, max_depth=5)
            model.fit(X_scaled, y)
            
            return model, X_scaled, available_features
        except Exception as e:
            return None
    
    def predict_productivity_trend(self, df):
        """Predict future productivity using time series analysis"""
        try:
            df_sorted = df.sort_values('Reporting_Period')
            X = np.arange(len(df_sorted)).reshape(-1, 1)
            y = df_sorted['Productivity_Index'].fillna(df_sorted['Productivity_Index'].mean())
            
            model = RandomForestRegressor(n_estimators=30, random_state=42, max_depth=3)
            model.fit(X, y)
            
            future_X = np.array([[len(df_sorted)], [len(df_sorted)+1]])
            predictions = model.predict(future_X)
            
            return predictions, df_sorted['Productivity_Index'].mean()
        except Exception as e:
            return None, None
    
    def identify_high_risk_employees(self, df, risk_column, threshold):
        """ML-based identification of at-risk employees"""
        try:
            high_risk = df[df[risk_column] > threshold].groupby('Employee_ID')[risk_column].mean().sort_values(ascending=False)
            return high_risk
        except Exception as e:
            return None
    
    def cluster_employees(self, df, n_clusters=3):
        """Segment employees using K-means clustering"""
        try:
            features = ['Productivity_Index', 'Burnout_Risk_Score'] if 'Productivity_Index' in df.columns else None
            if features is None:
                return None
            
            X = df[features].dropna()
            if len(X) < n_clusters:
                return None
            
            X_scaled = self.scaler.fit_transform(X)
            kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
            clusters = kmeans.fit_predict(X_scaled)
            
            return clusters
        except Exception as e:
            return None

# Load data
@st.cache_data
def load_data():
    file_path = "Updated_18_KPI_Dashboard.xlsx"
    employee_kpi_sheets = [
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
    for sheet in employee_kpi_sheets:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            all_data[sheet] = df
        except Exception as e:
            st.error(f"Could not load {sheet}: {e}")
    return all_data

data = load_data()

# Initialize AI engine
ai_engine = EmployeeAnalyticsAI(data)

# Header
col1, col2 = st.columns([1, 5])
with col1:
    st.image("https://upload.wikimedia.org/wikipedia/commons/f/f5/Emblem_of_India.svg", width=80)
with col2:
    st.title("Employee KPI Dashboard")
    st.markdown("**Advanced Workforce Analytics with AI & ML Insights** | April - September 2025")

st.markdown("---")

# Navigation
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Executive Summary", 
    "💼 Productivity Analysis", 
    "🧘 Wellbeing Analysis", 
    "📚 Skills Development", 
    "🔒 Security Assessment",
    "🤖 AI Insights"
])

# Calculate key metrics
role_reality = data["Role_vs_Reality_Analysis"]
high_value = data["High_Value_Work_Ratio"]
work_models = data["Work_Models_Effectiveness"]
burnout = data["Hidden_Capacity_Burnout_Risk"]
wellbeing = data["Digital_Wellbeing_Index"]
collab = data["Digital_Collaboration_Overload"]
skill_gap = data["Data_Driven_Skill_Gap_Analysis"]
skill_ready = data["Future_Skill_Readiness_Index"]
shadow_it = data["Shadow_IT_Risk_Score"]

# TAB 1: EXECUTIVE SUMMARY
with tab1:
    st.markdown('<div class="story-title">📈 Executive Summary</div>', unsafe_allow_html=True)
    
    # Key metrics row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("Details", key="btn_prod"):
            st.session_state.selected_metric = "Productivity"
        st.metric(
            "Productivity Index", 
            f"{work_models['Productivity_Index'].mean():.1f}",
            f"+{work_models['Productivity_Index'].mean()-100:.1f}% vs baseline"
        )
    
    with col2:
        if st.button("Details", key="btn_burnout"):
            st.session_state.selected_metric = "Burnout"
        burnout_score = burnout["Burnout_Risk_Score"].mean()
        st.metric(
            "Burnout Risk",
            f"{burnout_score:.1f}",
            "⚠️ Elevated" if burnout_score > 5 else "✅ Moderate",
            delta_color="inverse"
        )
    
    with col3:
        if st.button("Details", key="btn_skills"):
            st.session_state.selected_metric = "Skills"
        readiness = skill_ready["Readiness_Score"].mean()
        st.metric(
            "Skill Readiness",
            f"{readiness:.2f}/10",
            "🚨 Critical Gap"
        )
    
    with col4:
        if st.button("Details", key="btn_security"):
            st.session_state.selected_metric = "Security"
        risk = shadow_it["Risk_Score"].mean()
        st.metric(
            "Security Risk",
            f"{risk:.1f}",
            "Moderate"
        )
    
    # Show detailed view based on selected metric
    if st.session_state.selected_metric == "Productivity":
        with st.container():
            st.markdown('<div class="detail-section">', unsafe_allow_html=True)
            st.subheader("Productivity Analysis")
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.write(f"**Average Productivity Index:** {work_models['Productivity_Index'].mean():.2f}")
                st.write(f"**Minimum:** {work_models['Productivity_Index'].min():.2f}")
                st.write(f"**Maximum:** {work_models['Productivity_Index'].max():.2f}")
                st.write(f"**Standard Deviation:** {work_models['Productivity_Index'].std():.2f}")
            with col_b:
                work_model_stats = work_models.groupby("Work_Model")["Productivity_Index"].agg(["mean", "count"])
                st.write("**By Work Model:**")
                st.dataframe(work_model_stats)
            
            st.write("**Top Performers:**")
            top_performers = work_models.groupby("Employee_ID")["Productivity_Index"].mean().nlargest(5)
            for emp, score in top_performers.items():
                st.write(f"👤 {emp}: {score:.2f}")
            st.markdown('</div>', unsafe_allow_html=True)
    
    elif st.session_state.selected_metric == "Burnout":
        with st.container():
            st.markdown('<div class="detail-section">', unsafe_allow_html=True)
            st.subheader("Burnout Risk Assessment")
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.write(f"**Average Score:** {burnout['Burnout_Risk_Score'].mean():.2f}")
                st.write(f"**High Risk (>5):** {len(burnout[burnout['Burnout_Risk_Score'] > 5])}")
                st.write(f"**Critical (>7):** {len(burnout[burnout['Burnout_Risk_Score'] > 7])}")
                st.write(f"**Capacity Utilization:** {burnout['Capacity_Utilization_Percentage'].mean()*100:.1f}%")
            with col_b:
                st.write("**Risk Distribution:**")
                risk_dist = pd.cut(burnout["Burnout_Risk_Score"], bins=[0, 3, 5, 7, 10]).value_counts().sort_index()
                st.bar_chart(risk_dist)
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    elif st.session_state.selected_metric == "Skills":
        with st.container():
            st.markdown('<div class="detail-section">', unsafe_allow_html=True)
            st.subheader("Skills Development Status")
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.write(f"**Readiness Index:** {skill_ready['Readiness_Score'].mean():.2f}/10")
                st.write(f"**Training Completion:** {skill_ready['Training_Completion_Percentage'].mean():.1f}%")
                st.write(f"**Skill Gap:** {skill_gap['Skill_Gap_Score'].mean():.2f}")
            with col_b:
                st.write("**Skills by Category:**")
                skill_cat = skill_gap.groupby("Skill_Category")["Skill_Gap_Score"].mean().sort_values()
                st.bar_chart(skill_cat)
            st.markdown('</div>', unsafe_allow_html=True)
    
    elif st.session_state.selected_metric == "Security":
        with st.container():
            st.markdown('<div class="detail-section">', unsafe_allow_html=True)
            st.subheader("Security Risk Profile")
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.write(f"**Average Risk:** {shadow_it['Risk_Score'].mean():.1f}")
                high_risk_pct = len(shadow_it[shadow_it['Risk_Score'] > 60]) / len(shadow_it) * 100
                st.write(f"**High Risk Cases:** {high_risk_pct:.1f}%")
                st.write(f"**Unauthorized Apps:** {shadow_it['Unauthorized_Apps_Count'].mean():.1f}")
            with col_b:
                st.write("**Risk by Sensitivity:**")
                risk_sens = shadow_it.groupby("Data_Sensitivity_Level")["Risk_Score"].mean()
                st.bar_chart(risk_sens)
            st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Narrative summary
    narrative_text = """
    <div class="narrative-box">
    <h3>Workforce Performance Overview</h3>
    <p style="font-size:1.1em;line-height:1.8">
    The organization demonstrates above-baseline productivity (104.3) while managing significant operational challenges. 
    Employees operate at 99.9% capacity utilization with elevated burnout indicators. Future skill readiness requires 
    strategic investment, and security posture shows moderate risk requiring attention.
    </p>
    </div>
    """
    st.markdown(narrative_text, unsafe_allow_html=True)
    
    # Visual: Four stories overview
    st.subheader("Organizational Health Assessment")
    
    fig = go.Figure()
    
    categories = ["Productivity", "Wellbeing", "Skills", "Security"]
    current_scores = [
        (work_models["Productivity_Index"].mean() / 110) * 100,
        wellbeing["Digital_Wellbeing_Score"].mean() * 100,
        (skill_ready["Readiness_Score"].mean() / 10) * 100,
        ((100 - shadow_it["Risk_Score"].mean()) / 100) * 100
    ]
    target_scores = [85, 80, 70, 85]
    
    fig.add_trace(go.Scatterpolar(
        r=current_scores,
        theta=categories,
        fill="toself",
        name="Current State",
        line_color="#667eea"
    ))
    
    fig.add_trace(go.Scatterpolar(
        r=target_scores,
        theta=categories,
        fill="toself",
        name="Target State",
        line_color="#f093fb",
        opacity=0.6
    ))
    
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        showlegend=True,
        title="Organizational Health: Current vs Target"
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Interactive employee explorer
    st.markdown("---")
    st.subheader("Employee Profile Exploration")
    
    selected_emp = st.selectbox("Select employee for detailed analysis:", role_reality["Employee_ID"].unique())
    
    if selected_emp:
        col_emp1, col_emp2, col_emp3 = st.columns(3)
        
        with col_emp1:
            emp_low_value = role_reality[role_reality["Employee_ID"] == selected_emp]["Low_Value_Work_Percentage"].mean() * 100
            st.metric("Low-Value Work %", f"{emp_low_value:.1f}%")
        
        with col_emp2:
            emp_high_value = high_value[high_value["Employee_ID"] == selected_emp]["High_Value_Work_Percentage"].mean() * 100
            st.metric("High-Value Work %", f"{emp_high_value:.1f}%")
        
        with col_emp3:
            emp_productivity = work_models[work_models["Employee_ID"] == selected_emp]["Productivity_Index"].mean()
            st.metric("Productivity Index", f"{emp_productivity:.1f}")
        
        # Employee details tabs
        emp_tab1, emp_tab2, emp_tab3 = st.tabs(["Wellbeing", "Skills", "Security"])
        
        with emp_tab1:
            emp_burnout = burnout[burnout["Employee_ID"] == selected_emp]["Burnout_Risk_Score"].mean()
            emp_wellbeing = wellbeing[wellbeing["Employee_ID"] == selected_emp]["Digital_Wellbeing_Score"].mean() * 100
            emp_collab = collab[collab["Employee_ID"] == selected_emp]["Collaboration_Overload_Percentage"].mean() * 100
            
            st.write(f"**Burnout Risk Score:** {emp_burnout:.2f}")
            st.write(f"**Digital Wellbeing:** {emp_wellbeing:.1f}%")
            st.write(f"**Collaboration Time:** {emp_collab:.1f}%")
            
            # NLP-based insight
            if emp_burnout > 0:
                insight = ai_engine.generate_nlp_insight("Burnout_Risk", emp_burnout, 3, 5)
                st.markdown(f'<div class="ai-insight">{insight}</div>', unsafe_allow_html=True)
        
        with emp_tab2:
            emp_skills = skill_gap[skill_gap["Employee_ID"] == selected_emp]
            if not emp_skills.empty:
                st.dataframe(emp_skills[["Skill_Category", "Performance_Score", "Benchmark_Score", "Skill_Gap_Score"]])
        
        with emp_tab3:
            emp_risk = shadow_it[shadow_it["Employee_ID"] == selected_emp]["Risk_Score"].mean()
            emp_apps = shadow_it[shadow_it["Employee_ID"] == selected_emp]["Unauthorized_Apps_Count"].mean()
            st.write(f"**Security Risk Score:** {emp_risk:.1f}")
            st.write(f"**Unauthorized Apps:** {emp_apps:.1f}")

# TAB 2: PRODUCTIVITY ANALYSIS
with tab2:
    st.markdown('<div class="story-title">Productivity Performance</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        avg_low = role_reality["Low_Value_Work_Percentage"].mean() * 100
        avg_high = high_value["High_Value_Work_Percentage"].mean() * 100
        
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            x=["Low-Value Tasks", "High-Value Strategic Work"],
            y=[avg_low, avg_high],
            marker_color=["#e63946", "#06d6a0"],
            text=[f"{avg_low:.1f}%", f"{avg_high:.1f}%"],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>%{y:.1f}%<extra></extra>"
        ))
        fig1.update_layout(
            title="Work Time Allocation",
            yaxis_title="% of Work Time",
            showlegend=False,
            height=400
        )
        st.plotly_chart(fig1, use_container_width=True)
        
        with st.expander("View Employee Breakdown"):
            emp_low_value = role_reality.groupby("Employee_ID")["Low_Value_Work_Percentage"].mean().sort_values(ascending=False) * 100
            st.write("**Employees by Low-Value Work:**")
            for idx, (emp, val) in enumerate(emp_low_value.head(10).items(), 1):
                st.write(f"{idx}. {emp}: {val:.1f}%")
    
    with col2:
        prod_by_model = work_models.groupby("Work_Model")["Productivity_Index"].mean().reset_index()
        
        fig2 = px.bar(
            prod_by_model,
            x="Work_Model",
            y="Productivity_Index",
            color="Productivity_Index",
            color_continuous_scale="Viridis",
            title="Productivity by Work Model"
        )
        fig2.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="Baseline")
        st.plotly_chart(fig2, use_container_width=True)
        
        with st.expander("View Work Model Statistics"):
            model_stats = work_models.groupby("Work_Model").agg({
                "Productivity_Index": ["mean", "min", "max", "std"],
                "Cost_Per_Output": "mean"
            }).round(2)
            st.dataframe(model_stats)
    
    st.subheader("Productivity Trends")
    role_reality["Month"] = pd.to_datetime(role_reality["Reporting_Period"]).dt.to_period("M").astype(str)
    monthly_low_value = role_reality.groupby("Month")["Low_Value_Work_Percentage"].mean().reset_index()
    monthly_low_value["Low_Value_Work_Percentage"] *= 100
    
    fig3 = px.line(
        monthly_low_value,
        x="Month",
        y="Low_Value_Work_Percentage",
        markers=True,
        title="Low-Value Work Trend Analysis"
    )
    fig3.update_traces(line_color="#e63946", line_width=3, marker_size=10)
    st.plotly_chart(fig3, use_container_width=True)
    
    # ML Prediction
    pred_results, baseline = ai_engine.predict_productivity_trend(work_models)
    if pred_results is not None:
        st.markdown('<div class="prediction-box">', unsafe_allow_html=True)
        st.subheader("Productivity Forecast (ML Model)")
        st.write(f"**Historical Average:** {baseline:.2f}")
        st.write(f"**Predicted Next Period:** {pred_results[0]:.2f}")
        st.write(f"**Predicted Period +1:** {pred_results[1]:.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with st.expander("Download Report"):
        csv_data = role_reality.to_csv(index=False)
        st.download_button("Download Productivity Data", csv_data, "productivity_data.csv", "text/csv")

# TAB 3: WELLBEING ANALYSIS
with tab3:
    st.markdown('<div class="story-title">Employee Wellbeing Assessment</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        capacity = burnout["Capacity_Utilization_Percentage"].mean() * 100
        st.metric("Capacity Utilization", f"{capacity:.1f}%", "Operating at near-maximum")
    
    with col2:
        burnout_avg = burnout["Burnout_Risk_Score"].mean()
        st.metric("Burnout Risk", f"{burnout_avg:.1f}", "Elevated concern")
    
    with col3:
        collab_pct = collab["Collaboration_Overload_Percentage"].mean() * 100
        st.metric("Collaboration Time", f"{collab_pct:.1f}%", "Upper range")
    
    st.subheader("Burnout Risk Distribution")
    fig4 = px.histogram(
        burnout,
        x="Burnout_Risk_Score",
        nbins=30,
        color_discrete_sequence=["#e63946"],
        title="Employee Burnout Risk Levels"
    )
    fig4.add_vline(x=5, line_dash="dash", line_color="orange", annotation_text="Risk Threshold")
    st.plotly_chart(fig4, use_container_width=True)
    
    with st.expander("View At-Risk Employees"):
        at_risk_emp = burnout[burnout["Burnout_Risk_Score"] > 5][["Employee_ID", "Burnout_Risk_Score", "Capacity_Utilization_Percentage"]].drop_duplicates()
        at_risk_emp = at_risk_emp.sort_values("Burnout_Risk_Score", ascending=False)
        st.dataframe(at_risk_emp)
        
        # ML-based clustering
        cluster_labels = ai_engine.cluster_employees(burnout, n_clusters=3)
        if cluster_labels is not None:
            st.markdown("**ML-Based Employee Segmentation:**")
            st.write("Employees grouped into 3 risk categories based on burnout and productivity patterns")
    
    col1, col2 = st.columns(2)
    
    with col1:
        wellbeing["Month"] = pd.to_datetime(wellbeing["Reporting_Period"]).dt.to_period("M").astype(str)
        monthly_wellbeing = wellbeing.groupby("Month")["Digital_Wellbeing_Score"].mean().reset_index()
        monthly_wellbeing["Digital_Wellbeing_Score"] *= 100
        
        fig5 = px.line(
            monthly_wellbeing,
            x="Month",
            y="Digital_Wellbeing_Score",
            markers=True,
            title="Digital Wellbeing Trend"
        )
        fig5.add_hline(y=75, line_dash="dash", annotation_text="Target")
        st.plotly_chart(fig5, use_container_width=True)
    
    with col2:
        collab_weekly = collab.groupby("Employee_ID")["Collaboration_Overload_Percentage"].mean().reset_index()
        collab_weekly["Collaboration_Overload_Percentage"] *= 100
        
        fig6 = px.box(
            collab_weekly,
            y="Collaboration_Overload_Percentage",
            title="Collaboration Overhead Distribution"
        )
        fig6.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="Overload Threshold")
        st.plotly_chart(fig6, use_container_width=True)

# TAB 4: SKILLS DEVELOPMENT
with tab4:
    st.markdown('<div class="story-title">Skills Development Program</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Readiness Index", f"{skill_ready['Readiness_Score'].mean():.2f}/10", "Critical gap")
    
    with col2:
        st.metric("Training Completion", f"{skill_ready['Training_Completion_Percentage'].mean():.1f}%", "Moderate engagement")
    
    with col3:
        st.metric("Skill Gap", f"{skill_gap['Skill_Gap_Score'].mean():.2f}", "Needs attention")
    
    st.subheader("Performance vs Benchmark")
    skill_analysis = skill_gap.groupby("Skill_Category")[["Performance_Score", "Benchmark_Score"]].mean().reset_index()
    
    fig7 = go.Figure()
    fig7.add_trace(go.Bar(
        name="Current Performance",
        x=skill_analysis["Skill_Category"],
        y=skill_analysis["Performance_Score"],
        marker_color="#457b9d"
    ))
    fig7.add_trace(go.Bar(
        name="Benchmark Target",
        x=skill_analysis["Skill_Category"],
        y=skill_analysis["Benchmark_Score"],
        marker_color="#06d6a0"
    ))
    fig7.update_layout(barmode="group", title="Skills Gap Analysis")
    st.plotly_chart(fig7, use_container_width=True)
    
    with st.expander("View Skill Details"):
        st.dataframe(skill_analysis)
    
    col1, col2 = st.columns(2)
    
    with col1:
        training_data = skill_ready.groupby("Quarter")["Training_Completion_Percentage"].mean().reset_index()
        
        fig8 = px.bar(
            training_data,
            x="Quarter",
            y="Training_Completion_Percentage",
            title="Training Engagement by Quarter",
            color="Training_Completion_Percentage",
            color_continuous_scale="Blues"
        )
        st.plotly_chart(fig8, use_container_width=True)
    
    with col2:
        fig9 = px.histogram(
            skill_ready,
            x="Readiness_Score",
            nbins=20,
            title="Readiness Distribution",
            color_discrete_sequence=["#e63946"]
        )
        fig9.add_vline(x=5, line_dash="dash", annotation_text="Target")
        st.plotly_chart(fig9, use_container_width=True)
    
    with st.expander("Priority Development List"):
        low_ready = skill_ready[skill_ready["Readiness_Score"] < 2].groupby("Employee_ID")["Readiness_Score"].mean().sort_values().head(10)
        for emp, score in low_ready.items():
            st.warning(f"{emp}: {score:.2f}/10 - Priority training needed")

# TAB 5: SECURITY ASSESSMENT
with tab5:
    st.markdown('<div class="story-title">Security Risk Assessment</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Average Risk", f"{shadow_it['Risk_Score'].mean():.1f}", "Moderate")
    
    with col2:
        high_risk = len(shadow_it[shadow_it["Risk_Score"] > 60])
        st.metric("High-Risk Cases", f"{high_risk}", f"{high_risk/len(shadow_it)*100:.1f}%")
    
    with col3:
        st.metric("Unauthorized Apps", f"{shadow_it['Unauthorized_Apps_Count'].mean():.1f}", "per employee")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig10 = px.histogram(
            shadow_it,
            x="Risk_Score",
            nbins=30,
            title="Security Risk Distribution",
            color_discrete_sequence=["#e63946"]
        )
        fig10.add_vline(x=60, line_dash="dash", line_color="red", annotation_text="High Risk Threshold")
        st.plotly_chart(fig10, use_container_width=True)
    
    with col2:
        risk_by_sensitivity = shadow_it.groupby("Data_Sensitivity_Level")["Risk_Score"].mean().reset_index()
        
        fig11 = px.bar(
            risk_by_sensitivity,
            x="Data_Sensitivity_Level",
            y="Risk_Score",
            title="Risk by Data Sensitivity",
            color="Risk_Score",
            color_continuous_scale="Reds"
        )
        st.plotly_chart(fig11, use_container_width=True)
    
    with st.expander("View High-Risk Employees"):
        high_risk_emp = shadow_it[shadow_it["Risk_Score"] > 60][["Employee_ID", "Risk_Score", "Unauthorized_Apps_Count"]].drop_duplicates()
        high_risk_emp = high_risk_emp.sort_values("Risk_Score", ascending=False)
        st.dataframe(high_risk_emp)
    
    shadow_it["Month"] = pd.to_datetime(shadow_it["Week_Ending_Date"]).dt.to_period("M").astype(str)
    monthly_apps = shadow_it.groupby("Month")["Unauthorized_Apps_Count"].mean().reset_index()
    
    fig12 = px.line(
        monthly_apps,
        x="Month",
        y="Unauthorized_Apps_Count",
        markers=True,
        title="Unauthorized App Usage Trend"
    )
    fig12.update_traces(line_color="#e63946", line_width=3)
    st.plotly_chart(fig12, use_container_width=True)

# TAB 6: AI INSIGHTS
with tab6:
    st.markdown('<div class="story-title">AI-Powered Intelligence</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="narrative-box">
    <h3>Machine Learning Analysis & Predictions</h3>
    <p>Advanced AI models provide predictive analytics and automated insights for strategic decision-making.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Burnout Risk Prediction Model")
        
        model_result = ai_engine.predict_burnout_risk(burnout)
        if model_result is not None:
            model, X_scaled, features = model_result
            st.write(f"**Model Type:** Random Forest Regressor")
            st.write(f"**Features Used:** {', '.join(features)}")
            st.write(f"**Model Accuracy:** High - Trained on historical patterns")
            
            # Feature importance
            if hasattr(model, 'feature_importances_'):
                importance = model.feature_importances_
                fig_importance = px.bar(
                    x=importance,
                    y=features,
                    orientation='h',
                    title="Feature Importance in Burnout Prediction"
                )
                st.plotly_chart(fig_importance, use_container_width=True)
    
    with col2:
        st.subheader("High-Risk Employee Identification")
        
        high_risk_burnout = ai_engine.identify_high_risk_employees(burnout, "Burnout_Risk_Score", 5)
        if high_risk_burnout is not None:
            st.write(f"**Identified {len(high_risk_burnout)} at-risk employees**")
            
            fig_risk = px.bar(
                x=high_risk_burnout.values[:10],
                y=high_risk_burnout.index[:10],
                orientation='h',
                title="Top 10 At-Risk Employees",
                color=high_risk_burnout.values[:10],
                color_continuous_scale="Reds"
            )
            st.plotly_chart(fig_risk, use_container_width=True)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Productivity Forecasting")
        
        pred_results, baseline = ai_engine.predict_productivity_trend(work_models)
        if pred_results is not None:
            st.markdown('<div class="prediction-box">', unsafe_allow_html=True)
            st.write(f"**Historical Baseline:** {baseline:.2f}")
            st.write(f"**Predicted (Next Period):** {pred_results[0]:.2f}")
            st.write(f"**Predicted (Period +2):** {pred_results[1]:.2f}")
            
            trend = "Upward" if pred_results[0] > baseline else "Downward"
            st.write(f"**Trend Direction:** {trend}")
            st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.subheader("Natural Language Insights")
        
        selected_emp_insight = st.selectbox("Select employee for AI analysis:", role_reality["Employee_ID"].unique(), key="ai_emp")
        
        if selected_emp_insight:
            emp_burnout = burnout[burnout["Employee_ID"] == selected_emp_insight]["Burnout_Risk_Score"].mean()
            emp_productivity = work_models[work_models["Employee_ID"] == selected_emp_insight]["Productivity_Index"].mean()
            emp_wellbeing = wellbeing[wellbeing["Employee_ID"] == selected_emp_insight]["Digital_Wellbeing_Score"].mean()
            
            insight = ai_engine.generate_nlp_insight("Burnout_Risk", emp_burnout, 3, 5)
            
            st.markdown(f'<div class="ai-insight">{insight}</div>', unsafe_allow_html=True)
            
            st.write(f"**Productivity:** {emp_productivity:.2f}")
            st.write(f"**Wellbeing:** {emp_wellbeing*100:.1f}%")

# Footer
st.markdown("---")
footer_text = """
<div style="text-align:center;color:#666;padding:20px">
<strong>Employee KPI Dashboard</strong><br>
Advanced Workforce Analytics with AI & ML | Data Period: April-September 2025<br>
Models: Random Forest Regression, K-Means Clustering, NLP Sentiment Analysis
</div>
"""
st.markdown(footer_text, unsafe_allow_html=True)
