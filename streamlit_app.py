import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from datetime import datetime, timedelta
import io

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="COO Operational Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM CSS ====================
st.markdown("""
    <style>
    .main { padding: 0px; }

    /* Objective Card */
    .objective-card {
        background: white;
        border-radius: 12px;
        border: 2px solid #e5e7eb;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        cursor: pointer;
        transition: all 0.3s ease;
    }

    .objective-card:hover {
        border-color: #1e40af;
        box-shadow: 0 8px 16px rgba(30, 64, 175, 0.15);
        transform: translateY(-4px);
    }

    .objective-header {
        font-size: 18px;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    .objective-signal {
        font-size: 11px;
        color: #6b7280;
        font-style: italic;
        margin-bottom: 16px;
    }

    /* Sub-objective Box */
    .subobjective-box {
        background: #f9fafb;
        border-left: 4px solid #1e40af;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 10px;
    }

    .subobjective-box.cost { border-left-color: #ef4444; }
    .subobjective-box.quality { border-left-color: #059669; }
    .subobjective-box.efficiency { border-left-color: #f59e0b; }

    .subobjective-title {
        font-size: 12px;
        font-weight: 600;
        color: #1f2937;
    }

    .subobjective-value {
        font-size: 22px;
        font-weight: 700;
        color: #1e40af;
        margin: 8px 0;
    }

    .subobjective-trend {
        font-size: 11px;
        color: #6b7280;
    }

    /* Trend indicator */
    .trend-up { color: #059669; font-weight: 600; }
    .trend-down { color: #ef4444; font-weight: 600; }

    /* Detail Section */
    .detail-section {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 20px;
        border-radius: 10px;
        margin-top: 20px;
    }

    .detail-title {
        font-size: 16px;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 15px;
    }

    /* KPI Card in detail */
    .kpi-detail-card {
        background: white;
        border: 1px solid #e5e7eb;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
    }

    .kpi-detail-label {
        font-size: 12px;
        color: #6b7280;
        font-weight: 600;
    }

    .kpi-detail-value {
        font-size: 28px;
        font-weight: 700;
        color: #1e40af;
        margin: 8px 0;
    }

    .kpi-detail-trend {
        font-size: 11px;
        color: #6b7280;
    }
    </style>
""", unsafe_allow_html=True)

# ==================== LOAD DATA ====================
@st.cache_data
def load_excel_data():
    try:
        excel_file = 'COO_ROI_Dashboard_KPIs_Complete_12.xlsx'

        data = {
            'Role_vs_Reality': pd.read_excel(excel_file, sheet_name='Role_vs_Reality_Analysis'),
            'Automation_ROI': pd.read_excel(excel_file, sheet_name='Automation_ROI_Potential'),
            'Digital_Index': pd.read_excel(excel_file, sheet_name='Digital_Workplace_Index'),
            'Process_Rework': pd.read_excel(excel_file, sheet_name='Process_Rework_Cost'),
            'FTR_Rate': pd.read_excel(excel_file, sheet_name='First_Time_Right_Rate'),
            'Adherence': pd.read_excel(excel_file, sheet_name='Process_Adherence_Rate'),
            'Resilience': pd.read_excel(excel_file, sheet_name='Operational_Resilience_Score'),
            'Escalation': pd.read_excel(excel_file, sheet_name='Escalation_Exception_Patterns'),
            'Capacity': pd.read_excel(excel_file, sheet_name='Hidden_Capacity_Burnout'),
            'Model_Accuracy': pd.read_excel(excel_file, sheet_name='Capacity_Model_Accuracy'),
            'Work_Models': pd.read_excel(excel_file, sheet_name='Work_Models_Effectiveness'),
            'Collaboration': pd.read_excel(excel_file, sheet_name='Collaboration_Overload'),
        }
        return data
    except FileNotFoundError:
        st.error("❌ File not found: 'COO_ROI_Dashboard_KPIs_Complete_12.xlsx'")
        st.stop()

data = load_excel_data()

# ==================== SESSION STATE ====================
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'main'

# ==================== SIDEBAR FILTERS ====================
st.sidebar.markdown("## 🔧 Filters")

role_months = sorted(data['Role_vs_Reality']['Month'].unique())
selected_months = st.sidebar.multiselect(
    "Select Months",
    role_months,
    default=list(role_months),
)

all_departments = sorted(
    list(set(list(data['Role_vs_Reality'].get('Department', []).unique()) + 
             list(data['Capacity'].get('Department', []).unique())))
)
selected_depts = st.sidebar.multiselect(
    "Select Departments",
    all_departments,
    default=all_departments,
)

dept_filter = selected_depts if len(selected_depts) > 0 else None

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ==================== HELPER FUNCTIONS ====================
def filter_data(df, month_col='Month', dept_col='Department'):
    result = df[df[month_col].isin(selected_months)]
    if dept_filter and dept_col in result.columns:
        result = result[result[dept_col].isin(dept_filter)]
    return result

def get_trend_indicator(current, previous):
    if previous == 0:
        return "→ New"
    change = ((current - previous) / previous) * 100
    if change > 0:
        return f"↑ +{change:.1f}%"
    elif change < 0:
        return f"↓ {change:.1f}%"
    return "→ Flat"

# ==================== HEADER ====================
st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); 
                color: white; padding: 30px 20px; border-radius: 0; 
                margin: -25px -20px 25px -20px; text-align: center;">
        <h1 style="margin: 0; font-size: 36px;">📊 COO Operational Dashboard</h1>
        <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">Executive KPI Management & Analysis</p>
    </div>
""", unsafe_allow_html=True)

# ==================== MAIN PAGE (L1) ====================
if st.session_state.current_page == 'main':
    st.markdown("### 🎯 Key Objectives (Click to Explore)")

    col1, col2, col3 = st.columns(3, gap="medium")

    # -------- OBJECTIVE 1: COST & EFFICIENCY --------
    with col1:
        # Calculate metrics
        rework_data = filter_data(data['Process_Rework'])
        auto_data = filter_data(data['Automation_ROI'])
        digital_data = filter_data(data['Digital_Index'])
        role_data = filter_data(data['Role_vs_Reality'])

        rework_pct = rework_data['Rework_Cost_Percentage'].mean() if len(rework_data) > 0 else 0
        auto_roi = auto_data['ROI_Percentage_6M'].mean() if len(auto_data) > 0 else 0
        automation_coverage = (auto_data.shape[0] / max(data['Automation_ROI'].shape[0], 1)) * 100
        friction = digital_data['Friction_Index_Score'].mean() if len(digital_data) > 0 else 0

        if st.button("💰 Cost & Efficiency", key="btn_cost", use_container_width=True, 
                    help="ROI, Rework, Automation, Digital Readiness"):
            st.session_state.current_page = 'cost_efficiency'
            st.rerun()

        st.markdown(f"""
        <div class="objective-card">
            <div class="objective-signal">Monitor: ROI + Rework + Automation Coverage</div>

            <div class="subobjective-box cost">
                <div class="subobjective-title">📊 Rework Cost %</div>
                <div class="subobjective-value">{rework_pct:.1f}%</div>
                <div class="subobjective-trend">Monthly cost impact</div>
            </div>

            <div class="subobjective-box efficiency">
                <div class="subobjective-title">🤖 Automation ROI</div>
                <div class="subobjective-value">{auto_roi:.0f}%</div>
                <div class="subobjective-trend">6-month potential</div>
            </div>

            <div class="subobjective-box efficiency">
                <div class="subobjective-title">🔧 Automation Coverage</div>
                <div class="subobjective-value">{automation_coverage:.0f}%</div>
                <div class="subobjective-trend">Process automation</div>
            </div>

            <div class="subobjective-box efficiency">
                <div class="subobjective-title">📱 Digital Friction</div>
                <div class="subobjective-value">{friction:.1f}</div>
                <div class="subobjective-trend">Lower is better</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # -------- OBJECTIVE 2: EXECUTION & RESILIENCE --------
    with col2:
        ftr_data = filter_data(data['FTR_Rate'])
        adherence_data = filter_data(data['Adherence'])
        resilience_data = filter_data(data['Resilience'])
        escalation_data = filter_data(data['Escalation'])

        ftr_rate = ftr_data['FTR_Rate_Percentage'].mean() if len(ftr_data) > 0 else 0
        adherence = adherence_data['Adherence_Rate_Percentage'].mean() if len(adherence_data) > 0 else 0
        resilience = resilience_data['Resilience_Score'].mean() if len(resilience_data) > 0 else 0
        escalations = escalation_data['Step_Exception_Count'].sum() if len(escalation_data) > 0 else 0

        if st.button("⚙️ Execution & Resilience", key="btn_execution", use_container_width=True,
                    help="FTR, Adherence, Resilience, Exceptions"):
            st.session_state.current_page = 'execution_resilience'
            st.rerun()

        st.markdown(f"""
        <div class="objective-card">
            <div class="objective-signal">Monitor: Quality + Reliability + Risk</div>

            <div class="subobjective-box quality">
                <div class="subobjective-title">✅ FTR Rate</div>
                <div class="subobjective-value">{ftr_rate:.1f}%</div>
                <div class="subobjective-trend">First-time accuracy</div>
            </div>

            <div class="subobjective-box quality">
                <div class="subobjective-title">📋 Process Adherence</div>
                <div class="subobjective-value">{adherence:.1f}%</div>
                <div class="subobjective-trend">Policy compliance</div>
            </div>

            <div class="subobjective-box quality">
                <div class="subobjective-title">🛡️ Resilience Score</div>
                <div class="subobjective-value">{resilience:.1f}/10</div>
                <div class="subobjective-trend">Process robustness</div>
            </div>

            <div class="subobjective-box cost">
                <div class="subobjective-title">🚨 Escalations</div>
                <div class="subobjective-value">{escalations:.0f}</div>
                <div class="subobjective-trend">Exception volume</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # -------- OBJECTIVE 3: WORKFORCE & PRODUCTIVITY --------
    with col3:
        capacity_data = filter_data(data['Capacity'])
        work_data = filter_data(data['Work_Models'])
        model_data = filter_data(data['Model_Accuracy'])

        avg_capacity = capacity_data['Capacity_Utilization_Percentage'].mean() if len(capacity_data) > 0 else 0
        avg_output = work_data['Output_Per_Hour'].mean() if len(work_data) > 0 else 0
        burnout_count = capacity_data[capacity_data['Burnout_Risk_Flag'] == 'Yes'].shape[0]
        model_accuracy = model_data['Forecast_Accuracy_Percentage'].mean() if len(model_data) > 0 else 0

        if st.button("👥 Workforce & Productivity", key="btn_workforce", use_container_width=True,
                    help="Output, Capacity, Health, Model Accuracy"):
            st.session_state.current_page = 'workforce_productivity'
            st.rerun()

        st.markdown(f"""
        <div class="objective-card">
            <div class="objective-signal">Monitor: Output + Capacity + Health</div>

            <div class="subobjective-box efficiency">
                <div class="subobjective-title">⚡ Output/FTE</div>
                <div class="subobjective-value">{avg_output:.2f}</div>
                <div class="subobjective-trend">Productivity per agent</div>
            </div>

            <div class="subobjective-box efficiency">
                <div class="subobjective-title">📊 Capacity Utilization</div>
                <div class="subobjective-value">{avg_capacity:.0f}%</div>
                <div class="subobjective-trend">Workload balance</div>
            </div>

            <div class="subobjective-box cost">
                <div class="subobjective-title">🔥 At-Risk Employees</div>
                <div class="subobjective-value">{burnout_count}</div>
                <div class="subobjective-trend">Burnout risk count</div>
            </div>

            <div class="subobjective-box quality">
                <div class="subobjective-title">🎯 Model Accuracy</div>
                <div class="subobjective-value">{model_accuracy:.0f}%</div>
                <div class="subobjective-trend">Forecast precision</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ==================== DETAIL PAGE 1: COST & EFFICIENCY ====================
elif st.session_state.current_page == 'cost_efficiency':
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("← Back", key="back_cost"):
            st.session_state.current_page = 'main'
            st.rerun()
    with col2:
        st.markdown("### 💰 Cost & Efficiency - Deep Dive")

    st.markdown("---")

    # Calculate all metrics for this objective
    rework_data = filter_data(data['Process_Rework'])
    auto_data = filter_data(data['Automation_ROI'])
    digital_data = filter_data(data['Digital_Index'])
    role_data = filter_data(data['Role_vs_Reality'])
    work_data = filter_data(data['Work_Models'])

    # L2: Detail metrics in grid
    st.markdown("#### Level 2: Key Metrics")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        rework_cost = rework_data['Rework_Cost_Dollars'].sum()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Rework Cost</div>
            <div class="kpi-detail-value">${rework_cost:,.0f}</div>
            <div class="kpi-detail-trend">Monthly impact</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        rework_pct = rework_data['Rework_Cost_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Rework %</div>
            <div class="kpi-detail-value">{rework_pct:.1f}%</div>
            <div class="kpi-detail-trend">of operational cost</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        auto_roi = auto_data['ROI_Percentage_6M'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Automation ROI</div>
            <div class="kpi-detail-value">{auto_roi:.0f}%</div>
            <div class="kpi-detail-trend">6-month return</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        auto_savings = auto_data['Monthly_Cost_Savings'].sum()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Cost Savings</div>
            <div class="kpi-detail-value">${auto_savings:,.0f}</div>
            <div class="kpi-detail-trend">from automation</div>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        friction = digital_data['Friction_Index_Score'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Digital Friction</div>
            <div class="kpi-detail-value">{friction:.1f}</div>
            <div class="kpi-detail-trend">Workplace efficiency</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        low_value = role_data['Low_Value_Work_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Low-Value Work</div>
            <div class="kpi-detail-value">{low_value:.1f}%</div>
            <div class="kpi-detail-trend">optimization gap</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        work_model_eff = work_data['Output_Per_Hour'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Work Model Output</div>
            <div class="kpi-detail-value">{work_model_eff:.2f}</div>
            <div class="kpi-detail-trend">output per hour</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        automation_pct = (auto_data.shape[0] / max(data['Automation_ROI'].shape[0], 1)) * 100
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Automation Coverage</div>
            <div class="kpi-detail-value">{automation_pct:.0f}%</div>
            <div class="kpi-detail-trend">of processes</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # Charts/Analysis
    st.markdown("#### Trends & Analysis")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Rework Cost by Process**")
        if len(rework_data) > 0:
            process_rework = rework_data.groupby('Process_Name').agg({
                'Rework_Cost_Dollars': 'sum'
            }).sort_values('Rework_Cost_Dollars', ascending=False).head(8)

            fig = go.Figure(data=[
                go.Bar(y=process_rework.index, x=process_rework['Rework_Cost_Dollars'], 
                       orientation='h', marker_color='#ef4444')
            ])
            fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("**Top Automation Opportunities**")
        if len(auto_data) > 0:
            top_auto = auto_data.nlargest(8, 'ROI_Percentage_6M')[['Process_Name', 'ROI_Percentage_6M']]

            fig = go.Figure(data=[
                go.Bar(y=top_auto['Process_Name'], x=top_auto['ROI_Percentage_6M'],
                       orientation='h', marker_color='#059669')
            ])
            fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # L3: Detailed Data
    st.markdown("#### Level 3: Detailed Data")

    tabs = st.tabs(["Rework Cost", "Automation ROI", "Digital Index", "Work Models"])

    with tabs[0]:
        st.dataframe(rework_data.head(100), use_container_width=True, hide_index=True)
        csv = rework_data.to_csv(index=False)
        st.download_button("📥 Download Rework Data (CSV)", data=csv, file_name="rework_data.csv", mime="text/csv", key="dl_rework")

    with tabs[1]:
        st.dataframe(auto_data.head(100), use_container_width=True, hide_index=True)
        csv = auto_data.to_csv(index=False)
        st.download_button("📥 Download Automation Data (CSV)", data=csv, file_name="automation_data.csv", mime="text/csv", key="dl_auto")

    with tabs[2]:
        st.dataframe(digital_data.head(100), use_container_width=True, hide_index=True)
        csv = digital_data.to_csv(index=False)
        st.download_button("📥 Download Digital Index (CSV)", data=csv, file_name="digital_index.csv", mime="text/csv", key="dl_digital")

    with tabs[3]:
        st.dataframe(work_data.head(100), use_container_width=True, hide_index=True)
        csv = work_data.to_csv(index=False)
        st.download_button("📥 Download Work Models (CSV)", data=csv, file_name="work_models.csv", mime="text/csv", key="dl_work")

# ==================== DETAIL PAGE 2: EXECUTION & RESILIENCE ====================
elif st.session_state.current_page == 'execution_resilience':
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("← Back", key="back_exec"):
            st.session_state.current_page = 'main'
            st.rerun()
    with col2:
        st.markdown("### ⚙️ Execution & Resilience - Deep Dive")

    st.markdown("---")

    ftr_data = filter_data(data['FTR_Rate'])
    adherence_data = filter_data(data['Adherence'])
    resilience_data = filter_data(data['Resilience'])
    escalation_data = filter_data(data['Escalation'])

    st.markdown("#### Level 2: Key Metrics")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        ftr_rate = ftr_data['FTR_Rate_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">FTR Rate</div>
            <div class="kpi-detail-value">{ftr_rate:.1f}%</div>
            <div class="kpi-detail-trend">first-time accuracy</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        adherence = adherence_data['Adherence_Rate_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Process Adherence</div>
            <div class="kpi-detail-value">{adherence:.1f}%</div>
            <div class="kpi-detail-trend">compliance rate</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        resilience = resilience_data['Resilience_Score'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Resilience Score</div>
            <div class="kpi-detail-value">{resilience:.1f}/10</div>
            <div class="kpi-detail-trend">process robustness</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        escalations = escalation_data['Step_Exception_Count'].sum()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Escalations</div>
            <div class="kpi-detail-value">{escalations:.0f}</div>
            <div class="kpi-detail-trend">exception count</div>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        error_rate = ftr_data['Error_Rate_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Error Rate</div>
            <div class="kpi-detail-value">{error_rate:.1f}%</div>
            <div class="kpi-detail-trend">quality metric</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        risk_pct = resilience_data['Risk_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Risk Exposure</div>
            <div class="kpi-detail-value">{risk_pct:.1f}%</div>
            <div class="kpi-detail-trend">operational risk</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        exception_rate = escalation_data['Exception_Rate_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Exception Rate %</div>
            <div class="kpi-detail-value">{exception_rate:.1f}%</div>
            <div class="kpi-detail-trend">process deviation</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        critical_tasks = resilience_data[resilience_data['Risk_Level'] == 'Critical'].shape[0]
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Critical Tasks</div>
            <div class="kpi-detail-value">{critical_tasks}</div>
            <div class="kpi-detail-trend">at-risk items</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Trends & Analysis")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**FTR Rate by Process**")
        if len(ftr_data) > 0:
            process_ftr = ftr_data.groupby('Process').agg({
                'FTR_Rate_Percentage': 'mean'
            }).sort_values('FTR_Rate_Percentage')

            fig = go.Figure(data=[
                go.Bar(y=process_ftr.index, x=process_ftr['FTR_Rate_Percentage'],
                       orientation='h', marker_color='#059669')
            ])
            fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("**Resilience by Critical Task**")
        if len(resilience_data) > 0:
            task_resilience = resilience_data.groupby('Critical_Task').agg({
                'Resilience_Score': 'mean'
            }).sort_values('Resilience_Score').head(8)

            fig = go.Figure(data=[
                go.Bar(y=task_resilience.index, x=task_resilience['Resilience_Score'],
                       orientation='h', marker_color='#1e40af')
            ])
            fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Level 3: Detailed Data")

    tabs = st.tabs(["FTR Rate", "Process Adherence", "Resilience", "Escalations"])

    with tabs[0]:
        st.dataframe(ftr_data.head(100), use_container_width=True, hide_index=True)
        csv = ftr_data.to_csv(index=False)
        st.download_button("📥 Download FTR Data (CSV)", data=csv, file_name="ftr_data.csv", mime="text/csv", key="dl_ftr")

    with tabs[1]:
        st.dataframe(adherence_data.head(100), use_container_width=True, hide_index=True)
        csv = adherence_data.to_csv(index=False)
        st.download_button("📥 Download Adherence Data (CSV)", data=csv, file_name="adherence_data.csv", mime="text/csv", key="dl_adherence")

    with tabs[2]:
        st.dataframe(resilience_data.head(100), use_container_width=True, hide_index=True)
        csv = resilience_data.to_csv(index=False)
        st.download_button("📥 Download Resilience Data (CSV)", data=csv, file_name="resilience_data.csv", mime="text/csv", key="dl_resilience")

    with tabs[3]:
        st.dataframe(escalation_data.head(100), use_container_width=True, hide_index=True)
        csv = escalation_data.to_csv(index=False)
        st.download_button("📥 Download Escalation Data (CSV)", data=csv, file_name="escalation_data.csv", mime="text/csv", key="dl_escalation")

# ==================== DETAIL PAGE 3: WORKFORCE & PRODUCTIVITY ====================
elif st.session_state.current_page == 'workforce_productivity':
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("← Back", key="back_workforce"):
            st.session_state.current_page = 'main'
            st.rerun()
    with col2:
        st.markdown("### 👥 Workforce & Productivity - Deep Dive")

    st.markdown("---")

    capacity_data = filter_data(data['Capacity'])
    work_data = filter_data(data['Work_Models'])
    model_data = filter_data(data['Model_Accuracy'])
    collab_data = filter_data(data['Collaboration'])

    st.markdown("#### Level 2: Key Metrics")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        avg_output = work_data['Output_Per_Hour'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Output/FTE</div>
            <div class="kpi-detail-value">{avg_output:.2f}</div>
            <div class="kpi-detail-trend">productivity metric</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        avg_capacity = capacity_data['Capacity_Utilization_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Capacity Utilization</div>
            <div class="kpi-detail-value">{avg_capacity:.0f}%</div>
            <div class="kpi-detail-trend">workload balance</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        model_accuracy = model_data['Forecast_Accuracy_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Model Accuracy</div>
            <div class="kpi-detail-value">{model_accuracy:.0f}%</div>
            <div class="kpi-detail-trend">forecast precision</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        burnout_count = capacity_data[capacity_data['Burnout_Risk_Flag'] == 'Yes'].shape[0]
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">At-Risk Employees</div>
            <div class="kpi-detail-value">{burnout_count}</div>
            <div class="kpi-detail-trend">burnout risk</div>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        collab_hours = collab_data['Collaboration_Tools_Time_Hours'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Collab Hours</div>
            <div class="kpi-detail-value">{collab_hours:.1f} hrs</div>
            <div class="kpi-detail-trend">daily collaboration</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        over_capacity = capacity_data[capacity_data['Capacity_Utilization_Percentage'] > 100].shape[0]
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Over-Capacity</div>
            <div class="kpi-detail-value">{over_capacity}</div>
            <div class="kpi-detail-trend">employees</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        avg_cost_per_tx = work_data['Cost_Per_Transaction'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Cost per Transaction</div>
            <div class="kpi-detail-value">${avg_cost_per_tx:.2f}</div>
            <div class="kpi-detail-trend">unit economics</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        throughput = work_data[work_data['Work_Model'] == 'Remote']['Output_Per_Hour'].mean()
        st.markdown(f"""
        <div class="kpi-detail-card">
            <div class="kpi-detail-label">Remote Productivity</div>
            <div class="kpi-detail-value">{throughput:.2f}</div>
            <div class="kpi-detail-trend">output/hour</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Trends & Analysis")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Capacity Utilization by Department**")
        if len(capacity_data) > 0:
            dept_capacity = capacity_data.groupby('Department').agg({
                'Capacity_Utilization_Percentage': 'mean'
            }).sort_values('Capacity_Utilization_Percentage', ascending=False)

            fig = go.Figure(data=[
                go.Bar(y=dept_capacity.index, x=dept_capacity['Capacity_Utilization_Percentage'],
                       orientation='h', marker_color='#f59e0b')
            ])
            fig.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="Optimal")
            fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("**Output by Work Model**")
        if len(work_data) > 0:
            model_output = work_data.groupby('Work_Model').agg({
                'Output_Per_Hour': 'mean'
            })

            fig = go.Figure(data=[
                go.Bar(x=model_output.index, y=model_output['Output_Per_Hour'],
                       marker_color='#059669')
            ])
            fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Level 3: Detailed Data")

    tabs = st.tabs(["Capacity", "Work Models", "Model Accuracy", "Collaboration"])

    with tabs[0]:
        st.dataframe(capacity_data.head(100), use_container_width=True, hide_index=True)
        csv = capacity_data.to_csv(index=False)
        st.download_button("📥 Download Capacity Data (CSV)", data=csv, file_name="capacity_data.csv", mime="text/csv", key="dl_capacity")

    with tabs[1]:
        st.dataframe(work_data.head(100), use_container_width=True, hide_index=True)
        csv = work_data.to_csv(index=False)
        st.download_button("📥 Download Work Models (CSV)", data=csv, file_name="work_models_data.csv", mime="text/csv", key="dl_workmodels")

    with tabs[2]:
        st.dataframe(model_data.head(100), use_container_width=True, hide_index=True)
        csv = model_data.to_csv(index=False)
        st.download_button("📥 Download Model Accuracy (CSV)", data=csv, file_name="model_accuracy.csv", mime="text/csv", key="dl_model")

    with tabs[3]:
        st.dataframe(collab_data.head(100), use_container_width=True, hide_index=True)
        csv = collab_data.to_csv(index=False)
        st.download_button("📥 Download Collaboration (CSV)", data=csv, file_name="collaboration_data.csv", mime="text/csv", key="dl_collab")

# ==================== FOOTER ====================
st.divider()
st.markdown(f"""
    <div style="text-align: center; padding: 15px; color: #6b7280; font-size: 11px;">
        <strong>COO Dashboard v7.0 - Hierarchical Navigation</strong> | 
        {len(selected_months)} months | {len(dept_filter) if dept_filter else len(all_departments)} departments | 
        Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
    </div>
""", unsafe_allow_html=True)
