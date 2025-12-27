import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from datetime import datetime, timedelta

# Page configuration
st.set_page_config(
    page_title="COO Operational Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
    <style>
    .main {
        padding: 0px;
    }
    .kpi-card {
        background: white;
        padding: 20px;
        border-radius: 8px;
        border-left: 5px solid #1e40af;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .kpi-card.green {
        border-left-color: #059669;
    }
    .kpi-card.amber {
        border-left-color: #d97706;
    }
    .kpi-card.red {
        border-left-color: #dc2626;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: #1f2937;
    }
    .metric-label {
        font-size: 12px;
        color: #6b7280;
        text-transform: uppercase;
        margin-bottom: 8px;
    }
    .status-pill {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        margin-top: 8px;
    }
    .status-good { background: #d1fae5; color: #065f46; }
    .status-amber { background: #fef3c7; color: #92400e; }
    .status-red { background: #fee2e2; color: #991b1b; }
    .trend {
        font-size: 12px;
        font-weight: 600;
        padding: 4px 8px;
        border-radius: 3px;
        display: inline-block;
        margin-top: 6px;
    }
    .trend-positive { background: #d1fae5; color: #047857; }
    .trend-negative { background: #fee2e2; color: #991b1b; }
    </style>
""", unsafe_allow_html=True)

# ==================== LOAD DATA FROM EXCEL ====================
@st.cache_data
def load_excel_data():
    """Load all data from Excel file"""
    try:
        excel_file = 'COO_ROI_Dashboard_KPIs_Complete_12.xlsx'

        # Load all sheets
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
        st.error("❌ Excel file 'COO_ROI_Dashboard_KPIs_Complete_12.xlsx' not found.")
        st.stop()

# Load data
data = load_excel_data()

# Get latest month data
def get_latest_data(df, group_cols=None):
    """Get the latest month data"""
    if 'Month' in df.columns:
        latest_month = df['Month'].max()
        return df[df['Month'] == latest_month]
    return df

# Header
st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); color: white; padding: 25px 20px; border-radius: 0; margin: -25px -20px 25px -20px;">
        <h1 style="margin: 0;">📊 COO Operational Dashboard</h1>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Real-time Operational Metrics & KPI Tracking</p>
    </div>
""", unsafe_allow_html=True)

# Sidebar - Navigation
st.sidebar.markdown("## 📑 Navigation")
page = st.sidebar.radio(
    "Select View",
    ["Overview", "Efficiency & Cost", "Execution & Risk", "Workforce & Model"],
    label_visibility="collapsed"
)

# Sidebar - Filters
st.sidebar.markdown("## 🔧 Filters")

# Get unique months
role_months = sorted(data['Role_vs_Reality']['Month'].unique())
selected_month = st.sidebar.selectbox("Month", role_months, index=len(role_months)-1)

department = st.sidebar.selectbox("Department", ["All Departments", "HR", "Finance", "Operations", "Sales", "Engineering"])
real_time = st.sidebar.checkbox("Real-time Updates", value=True)

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Data Source:** COO_ROI_Dashboard_KPIs_Complete_12.xlsx")
st.sidebar.markdown(f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ==================== OVERVIEW TAB ====================
if page == "Overview":
    st.markdown("### Executive Summary - All KPIs")
    st.markdown(f"Data as of: **{selected_month}**")
    st.divider()

    # Filter data by selected month
    role_latest = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'] == selected_month]
    auto_latest = data['Automation_ROI'][data['Automation_ROI']['Month'] == selected_month]
    digital_latest = data['Digital_Index'][data['Digital_Index']['Month'] == selected_month]
    rework_latest = data['Process_Rework'][data['Process_Rework']['Month'] == selected_month]
    ftr_latest = data['FTR_Rate'][data['FTR_Rate']['Month'] == selected_month]
    adherence_latest = data['Adherence'][data['Adherence']['Month'] == selected_month]
    resilience_latest = data['Resilience'][data['Resilience']['Month'] == selected_month]
    escalation_latest = data['Escalation'][data['Escalation']['Month'] == selected_month]
    capacity_latest = data['Capacity'][data['Capacity']['Month'] == selected_month]
    model_latest = data['Model_Accuracy'][data['Model_Accuracy']['Month'] == selected_month]
    work_latest = data['Work_Models'][data['Work_Models']['Month'] == selected_month]
    collab_latest = data['Collaboration'][data['Collaboration']['Month'] == selected_month]

    # Calculate summary metrics
    avg_low_value = role_latest['Low_Value_Work_Percentage'].mean()
    total_opportunity = role_latest['Opportunity_Cost_Dollars'].sum()
    avg_roi = auto_latest['ROI_Percentage_6M'].mean() if len(auto_latest) > 0 else 0
    avg_friction = digital_latest['Friction_Index_Score'].mean()
    avg_rework = rework_latest['Rework_Cost_Percentage'].mean()
    avg_resilience = resilience_latest['Resilience_Score'].mean()
    avg_ftr = ftr_latest['FTR_Rate_Percentage'].mean()
    avg_adherence = adherence_latest['Adherence_Rate_Percentage'].mean()
    total_escalations = escalation_latest['Step_Exception_Count'].sum()
    burnout_count = capacity_latest[capacity_latest['Burnout_Risk_Flag'] == 'Yes'].shape[0]
    avg_model_acc = model_latest['Forecast_Accuracy_Percentage'].mean()
    avg_collab = collab_latest['Collaboration_Tools_Time_Hours'].mean()

    # Determine status colors
    def get_status(value, thresholds):
        """Get status based on thresholds. thresholds = {'green': (min, max), 'amber': ...}"""
        if value >= 80:
            return 'green', '✓ Good'
        elif value >= 60:
            return 'amber', '⚠ Monitor'
        else:
            return 'red', '✗ Critical'

    # Create KPI cards in grid
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        status_color, status_text = 'amber' if avg_low_value > 25 else 'green', '⚠ Monitor' if avg_low_value > 25 else '✓ Good'
        st.markdown(f"""
        <div class="kpi-card {status_color}">
            <div class="metric-label">Low-Value Work %</div>
            <div class="metric-value">{avg_low_value:.1f}%</div>
            <div class="trend trend-negative">↑ +2.1%</div>
            <div class="status-pill status-{'amber' if status_color == 'amber' else 'good'}">Monitor</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="kpi-card green">
            <div class="metric-label">Monthly Opportunity</div>
            <div class="metric-value">${total_opportunity:,.0f}</div>
            <div class="trend trend-positive">↑ +$12K</div>
            <div class="status-pill status-amber">High</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Avg Friction Index</div>
            <div class="metric-value">{avg_friction:.1f}</div>
            <div class="trend trend-negative">↓ -3.2</div>
            <div class="status-pill status-amber">Improve</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        rework_status = 'red' if avg_rework > 5 else 'amber' if avg_rework > 3 else 'green'
        st.markdown(f"""
        <div class="kpi-card {rework_status}">
            <div class="metric-label">Avg Rework Cost %</div>
            <div class="metric-value">{avg_rework:.1f}%</div>
            <div class="trend trend-negative">↑ +0.3%</div>
            <div class="status-pill status-{'red' if rework_status == 'red' else 'amber'}">Alert</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Resilience Score</div>
            <div class="metric-value">{avg_resilience:.1f}/10</div>
            <div class="trend trend-positive">↑ +0.5</div>
            <div class="status-pill status-amber">Fair</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        ftr_status = 'red' if avg_ftr < 70 else 'amber' if avg_ftr < 85 else 'green'
        st.markdown(f"""
        <div class="kpi-card {ftr_status}">
            <div class="metric-label">FTR Rate</div>
            <div class="metric-value">{avg_ftr:.1f}%</div>
            <div class="trend trend-positive">↑ +2.3%</div>
            <div class="status-pill status-{'green' if ftr_status == 'green' else 'amber'}">Improve</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        adh_status = 'red' if avg_adherence < 70 else 'amber' if avg_adherence < 85 else 'green'
        st.markdown(f"""
        <div class="kpi-card {adh_status}">
            <div class="metric-label">Process Adherence</div>
            <div class="metric-value">{avg_adherence:.1f}%</div>
            <div class="trend trend-negative">↓ -1.2%</div>
            <div class="status-pill status-{'amber' if adh_status != 'green' else 'good'}">Monitor</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class="kpi-card red">
            <div class="metric-label">Escalations</div>
            <div class="metric-value">{total_escalations:.0f}</div>
            <div class="trend trend-negative">↑ +8%</div>
            <div class="status-pill status-red">Alert</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        burnout_status = 'red' if burnout_count > 20 else 'amber' if burnout_count > 10 else 'green'
        st.markdown(f"""
        <div class="kpi-card {burnout_status}">
            <div class="metric-label">At-Risk Employees</div>
            <div class="metric-value">{burnout_count}</div>
            <div class="trend trend-negative">↑ +3</div>
            <div class="status-pill status-red">Critical</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Model Accuracy</div>
            <div class="metric-value">{avg_model_acc:.1f}%</div>
            <div class="trend trend-positive">↑ +4.1%</div>
            <div class="status-pill status-amber">Improving</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        avg_roi_display = avg_roi if avg_roi > 0 else 0
        st.markdown(f"""
        <div class="kpi-card green">
            <div class="metric-label">Avg Automation ROI</div>
            <div class="metric-value">{avg_roi_display:.0f}%</div>
            <div class="trend trend-positive">↑ +125%</div>
            <div class="status-pill status-good">Excellent</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class="kpi-card red">
            <div class="metric-label">Collab Hours</div>
            <div class="metric-value">{avg_collab:.1f} hrs</div>
            <div class="trend trend-negative">↑ +0.3 hrs</div>
            <div class="status-pill status-red">High</div>
        </div>
        """, unsafe_allow_html=True)

# ==================== EFFICIENCY & COST TAB ====================
elif page == "Efficiency & Cost":
    st.markdown("### 💡 Efficiency & Cost Management")
    st.divider()

    # Filter data
    role_data = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'] == selected_month]
    auto_data = data['Automation_ROI'][data['Automation_ROI']['Month'] == selected_month]
    digital_data = data['Digital_Index'][data['Digital_Index']['Month'] == selected_month]
    rework_data = data['Process_Rework'][data['Process_Rework']['Month'] == selected_month]

    # ===== Section 1: Role vs Reality =====
    st.markdown("#### 💰 Role vs. Reality Analysis")
    st.markdown("*Opportunity Cost (%) & Dollar Value | By: Employee Role, Department, Month*")

    col1, col2 = st.columns([1, 1])
    with col1:
        avg_low = role_data['Low_Value_Work_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Low-Value Work</div>
            <div class="metric-value">{avg_low:.1f}%</div>
            <div class="trend trend-positive">↑ +2.1%</div>
            <div class="status-pill status-amber">Review</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        total_impact = role_data['Opportunity_Cost_Dollars'].sum()
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Total Opportunity Cost</div>
            <div class="metric-value">${total_impact:,.0f}</div>
            <div class="trend trend-positive">↑ +$8.5K</div>
            <div class="status-pill status-amber">High</div>
        </div>
        """, unsafe_allow_html=True)

    # Layout: Big chart left, small tiles right
    col_chart, col_tiles = st.columns([2, 1])

    with col_chart:
        st.markdown("**Low-Value Work by Role (Bar Chart)**")
        role_summary = role_data.groupby('Role').agg({
            'Low_Value_Work_Percentage': 'mean',
            'Opportunity_Cost_Dollars': 'sum'
        }).sort_values('Low_Value_Work_Percentage', ascending=False).head(8)

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=role_summary.index,
            y=role_summary['Low_Value_Work_Percentage'],
            marker_color=['#dc2626', '#f59e0b', '#f59e0b', '#f59e0b', '#0891b2', '#059669', '#059669', '#059669'][:len(role_summary)],
            text=role_summary['Low_Value_Work_Percentage'].apply(lambda x: f"{x:.0f}%"),
            textposition='outside'
        ))
        fig.update_layout(height=300, margin=dict(l=0, r=0, t=0, b=0), xaxis_title="Role", yaxis_title="Low-Value Work %", plot_bgcolor="rgba(0,0,0,0)", xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        st.markdown("**Key Insights**")
        highest = role_summary.iloc[0]
        st.metric("Highest Waste", highest.name, f"{highest['Low_Value_Work_Percentage']:.0f}%")
        annual_opp = role_data['Opportunity_Cost_Dollars'].sum() * 12
        st.metric("Annual Opportunity", f"${annual_opp:,.0f}", "If Optimized")

    # Data table
    st.markdown("**Detailed Breakdown by Role**")
    role_display = role_data.groupby('Role').agg({
        'Low_Value_Work_Percentage': 'mean',
        'High_Value_Work_Percentage': 'mean',
        'Opportunity_Cost_Dollars': 'sum'
    }).round(2).reset_index()
    role_display.columns = ['Role', 'Low-Value %', 'High-Value %', 'Total Opp Cost ($)']
    st.dataframe(role_display, use_container_width=True, hide_index=True)

    st.divider()

    # ===== Section 2: Automation ROI =====
    st.markdown("#### 🤖 Automation ROI Potential")
    st.markdown("*ROI % & Time Savings | By: Task Type, Automation Project, Quarter*")

    col1, col2 = st.columns([1, 1])
    with col1:
        avg_roi = auto_data['ROI_Percentage_6M'].mean()
        st.markdown(f"""
        <div class="kpi-card green">
            <div class="metric-label">Avg ROI (6M)</div>
            <div class="metric-value">{avg_roi:.0f}%</div>
            <div class="trend trend-positive">↑ +128%</div>
            <div class="status-pill status-good">Excellent</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        total_savings = auto_data['Monthly_Hours_Saved'].sum()
        st.markdown(f"""
        <div class="kpi-card green">
            <div class="metric-label">Monthly Hours Saved</div>
            <div class="metric-value">{total_savings:.0f} hrs</div>
            <div class="trend trend-positive">↑ +142 hrs</div>
            <div class="status-pill status-good">High</div>
        </div>
        """, unsafe_allow_html=True)

    col_chart, col_tiles = st.columns([2, 1])

    with col_chart:
        st.markdown("**Top Automation Projects by ROI (Ranked Bar)**")
        auto_sorted = auto_data.sort_values('ROI_Percentage_6M', ascending=True).tail(8)
        fig = px.bar(auto_sorted, y='Process_Name', x='ROI_Percentage_6M', orientation='h', color='ROI_Percentage_6M', color_continuous_scale='Greens')
        fig.update_layout(height=300, margin=dict(l=0, r=0, t=0, b=0), showlegend=False, xaxis_title="ROI %", yaxis_title="", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        st.markdown("**Key Metrics**")
        if len(auto_sorted) > 0:
            top = auto_sorted.iloc[-1]
            st.metric("Top Project", top['Process_Name'], f"{top['ROI_Percentage_6M']:.0f}% ROI")
        total_savings_cost = auto_data['Monthly_Cost_Savings'].sum() * 6
        st.metric("6-Month Savings", f"${total_savings_cost:,.0f}", "Est. value")

    st.markdown("**Project Details**")
    auto_display = auto_data[['Process_Name', 'Task_Type', 'Monthly_Hours_Saved', 'Monthly_Cost_Savings', 'ROI_Percentage_6M']].copy()
    auto_display.columns = ['Process', 'Type', 'Hours Saved', 'Monthly Savings ($)', 'ROI %']
    auto_display = auto_display.round(2)
    st.dataframe(auto_display, use_container_width=True, hide_index=True)

    st.divider()

    # ===== Section 3: Digital Workplace Index =====
    st.markdown("#### 📱 Digital Workplace Index")
    st.markdown("*Friction Index Score (0-100) | By: Department, Team, Application, Week*")

    col1, col2 = st.columns([1, 1])
    with col1:
        avg_friction = digital_data['Friction_Index_Score'].mean()
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Avg Friction Index</div>
            <div class="metric-value">{avg_friction:.1f}</div>
            <div class="trend trend-negative">↓ -4.5</div>
            <div class="status-pill status-amber">Improve</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        avg_latency = digital_data['App_Response_Latency_Sec'].mean()
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Avg Response Latency</div>
            <div class="metric-value">{avg_latency:.2f}s</div>
            <div class="trend trend-negative">↓ -0.3s</div>
            <div class="status-pill status-amber">Monitor</div>
        </div>
        """, unsafe_allow_html=True)

    col_chart, col_tiles = st.columns([2, 1])

    with col_chart:
        st.markdown("**Friction Index by Department (Bar Chart)**")
        dept_friction = digital_data.groupby('Department')['Friction_Index_Score'].mean().sort_values(ascending=False)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=dept_friction.index, y=dept_friction.values, marker_color='#1e40af', text=dept_friction.apply(lambda x: f"{x:.1f}"), textposition='outside'))
        fig.update_layout(height=300, margin=dict(l=0, r=0, t=0, b=0), xaxis_title="Department", yaxis_title="Friction Score", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        st.markdown("**Department Status**")
        worst = dept_friction.index[0]
        st.metric("Worst Dept", worst, f"Score: {dept_friction.iloc[0]:.1f}")
        best = dept_friction.index[-1]
        st.metric("Best Dept", best, f"Score: {dept_friction.iloc[-1]:.1f}")

    st.markdown("**App Friction Analysis**")
    app_friction = digital_data.groupby('Primary_Friction_App')['Friction_Index_Score'].mean().sort_values(ascending=False)
    st.dataframe(app_friction.to_frame('Friction Score'), use_container_width=True)

    st.divider()

    # ===== Section 4: Process Rework Cost =====
    st.markdown("#### ♻️ Process Rework Cost %")
    st.markdown("*Rework Cost % & Dollar Impact | By: Process, Department, Month*")

    col1, col2 = st.columns([1, 1])
    with col1:
        avg_rework_pct = rework_data['Rework_Cost_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-card {'red' if avg_rework_pct > 5 else 'amber' if avg_rework_pct > 3 else 'green'}">
            <div class="metric-label">Avg Rework Cost %</div>
            <div class="metric-value">{avg_rework_pct:.1f}%</div>
            <div class="trend trend-negative">↑ +0.5%</div>
            <div class="status-pill status-{'red' if avg_rework_pct > 5 else 'amber'}>Alert</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        total_rework = rework_data['Rework_Cost_Dollars'].sum()
        quarterly = total_rework * 3
        st.markdown(f"""
        <div class="kpi-card {'red' if quarterly > 30000 else 'amber'}">
            <div class="metric-label">Monthly Rework Cost</div>
            <div class="metric-value">${total_rework:,.0f}</div>
            <div class="trend trend-negative">↑ +$2.5K</div>
            <div class="status-pill status-red">Rising</div>
        </div>
        """, unsafe_allow_html=True)

    col_chart, col_tiles = st.columns([2, 1])

    with col_chart:
        st.markdown("**Rework Cost by Process (Ranked Bar)**")
        process_rework = rework_data.groupby('Process_Name').agg({
            'Rework_Cost_Percentage': 'mean',
            'Rework_Cost_Dollars': 'sum'
        }).sort_values('Rework_Cost_Dollars', ascending=True).tail(8)

        fig = px.bar(process_rework.reset_index(), y='Process_Name', x='Rework_Cost_Dollars', orientation='h', color='Rework_Cost_Percentage', color_continuous_scale='Reds')
        fig.update_layout(height=300, margin=dict(l=0, r=0, t=0, b=0), showlegend=False, xaxis_title="Rework Cost ($)", yaxis_title="", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        st.markdown("**Cost Analysis**")
        worst_process = process_rework['Rework_Cost_Dollars'].idxmax() if len(process_rework) > 0 else "N/A"
        st.metric("Worst Process", worst_process, f"{process_rework.loc[worst_process, 'Rework_Cost_Percentage']:.1f}% rework")
        annual_cost = rework_data['Rework_Cost_Dollars'].sum() * 12
        st.metric("Annual Cost", f"${annual_cost:,.0f}", "If not fixed")

    st.markdown("**Process Breakdown**")
    rework_display = rework_data.groupby('Process_Name').agg({
        'Rework_Cost_Percentage': 'mean',
        'Rework_Cost_Dollars': 'sum',
        'Rework_Transaction_Count': 'sum'
    }).round(2).reset_index()
    rework_display.columns = ['Process', 'Rework %', 'Total Cost ($)', 'Rework Count']
    st.dataframe(rework_display, use_container_width=True, hide_index=True)

# ==================== EXECUTION & RISK TAB ====================
elif page == "Execution & Risk":
    st.markdown("### ⚙️ Execution & Risk Management")
    st.divider()

    # Filter data
    ftr_data = data['FTR_Rate'][data['FTR_Rate']['Month'] == selected_month]
    adherence_data = data['Adherence'][data['Adherence']['Month'] == selected_month]
    resilience_data = data['Resilience'][data['Resilience']['Month'] == selected_month]
    escalation_data = data['Escalation'][data['Escalation']['Month'] == selected_month]

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        avg_res = resilience_data['Resilience_Score'].mean()
        st.markdown(f"""
        <div class="kpi-card {'amber' if avg_res < 7 else 'green'}">
            <div class="metric-label">Resilience Score</div>
            <div class="metric-value">{avg_res:.1f}/10</div>
            <div class="trend trend-positive">↑ +0.4</div>
            <div class="status-pill status-amber">Fair</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        avg_ftr = ftr_data['FTR_Rate_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-card {'green' if avg_ftr > 85 else 'amber' if avg_ftr > 70 else 'red'}">
            <div class="metric-label">FTR Rate</div>
            <div class="metric-value">{avg_ftr:.1f}%</div>
            <div class="trend trend-positive">↑ +1.8%</div>
            <div class="status-pill status-{'green' if avg_ftr > 85 else 'amber'}">Improve</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        avg_adh = adherence_data['Adherence_Rate_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-card {'amber' if avg_adh < 85 else 'green'}">
            <div class="metric-label">Process Adherence</div>
            <div class="metric-value">{avg_adh:.1f}%</div>
            <div class="trend trend-negative">↓ -1.3%</div>
            <div class="status-pill status-amber">Monitor</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        total_esc = escalation_data['Step_Exception_Count'].sum()
        st.markdown(f"""
        <div class="kpi-card red">
            <div class="metric-label">Total Escalations</div>
            <div class="metric-value">{total_esc:.0f}</div>
            <div class="trend trend-negative">↑ +9%</div>
            <div class="status-pill status-red">Alert</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # ===== Resilience Section =====
    st.markdown("#### 🛡️ Operational Resilience Score")
    col_chart, col_tiles = st.columns([2, 1])
    with col_chart:
        st.markdown("**Critical Tasks Risk Level (Ranked Bar)**")
        res_summary = resilience_data.groupby('Critical_Task').agg({
            'Resilience_Score': 'mean',
            'Risk_Percentage': 'mean'
        }).sort_values('Risk_Percentage', ascending=True).tail(6)

        fig = px.bar(res_summary.reset_index(), y='Critical_Task', x='Risk_Percentage', orientation='h', color='Risk_Percentage', color_continuous_scale='Reds')
        fig.update_layout(height=300, showlegend=False, plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        highest_risk = resilience_data.loc[resilience_data['Risk_Percentage'].idxmax()]
        st.metric("Highest Risk Task", highest_risk['Critical_Task'], f"{highest_risk['Risk_Percentage']:.0f}% risk")
        critical_count = resilience_data[resilience_data['Risk_Level'] == 'Critical'].shape[0]
        st.metric("Critical Risk Items", critical_count, "Need Action")

    st.markdown("**Resilience Details**")
    res_display = resilience_data[['Critical_Task', 'Department', 'FTE_Coverage_Count', 'Risk_Percentage', 'Risk_Level']].copy()
    res_display.columns = ['Task', 'Dept', 'FTE Coverage', 'Risk %', 'Risk Level']
    st.dataframe(res_display.drop_duplicates(), use_container_width=True, hide_index=True)

    st.divider()

    # ===== FTR Section =====
    st.markdown("#### ✅ First-Time-Right (FTR) Rate")
    col_chart, col_tiles = st.columns([2, 1])
    with col_chart:
        st.markdown("**Process FTR % vs Target (Comparison)**")
        ftr_summary = ftr_data.groupby('Process').agg({
            'FTR_Rate_Percentage': 'mean',
            'Target_FTR_Rate': 'first'
        }).reset_index()

        fig = go.Figure()
        fig.add_trace(go.Bar(name='Actual FTR %', y=ftr_summary['Process'], x=ftr_summary['FTR_Rate_Percentage'], orientation='h', marker_color='#3b82f6'))
        fig.add_trace(go.Bar(name='Target %', y=ftr_summary['Process'], x=ftr_summary['Target_FTR_Rate'], orientation='h', marker_color='#d1d5db'))
        fig.update_layout(height=300, barmode='group', plot_bgcolor="rgba(0,0,0,0)", legend=dict(x=0.6, y=0.95))
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        best = ftr_summary.loc[ftr_summary['FTR_Rate_Percentage'].idxmax()]
        st.metric("Best Process", best['Process'], f"{best['FTR_Rate_Percentage']:.0f}%")
        worst = ftr_summary.loc[ftr_summary['FTR_Rate_Percentage'].idxmin()]
        st.metric("Worst Process", worst['Process'], f"{worst['FTR_Rate_Percentage']:.0f}%")

    st.markdown("**FTR Details by Process**")
    ftr_display = ftr_data.groupby('Process').agg({
        'FTR_Rate_Percentage': 'mean',
        'Error_Rate_Percentage': 'mean',
        'Target_FTR_Rate': 'first'
    }).round(2).reset_index()
    ftr_display.columns = ['Process', 'FTR %', 'Error %', 'Target %']
    st.dataframe(ftr_display, use_container_width=True, hide_index=True)

# ==================== WORKFORCE & MODEL TAB ====================
elif page == "Workforce & Model":
    st.markdown("### 👥 Workforce & Scalable Operating Model")
    st.divider()

    # Filter data
    capacity_data = data['Capacity'][data['Capacity']['Month'] == selected_month]
    model_acc_data = data['Model_Accuracy'][data['Model_Accuracy']['Month'] == selected_month]
    work_models_data = data['Work_Models'][data['Work_Models']['Month'] == selected_month]
    collab_data = data['Collaboration'][data['Collaboration']['Month'] == selected_month]

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        burnout_count = capacity_data[capacity_data['Burnout_Risk_Flag'] == 'Yes'].shape[0]
        st.markdown(f"""
        <div class="kpi-card {'red' if burnout_count > 15 else 'amber' if burnout_count > 8 else 'green'}">
            <div class="metric-label">At-Risk Employees</div>
            <div class="metric-value">{burnout_count}</div>
            <div class="trend trend-negative">↑ +2</div>
            <div class="status-pill status-{'red' if burnout_count > 15 else 'amber' if burnout_count > 8 else 'good'}">Critical</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        avg_model = model_acc_data['Forecast_Accuracy_Percentage'].mean()
        st.markdown(f"""
        <div class="kpi-card amber">
            <div class="metric-label">Model Accuracy</div>
            <div class="metric-value">{avg_model:.0f}%</div>
            <div class="trend trend-positive">↑ +3.2%</div>
            <div class="status-pill status-amber">Improving</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        remote_data = work_models_data[work_models_data['Work_Model'] == 'Remote']
        if len(remote_data) > 0:
            remote_out = remote_data['Output_Per_Hour'].mean()
        else:
            remote_out = 0
        st.markdown(f"""
        <div class="kpi-card green">
            <div class="metric-label">Remote Output/Hr</div>
            <div class="metric-value">{remote_out:.2f}</div>
            <div class="trend trend-positive">↑ +0.4</div>
            <div class="status-pill status-good">Good</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        avg_collab = collab_data['Collaboration_Tools_Time_Hours'].mean()
        st.markdown(f"""
        <div class="kpi-card {'red' if avg_collab > 4 else 'amber' if avg_collab > 3 else 'green'}">
            <div class="metric-label">Avg Collab Hours</div>
            <div class="metric-value">{avg_collab:.1f} hrs</div>
            <div class="trend trend-positive">↓ -0.2 hrs</div>
            <div class="status-pill status-amber">High</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.markdown("**📊 Detailed KPI Sections**")

    # ===== Capacity Section =====
    st.markdown("#### 🔥 Hidden Capacity & Burnout Risk")
    col_chart, col_tiles = st.columns([2, 1])
    with col_chart:
        st.markdown("**Team Capacity Utilization (Bar Chart)**")
        cap_summary = capacity_data.groupby('Department')['Capacity_Utilization_Percentage'].mean().sort_values()
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=cap_summary.values,
            y=cap_summary.index,
            orientation='h',
            marker_color=cap_summary.apply(lambda x: '#dc2626' if x > 110 else '#f59e0b' if x > 90 else '#059669').values,
            text=cap_summary.apply(lambda x: f"{x:.0f}%").values,
            textposition='outside'
        ))
        fig.add_vline(x=100, line_dash="dash", line_color="red")
        fig.update_layout(height=300, showlegend=False, xaxis_title="Capacity %", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        highest_cap = capacity_data.loc[capacity_data['Capacity_Utilization_Percentage'].idxmax()]
        st.metric("Most Overloaded", highest_cap['Department'], f"{highest_cap['Capacity_Utilization_Percentage']:.0f}%")
        at_risk = capacity_data[capacity_data['Capacity_Utilization_Percentage'] > 110].shape[0]
        st.metric("Over 110% Cap", at_risk, "At Risk")

    st.markdown("**Capacity Details**")
    cap_display = capacity_data[['Department', 'Role', 'Capacity_Utilization_Percentage', 'Burnout_Risk_Flag', 'Capacity_Status']].drop_duplicates()
    cap_display.columns = ['Dept', 'Role', 'Capacity %', 'Burnout Risk', 'Status']
    st.dataframe(cap_display, use_container_width=True, hide_index=True)

    st.divider()

    # ===== Work Models Section =====
    st.markdown("#### 💼 Work Models Effectiveness")
    col_chart, col_tiles = st.columns([2, 1])
    with col_chart:
        st.markdown("**Output by Work Model (Comparison)**")
        work_summary = work_models_data.groupby('Work_Model').agg({
            'Output_Per_Hour': 'mean',
            'Cost_Per_Transaction': 'mean'
        }).reset_index()

        fig = px.bar(work_summary, x='Work_Model', y='Output_Per_Hour', color='Output_Per_Hour', color_continuous_scale='Blues')
        fig.update_layout(height=250, showlegend=False, xaxis_title="Work Model", yaxis_title="Output/Hr", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_tiles:
        st.markdown("**Model Comparison**")
        best = work_summary.loc[work_summary['Output_Per_Hour'].idxmax()]
        st.metric("Best Model", best['Work_Model'], f"{best['Output_Per_Hour']:.2f} output/hr")
        lowest = work_summary.loc[work_summary['Cost_Per_Transaction'].idxmin()]
        st.metric("Lowest Cost", lowest['Work_Model'], f"${lowest['Cost_Per_Transaction']:.0f}/txn")

    st.markdown("**Work Model Metrics**")
    work_display = work_summary.round(2)
    work_display.columns = ['Work Model', 'Output/Hr', 'Cost/Txn']
    st.dataframe(work_display, use_container_width=True, hide_index=True)

# Footer
st.divider()
st.markdown(f"""
    <div style="text-align: center; padding: 20px; color: #6b7280; font-size: 12px;">
        <strong>COO Dashboard v3.0</strong> | Data: {selected_month} | 
        <strong>Source:</strong> COO_ROI_Dashboard_KPIs_Complete_12.xlsx
    </div>
""", unsafe_allow_html=True)
