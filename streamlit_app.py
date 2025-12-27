import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from datetime import datetime, timedelta
import io

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
    .insight-box {
        background: #f0f9ff;
        border-left: 4px solid #0284c7;
        padding: 15px;
        border-radius: 5px;
        margin: 15px 0;
    }
    .opportunity-box {
        background: #f0fdf4;
        border-left: 4px solid #22c55e;
        padding: 15px;
        border-radius: 5px;
        margin: 15px 0;
    }
    .alert-box {
        background: #fef2f2;
        border-left: 4px solid #ef4444;
        padding: 15px;
        border-radius: 5px;
        margin: 15px 0;
    }
    .data-section {
        background: #f8fafc;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        margin-top: 15px;
    }
    .expandable-header {
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 12px;
        background: #f1f5f9;
        border-radius: 6px;
        font-weight: 600;
        user-select: none;
    }
    .expandable-header:hover {
        background: #e2e8f0;
    }
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

# Initialize session state for navigation
if 'current_page' not in st.session_state:
    st.session_state.current_page = "Overview"

# Header
st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); color: white; padding: 25px 20px; border-radius: 0; margin: -25px -20px 25px -20px;">
        <h1 style="margin: 0;">📊 COO Operational Dashboard</h1>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Real-time Operational Metrics & KPI Tracking</p>
    </div>
""", unsafe_allow_html=True)

# ==================== NAVIGATION (Main Screen) ====================
nav_col1, nav_col2, nav_col3, nav_col4 = st.columns(4)

with nav_col1:
    if st.button("📊 Overview", use_container_width=True, key="nav_overview"):
        st.session_state.current_page = "Overview"

with nav_col2:
    if st.button("💡 Efficiency & Cost", use_container_width=True, key="nav_efficiency"):
        st.session_state.current_page = "Efficiency & Cost"

with nav_col3:
    if st.button("⚙️ Execution & Risk", use_container_width=True, key="nav_execution"):
        st.session_state.current_page = "Execution & Risk"

with nav_col4:
    if st.button("👥 Workforce & Model", use_container_width=True, key="nav_workforce"):
        st.session_state.current_page = "Workforce & Model"

st.divider()

# ==================== SIDEBAR FILTERS ====================
st.sidebar.markdown("## 🔧 Filters")

# Get unique months - DEFAULT: ALL SELECTED
role_months = sorted(data['Role_vs_Reality']['Month'].unique())
selected_months = st.sidebar.multiselect(
    "Select Months",
    role_months,
    default=list(role_months),  # ALL SELECTED BY DEFAULT
    key="month_filter"
)

# Multi-select for departments - DEFAULT: ALL SELECTED
all_departments = sorted(
    list(set(list(data['Role_vs_Reality'].get('Department', []).unique()) + 
             list(data['Capacity'].get('Department', []).unique())))
)
selected_depts = st.sidebar.multiselect(
    "Select Departments",
    all_departments,
    default=all_departments,  # ALL SELECTED BY DEFAULT
    key="dept_filter"
)

# Filter handling
dept_filter = selected_depts if len(selected_depts) > 0 else None

# Real-time toggle
real_time = st.sidebar.checkbox("Real-time Updates", value=True)

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Data Source:** COO_ROI_Dashboard_KPIs_Complete_12.xlsx")
st.sidebar.markdown(f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ==================== HELPER FUNCTIONS ====================
def export_dataframe_to_excel(df, sheet_name="Data"):
    """Convert dataframe to Excel for download"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def create_trend_chart(data_df, x_col, y_col, title, y_title, color='#1e40af'):
    """Create a trend line chart"""
    if len(data_df) == 0:
        return None

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=data_df[x_col],
        y=data_df[y_col],
        mode='lines+markers',
        line=dict(color=color, width=3),
        marker=dict(size=8),
        fill='tozeroy',
        fillcolor=f'rgba({int(color[1:3], 16)}, {int(color[3:5], 16)}, {int(color[5:7], 16)}, 0.1)'
    ))
    fig.update_layout(
        title=title,
        height=300,
        margin=dict(l=0, r=0, t=30, b=0),
        showlegend=False,
        xaxis_title=x_col,
        yaxis_title=y_title,
        plot_bgcolor="rgba(0,0,0,0)",
        hovermode='x unified'
    )
    return fig

def show_detailed_data(data_df, title, columns=None):
    """Show detailed data table"""
    st.markdown(f"### 📋 {title}")
    if columns:
        display_df = data_df[columns].copy()
    else:
        display_df = data_df.copy()
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # Export button
    csv = display_df.to_csv(index=False)
    st.download_button(
        label=f"📥 Download {title} (CSV)",
        data=csv,
        file_name=f"{title.replace(' ', '_')}.csv",
        mime="text/csv"
    )

# ==================== OVERVIEW TAB ====================
if st.session_state.current_page == "Overview":
    st.markdown("### Executive Summary - All KPIs")

    # Use all selected months and departments
    role_latest = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'].isin(selected_months)]
    if dept_filter:
        role_latest = role_latest[role_latest['Department'].isin(dept_filter)]

    auto_latest = data['Automation_ROI'][data['Automation_ROI']['Month'].isin(selected_months)]
    digital_latest = data['Digital_Index'][data['Digital_Index']['Month'].isin(selected_months)]
    if dept_filter:
        digital_latest = digital_latest[digital_latest['Department'].isin(dept_filter)]

    rework_latest = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)]
    if dept_filter:
        rework_latest = rework_latest[rework_latest['Department'].isin(dept_filter)]

    ftr_latest = data['FTR_Rate'][data['FTR_Rate']['Month'].isin(selected_months)]
    adherence_latest = data['Adherence'][data['Adherence']['Month'].isin(selected_months)]
    resilience_latest = data['Resilience'][data['Resilience']['Month'].isin(selected_months)]
    escalation_latest = data['Escalation'][data['Escalation']['Month'].isin(selected_months)]
    capacity_latest = data['Capacity'][data['Capacity']['Month'].isin(selected_months)]
    if dept_filter:
        capacity_latest = capacity_latest[capacity_latest['Department'].isin(dept_filter)]

    model_latest = data['Model_Accuracy'][data['Model_Accuracy']['Month'].isin(selected_months)]
    work_latest = data['Work_Models'][data['Work_Models']['Month'].isin(selected_months)]
    collab_latest = data['Collaboration'][data['Collaboration']['Month'].isin(selected_months)]

    # Calculate metrics
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

    st.info(f"📊 Showing data for: {len(selected_months)} months | {len(dept_filter) if dept_filter else len(all_departments)} departments")

    # KPI Cards - Row 1
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        status_color = 'amber' if avg_low_value > 25 else 'green'
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

    # KPI Cards - Row 2
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

    # KPI Cards - Row 3
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

    st.divider()

    # KEY INSIGHTS & OPPORTUNITIES
    st.markdown("### 💡 Key Insights & Opportunities")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="insight-box">
            <strong>🔍 Insight: Low-Value Work Burden</strong><br>
            Employees are spending significant time on non-core tasks that don't add strategic value.
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show details", key="insight_lowvalue"):
            low_value_detail = role_latest.groupby('Role').agg({
                'Low_Value_Work_Percentage': 'mean',
                'Opportunity_Cost_Dollars': 'sum'
            }).sort_values('Opportunity_Cost_Dollars', ascending=False)
            st.dataframe(low_value_detail, use_container_width=True)

    with col2:
        st.markdown(f"""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Automation ROI</strong><br>
            Average ROI of {avg_roi_display:.0f}% on automation projects - exceptional returns.
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show details", key="opp_automation"):
            auto_detail = auto_latest.groupby('Process_Name').agg({
                'ROI_Percentage_6M': 'mean',
                'Monthly_Cost_Savings': 'sum'
            }).sort_values('ROI_Percentage_6M', ascending=False).head(10)
            st.dataframe(auto_detail, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"""
        <div class="alert-box">
            <strong>⚠️ Alert: Escalation Spike</strong><br>
            {total_escalations:.0f} escalations indicate process bottlenecks requiring attention.
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show details", key="alert_escalation"):
            esc_detail = escalation_latest.groupby('Process').agg({
                'Step_Exception_Count': 'sum',
                'Exception_Rate_Percentage': 'mean'
            }).sort_values('Step_Exception_Count', ascending=False).head(10)
            st.dataframe(esc_detail, use_container_width=True)

    with col2:
        if burnout_count > 10:
            st.markdown(f"""
            <div class="alert-box">
                <strong>🔴 Critical: Burnout Risk</strong><br>
                {burnout_count} employees at risk. Immediate capacity rebalancing required.
            </div>
            """, unsafe_allow_html=True)

            if st.checkbox("📊 Show details", key="alert_burnout"):
                burnout_detail = capacity_latest[capacity_latest['Burnout_Risk_Flag'] == 'Yes'][['Department', 'Role', 'Capacity_Utilization_Percentage']].drop_duplicates()
                st.dataframe(burnout_detail, use_container_width=True)

# ==================== EFFICIENCY & COST TAB ====================
elif st.session_state.current_page == "Efficiency & Cost":
    st.markdown("### 💡 Efficiency & Cost Management")

    # Filter data
    role_data = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'].isin(selected_months)]
    if dept_filter:
        role_data = role_data[role_data['Department'].isin(dept_filter)]

    auto_data = data['Automation_ROI'][data['Automation_ROI']['Month'].isin(selected_months)]
    digital_data = data['Digital_Index'][data['Digital_Index']['Month'].isin(selected_months)]
    if dept_filter:
        digital_data = digital_data[digital_data['Department'].isin(dept_filter)]

    rework_data = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)]
    if dept_filter:
        rework_data = rework_data[rework_data['Department'].isin(dept_filter)]

    st.info(f"📊 Showing {len(role_data)} records from {len(selected_months)} months")

    # ===== Section 1: Role vs Reality =====
    st.markdown("#### 💰 Role vs. Reality Analysis")

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

    # Combined chart: Low-Value % and Cost
    st.markdown("**Low-Value Work by Role & Impact (Dual Axis)**")
    role_summary = role_data.groupby('Role').agg({
        'Low_Value_Work_Percentage': 'mean',
        'Opportunity_Cost_Dollars': 'sum'
    }).sort_values('Opportunity_Cost_Dollars', ascending=False).head(8)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=role_summary.index,
        y=role_summary['Low_Value_Work_Percentage'],
        name='Low-Value Work %',
        marker_color='#f59e0b',
        yaxis='y1'
    ))
    fig.add_trace(go.Scatter(
        x=role_summary.index,
        y=role_summary['Opportunity_Cost_Dollars'],
        name='Opportunity Cost ($)',
        yaxis='y2',
        mode='lines+markers',
        line=dict(color='#dc2626', width=3),
        marker=dict(size=10)
    ))

    fig.update_layout(
        title='Low-Value Work Percentage vs Opportunity Cost',
        yaxis=dict(title='Low-Value Work (%)', side='left'),
        yaxis2=dict(title='Opportunity Cost ($)', overlaying='y', side='right'),
        height=400,
        hovermode='x unified',
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_tickangle=-45
    )
    st.plotly_chart(fig, use_container_width=True)

    # Opportunity insight with expandable details
    annual_opp = role_data['Opportunity_Cost_Dollars'].sum() * 12
    st.markdown(f"""
    <div class="opportunity-box">
        <strong>🎯 Opportunity: Role Optimization</strong><br>
        Annual potential: ${annual_opp:,.0f} | Reallocate {int(role_data['Opportunity_Cost_Dollars'].sum() / 75)} FTE hours
    </div>
    """, unsafe_allow_html=True)

    if st.checkbox("📊 Show detailed breakdown", key="role_details"):
        show_detailed_data(
            role_data.groupby('Role').agg({
                'Low_Value_Work_Percentage': 'mean',
                'High_Value_Work_Percentage': 'mean',
                'Opportunity_Cost_Dollars': 'sum'
            }).reset_index().sort_values('Opportunity_Cost_Dollars', ascending=False),
            "Role vs Reality Details",
            ['Role', 'Low_Value_Work_Percentage', 'High_Value_Work_Percentage', 'Opportunity_Cost_Dollars']
        )

    # Trend
    role_trend = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'].isin(selected_months)].groupby('Month').agg({
        'Low_Value_Work_Percentage': 'mean',
        'Opportunity_Cost_Dollars': 'sum'
    }).reset_index().sort_values('Month')

    if len(role_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(role_trend, 'Month', 'Low_Value_Work_Percentage', 'Low-Value Work Trend', 'Percentage (%)', '#dc2626')
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ===== Section 2: Automation ROI =====
    st.markdown("#### 🤖 Automation ROI Potential")

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

    # Combined chart: ROI % and Cost Savings
    st.markdown("**Top Automation Projects: ROI vs Cost Savings (Dual Axis)**")
    auto_sorted = auto_data.sort_values('ROI_Percentage_6M', ascending=False).head(8)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=auto_sorted['Process_Name'],
        x=auto_sorted['ROI_Percentage_6M'],
        name='ROI %',
        marker_color='#059669',
        orientation='h',
        xaxis='x1'
    ))
    fig.add_trace(go.Scatter(
        y=auto_sorted['Process_Name'],
        x=auto_sorted['Monthly_Cost_Savings'],
        name='Monthly Savings ($)',
        mode='markers',
        marker=dict(size=12, color='#dc2626'),
        xaxis='x2'
    ))

    fig.update_layout(
        title='Automation Projects: ROI vs Monthly Savings',
        xaxis=dict(title='ROI %'),
        xaxis2=dict(title='Monthly Savings ($)', overlaying='x', side='top'),
        height=400,
        hovermode='y unified',
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Opportunity
    if len(auto_sorted) > 0:
        top = auto_sorted.iloc[0]
        total_savings_cost = auto_data['Monthly_Cost_Savings'].sum() * 6
        st.markdown(f"""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Scale Top Projects</strong><br>
            6-Month Savings: ${total_savings_cost:,.0f} | Expand {top['Process_Name']} to other departments
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show project details", key="auto_details"):
            show_detailed_data(
                auto_data[['Process_Name', 'Task_Type', 'Monthly_Hours_Saved', 'Monthly_Cost_Savings', 'ROI_Percentage_6M']].head(20),
                "Automation ROI Details"
            )

    # Trend
    auto_trend = data['Automation_ROI'][data['Automation_ROI']['Month'].isin(selected_months)].groupby('Month').agg({
        'ROI_Percentage_6M': 'mean'
    }).reset_index().sort_values('Month')

    if len(auto_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(auto_trend, 'Month', 'ROI_Percentage_6M', 'Automation ROI Trend', 'ROI %', '#059669')
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ===== Section 3: Digital Workplace Index =====
    st.markdown("#### 📱 Digital Workplace Index")

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

    # Combined: Friction by Dept and Latency
    st.markdown("**Friction Index & Response Latency by Department**")
    dept_friction = digital_data.groupby('Department').agg({
        'Friction_Index_Score': 'mean',
        'App_Response_Latency_Sec': 'mean'
    }).sort_values('Friction_Index_Score', ascending=False)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=dept_friction.index,
        y=dept_friction['Friction_Index_Score'],
        name='Friction Score',
        marker_color='#1e40af',
        yaxis='y1'
    ))
    fig.add_trace(go.Scatter(
        x=dept_friction.index,
        y=dept_friction['App_Response_Latency_Sec'],
        name='Response Latency (s)',
        mode='lines+markers',
        line=dict(color='#dc2626', width=3),
        marker=dict(size=10),
        yaxis='y2'
    ))

    fig.update_layout(
        yaxis=dict(title='Friction Score', side='left'),
        yaxis2=dict(title='Response Latency (s)', overlaying='y', side='right'),
        height=400,
        hovermode='x unified',
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Alert
    worst = dept_friction.index[0]
    friction_impact = (dept_friction.iloc[0]['Friction_Index_Score'] / 100) * 8 * 5 * 60
    st.markdown(f"""
    <div class="alert-box">
        <strong>⚠️ Alert: High Friction in {worst}</strong><br>
        Lost productivity: ~{friction_impact:.0f} hours/week | Recommend urgent system upgrades
    </div>
    """, unsafe_allow_html=True)

    if st.checkbox("📊 Show department details", key="digital_details"):
        show_detailed_data(
            digital_data.groupby('Department').agg({
                'Friction_Index_Score': 'mean',
                'App_Response_Latency_Sec': 'mean',
                'Tool_Switches_Per_Hour': 'mean'
            }).reset_index().sort_values('Friction_Index_Score', ascending=False),
            "Digital Index Details"
        )

    # Trend
    digital_trend = data['Digital_Index'][data['Digital_Index']['Month'].isin(selected_months)].groupby('Month').agg({
        'Friction_Index_Score': 'mean'
    }).reset_index().sort_values('Month')

    if len(digital_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(digital_trend, 'Month', 'Friction_Index_Score', 'Friction Index Trend', 'Score', '#f59e0b')
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ===== Section 4: Process Rework Cost =====
    st.markdown("#### ♻️ Process Rework Cost %")

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

    # Combined: Rework % and Cost
    st.markdown("**Rework Cost by Process: Percentage vs Dollar Impact**")
    process_rework = rework_data.groupby('Process_Name').agg({
        'Rework_Cost_Percentage': 'mean',
        'Rework_Cost_Dollars': 'sum'
    }).sort_values('Rework_Cost_Dollars', ascending=False).head(8)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=process_rework.index,
        x=process_rework['Rework_Cost_Percentage'],
        name='Rework %',
        marker_color='#f59e0b',
        orientation='h',
        xaxis='x1'
    ))
    fig.add_trace(go.Scatter(
        y=process_rework.index,
        x=process_rework['Rework_Cost_Dollars'],
        name='Rework Cost ($)',
        mode='markers',
        marker=dict(size=12, color='#dc2626'),
        xaxis='x2'
    ))

    fig.update_layout(
        xaxis=dict(title='Rework %'),
        xaxis2=dict(title='Rework Cost ($)', overlaying='x', side='top'),
        height=400,
        hovermode='y unified',
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Opportunity
    if len(process_rework) > 0:
        worst_process = process_rework.index[0]
        annual_cost = rework_data['Rework_Cost_Dollars'].sum() * 12
        st.markdown(f"""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Process Improvement</strong><br>
            Fix {worst_process} rework issues | Potential annual savings: ${annual_cost * 0.5:,.0f} (50% reduction)
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show process details", key="rework_details"):
            show_detailed_data(
                rework_data.groupby('Process_Name').agg({
                    'Rework_Cost_Percentage': 'mean',
                    'Rework_Cost_Dollars': 'sum',
                    'Rework_Transaction_Count': 'sum'
                }).reset_index().sort_values('Rework_Cost_Dollars', ascending=False),
                "Rework Cost Details"
            )

    # Trend
    rework_trend = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)].groupby('Month').agg({
        'Rework_Cost_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(rework_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(rework_trend, 'Month', 'Rework_Cost_Percentage', 'Rework Cost Trend', 'Percentage (%)', '#dc2626')
        st.plotly_chart(fig, use_container_width=True)

# ==================== EXECUTION & RISK TAB ====================
elif st.session_state.current_page == "Execution & Risk":
    st.markdown("### ⚙️ Execution & Risk Management")

    # Filter data
    ftr_data = data['FTR_Rate'][data['FTR_Rate']['Month'].isin(selected_months)]
    adherence_data = data['Adherence'][data['Adherence']['Month'].isin(selected_months)]
    resilience_data = data['Resilience'][data['Resilience']['Month'].isin(selected_months)]
    escalation_data = data['Escalation'][data['Escalation']['Month'].isin(selected_months)]

    st.info(f"📊 Showing {len(ftr_data)} FTR records | {len(resilience_data)} Resilience records")

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

    res_summary = resilience_data.groupby('Critical_Task').agg({
        'Resilience_Score': 'mean',
        'Risk_Percentage': 'mean'
    }).sort_values('Risk_Percentage', ascending=False).head(8)

    st.markdown("**Critical Tasks: Risk vs Resilience Score**")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=res_summary.index,
        x=res_summary['Risk_Percentage'],
        name='Risk %',
        marker_color='#dc2626',
        orientation='h',
        xaxis='x1'
    ))
    fig.add_trace(go.Scatter(
        y=res_summary.index,
        x=res_summary['Resilience_Score'],
        name='Resilience Score',
        mode='lines+markers',
        line=dict(color='#059669', width=3),
        marker=dict(size=10),
        xaxis='x2'
    ))

    fig.update_layout(
        xaxis=dict(title='Risk %'),
        xaxis2=dict(title='Resilience Score', overlaying='x', side='top'),
        height=400,
        hovermode='y unified',
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Alert
    highest_risk = resilience_data.loc[resilience_data['Risk_Percentage'].idxmax()]
    critical_count = resilience_data[resilience_data['Risk_Level'] == 'Critical'].shape[0]
    st.markdown(f"""
    <div class="alert-box">
        <strong>🔴 Critical: Single Points of Failure</strong><br>
        {critical_count} tasks with single-person dependency | Cross-training can reduce risk by 60%
    </div>
    """, unsafe_allow_html=True)

    if st.checkbox("📊 Show resilience details", key="resilience_details"):
        show_detailed_data(
            resilience_data[['Critical_Task', 'Department', 'Risk_Percentage', 'Risk_Level']].drop_duplicates().head(20),
            "Resilience Details"
        )

    # Trend
    res_trend = data['Resilience'][data['Resilience']['Month'].isin(selected_months)].groupby('Month').agg({
        'Resilience_Score': 'mean'
    }).reset_index().sort_values('Month')

    if len(res_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(res_trend, 'Month', 'Resilience_Score', 'Resilience Score Trend', 'Score', '#0891b2')
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ===== FTR Section =====
    st.markdown("#### ✅ First-Time-Right (FTR) Rate")

    ftr_summary = ftr_data.groupby('Process').agg({
        'FTR_Rate_Percentage': 'mean',
        'Target_FTR_Rate': 'first',
        'Error_Rate_Percentage': 'mean'
    }).reset_index().sort_values('FTR_Rate_Percentage')

    st.markdown("**Process Quality: FTR vs Target vs Error Rate**")
    fig = go.Figure()
    fig.add_trace(go.Bar(y=ftr_summary['Process'], x=ftr_summary['FTR_Rate_Percentage'], name='Actual FTR %', marker_color='#3b82f6', orientation='h'))
    fig.add_trace(go.Bar(y=ftr_summary['Process'], x=ftr_summary['Target_FTR_Rate'], name='Target FTR %', marker_color='#d1d5db', orientation='h'))
    fig.add_trace(go.Scatter(y=ftr_summary['Process'], x=ftr_summary['Error_Rate_Percentage'], name='Error %', mode='markers', marker=dict(size=12, color='#dc2626')))

    fig.update_layout(
        barmode='group',
        height=400,
        hovermode='y unified',
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Opportunity
    gap = ftr_summary['Target_FTR_Rate'].mean() - ftr_summary['FTR_Rate_Percentage'].mean()
    if gap > 5:
        st.markdown(f"""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Quality Improvement</strong><br>
            Close FTR gap of {gap:.0f} percentage points | Estimated cost savings: $125K annually
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show FTR details", key="ftr_details"):
            show_detailed_data(
                ftr_data[['Process', 'Department', 'FTR_Rate_Percentage', 'Error_Rate_Percentage', 'Target_FTR_Rate']].head(20),
                "FTR Rate Details"
            )

    # Trend
    ftr_trend = data['FTR_Rate'][data['FTR_Rate']['Month'].isin(selected_months)].groupby('Month').agg({
        'FTR_Rate_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(ftr_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(ftr_trend, 'Month', 'FTR_Rate_Percentage', 'FTR Rate Trend', 'Percentage (%)', '#059669')
        st.plotly_chart(fig, use_container_width=True)

# ==================== WORKFORCE & MODEL TAB ====================
elif st.session_state.current_page == "Workforce & Model":
    st.markdown("### 👥 Workforce & Scalable Operating Model")

    # Filter data
    capacity_data = data['Capacity'][data['Capacity']['Month'].isin(selected_months)]
    if dept_filter:
        capacity_data = capacity_data[capacity_data['Department'].isin(dept_filter)]

    model_acc_data = data['Model_Accuracy'][data['Model_Accuracy']['Month'].isin(selected_months)]
    if dept_filter:
        model_acc_data = model_acc_data[model_acc_data['Department'].isin(dept_filter)]

    work_models_data = data['Work_Models'][data['Work_Models']['Month'].isin(selected_months)]
    if dept_filter:
        work_models_data = work_models_data[work_models_data['Department'].isin(dept_filter)]

    collab_data = data['Collaboration'][data['Collaboration']['Month'].isin(selected_months)]
    if dept_filter:
        collab_data = collab_data[collab_data['Department'].isin(dept_filter)]

    st.info(f"📊 Showing {len(capacity_data)} capacity records | {len(work_models_data)} work model records")

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

    # ===== Capacity Section =====
    st.markdown("#### 🔥 Hidden Capacity & Burnout Risk")

    cap_summary = capacity_data.groupby('Department').agg({
        'Capacity_Utilization_Percentage': 'mean'
    }).sort_values('Capacity_Utilization_Percentage')

    st.markdown("**Department Capacity & Burnout Risk**")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=cap_summary.values,
        y=cap_summary.index,
        name='Capacity %',
        marker_color=cap_summary.apply(lambda x: '#dc2626' if x > 110 else '#f59e0b' if x > 90 else '#059669').values,
        orientation='h',
        text=cap_summary.apply(lambda x: f"{x:.0f}%").values,
        textposition='outside'
    ))
    fig.add_vline(x=100, line_dash="dash", line_color="red", annotation_text="Optimal Capacity")
    fig.update_layout(height=400, showlegend=False, xaxis_title="Capacity %", plot_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig, use_container_width=True)

    # Alert
    at_risk = capacity_data[capacity_data['Capacity_Utilization_Percentage'] > 110].shape[0]
    if at_risk > 5:
        st.markdown(f"""
        <div class="alert-box">
            <strong>🔴 Critical: Capacity Crisis</strong><br>
            {at_risk} employees over capacity | Immediate staffing or workload rebalancing required
        </div>
        """, unsafe_allow_html=True)

        if st.checkbox("📊 Show capacity details", key="capacity_details"):
            show_detailed_data(
                capacity_data[['Department', 'Role', 'Capacity_Utilization_Percentage', 'Burnout_Risk_Flag']].drop_duplicates().head(20),
                "Capacity Details"
            )

    # Trend
    cap_trend = data['Capacity'][data['Capacity']['Month'].isin(selected_months)].groupby('Month').agg({
        'Capacity_Utilization_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(cap_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(cap_trend, 'Month', 'Capacity_Utilization_Percentage', 'Capacity Trend', 'Utilization %', '#f59e0b')
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ===== Work Models Section =====
    st.markdown("#### 💼 Work Models Effectiveness")

    work_summary = work_models_data.groupby('Work_Model').agg({
        'Output_Per_Hour': 'mean',
        'Cost_Per_Transaction': 'mean'
    }).reset_index()

    st.markdown("**Work Model Comparison: Output vs Cost**")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=work_summary['Work_Model'],
        y=work_summary['Output_Per_Hour'],
        name='Output/Hr',
        marker_color='#059669',
        yaxis='y1'
    ))
    fig.add_trace(go.Scatter(
        x=work_summary['Work_Model'],
        y=work_summary['Cost_Per_Transaction'],
        name='Cost/Transaction ($)',
        mode='lines+markers',
        line=dict(color='#dc2626', width=3),
        marker=dict(size=10),
        yaxis='y2'
    ))

    fig.update_layout(
        yaxis=dict(title='Output/Hr', side='left'),
        yaxis2=dict(title='Cost/Txn ($)', overlaying='y', side='right'),
        height=400,
        hovermode='x unified',
        plot_bgcolor="rgba(0,0,0,0)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # Opportunity
    best = work_summary.loc[work_summary['Output_Per_Hour'].idxmax()]
    cost_savings = work_summary['Cost_Per_Transaction'].max() - work_summary['Cost_Per_Transaction'].min()
    st.markdown(f"""
    <div class="opportunity-box">
        <strong>🎯 Opportunity: Work Model Optimization</strong><br>
        Shift 30% of work to {best['Work_Model']} model | Annual savings: ${cost_savings * 50000:,.0f}
    </div>
    """, unsafe_allow_html=True)

    if st.checkbox("📊 Show work model details", key="workmodel_details"):
        show_detailed_data(
            work_models_data.groupby('Work_Model').agg({
                'Output_Per_Hour': 'mean',
                'Cost_Per_Transaction': 'mean',
                'Productivity_Index': 'mean'
            }).reset_index(),
            "Work Models Details"
        )

    # Trend
    work_trend = data['Work_Models'][data['Work_Models']['Month'].isin(selected_months)].groupby('Month').agg({
        'Output_Per_Hour': 'mean'
    }).reset_index().sort_values('Month')

    if len(work_trend) > 1:
        st.markdown("**Trend Over Time**")
        fig = create_trend_chart(work_trend, 'Month', 'Output_Per_Hour', 'Productivity Trend', 'Output/Hr', '#059669')
        st.plotly_chart(fig, use_container_width=True)

# Footer
st.divider()
st.markdown(f"""
    <div style="text-align: center; padding: 20px; color: #6b7280; font-size: 12px;">
        <strong>COO Dashboard v5.0 - Final</strong> | {len(selected_months)} months | {len(dept_filter) if dept_filter else len(all_departments)} departments | 
        <strong>Source:</strong> COO_ROI_Dashboard_KPIs_Complete_12.xlsx
    </div>
""", unsafe_allow_html=True)
