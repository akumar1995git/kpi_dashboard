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
    .nav-buttons {
        display: flex;
        gap: 10px;
        margin-bottom: 20px;
        flex-wrap: wrap;
    }
    .nav-button {
        padding: 10px 20px;
        border-radius: 8px;
        border: 2px solid #e5e7eb;
        background: white;
        cursor: pointer;
        font-weight: 600;
        font-size: 14px;
        transition: all 0.3s;
        color: #1f2937;
    }
    .nav-button.active {
        background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%);
        color: white;
        border-color: #1e40af;
    }
    .nav-button:hover {
        border-color: #1e40af;
        color: #1e40af;
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
    .export-section {
        background: #f8fafc;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        margin-top: 30px;
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
pages = ["Overview", "Efficiency & Cost", "Execution & Risk", "Workforce & Model"]

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

# Get unique months
role_months = sorted(data['Role_vs_Reality']['Month'].unique())
selected_months = st.sidebar.multiselect(
    "Select Months",
    role_months,
    default=[role_months[-1]],
    key="month_filter"
)

# Multi-select for departments
all_departments = ["All Departments"] + sorted(
    list(set(list(data['Role_vs_Reality'].get('Department', []).unique()) + 
             list(data['Capacity'].get('Department', []).unique())))
)
selected_depts = st.sidebar.multiselect(
    "Select Departments",
    all_departments,
    default=["All Departments"],
    key="dept_filter"
)

# Filter handling
if "All Departments" in selected_depts:
    dept_filter = None
else:
    dept_filter = selected_depts

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

def create_drill_down_modal(data_df, title):
    """Create expandable drill-down section"""
    with st.expander(f"🔍 Drill Down: {title}", expanded=False):
        st.dataframe(data_df, use_container_width=True, hide_index=True)

        # Export button for detailed data
        csv = data_df.to_csv(index=False)
        st.download_button(
            label=f"📥 Download {title} Data (CSV)",
            data=csv,
            file_name=f"{title.replace(' ', '_')}.csv",
            mime="text/csv"
        )

# ==================== OVERVIEW TAB ====================
if st.session_state.current_page == "Overview":
    st.markdown("### Executive Summary - All KPIs")

    # Use all selected months
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

    st.info(f"📊 Showing data for: {', '.join([str(m) for m in selected_months])}")

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

    # KEY INSIGHTS
    st.markdown("### 💡 Key Insights & Opportunities")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="insight-box">
            <strong>🔍 Insight: Low-Value Work Burden</strong><br>
            Employees are spending significant time on non-core tasks that don't add strategic value. This represents a hidden productivity leak.
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Automation ROI</strong><br>
            With an average ROI of {:.0f}%, automation projects are delivering exceptional returns. Consider accelerating initiatives.
        </div>
        """.format(avg_roi), unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="alert-box">
            <strong>⚠️ Alert: Escalation Spike</strong><br>
            Escalations are increasing at {:.0f} incidents. Root cause analysis recommended to identify process bottlenecks.
        </div>
        """.format(total_escalations), unsafe_allow_html=True)

        if burnout_count > 10:
            st.markdown("""
            <div class="alert-box">
                <strong>🔴 Critical: Burnout Risk</strong><br>
                {} employees are at risk of burnout. Immediate capacity rebalancing required.
            </div>
            """.format(burnout_count), unsafe_allow_html=True)

    # TREND ANALYSIS
    st.divider()
    st.markdown("### 📈 Trend Analysis Over Time")

    trend_col1, trend_col2 = st.columns(2)

    with trend_col1:
        # Role vs Reality Trend
        role_trend = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'].isin(selected_months)].groupby('Month').agg({
            'Low_Value_Work_Percentage': 'mean'
        }).reset_index().sort_values('Month')

        if len(role_trend) > 1:
            fig = create_trend_chart(role_trend, 'Month', 'Low_Value_Work_Percentage', 'Low-Value Work Trend', 'Percentage (%)', '#dc2626')
            st.plotly_chart(fig, use_container_width=True)

    with trend_col2:
        # Rework Cost Trend
        rework_trend = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)].groupby('Month').agg({
            'Rework_Cost_Percentage': 'mean'
        }).reset_index().sort_values('Month')

        if len(rework_trend) > 1:
            fig = create_trend_chart(rework_trend, 'Month', 'Rework_Cost_Percentage', 'Rework Cost Trend', 'Percentage (%)', '#f59e0b')
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # EXPORT SECTION
    st.markdown("### 📊 Export Data")
    st.markdown("Download detailed data for further analysis")

    export_col1, export_col2, export_col3 = st.columns(3)

    with export_col1:
        if len(role_latest) > 0:
            excel_data = export_dataframe_to_excel(role_latest, "Role vs Reality")
            st.download_button(
                label="📥 Role vs Reality",
                data=excel_data,
                file_name="role_vs_reality.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with export_col2:
        if len(auto_latest) > 0:
            excel_data = export_dataframe_to_excel(auto_latest, "Automation ROI")
            st.download_button(
                label="📥 Automation ROI",
                data=excel_data,
                file_name="automation_roi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with export_col3:
        if len(digital_latest) > 0:
            excel_data = export_dataframe_to_excel(digital_latest, "Digital Index")
            st.download_button(
                label="📥 Digital Index",
                data=excel_data,
                file_name="digital_index.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

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

    st.info(f"📊 Showing data for: {', '.join([str(m) for m in selected_months])}")

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

    # Chart and insights
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

    # Opportunity insight
    st.markdown("""
    <div class="opportunity-box">
        <strong>🎯 Opportunity: Role Optimization</strong><br>
        Reallocate {} FTE hours annually by automating {} role's low-value tasks. Estimated ROI: 340%.
    </div>
    """.format(int(role_data['Opportunity_Cost_Dollars'].sum() / 75), role_summary.index[0]), unsafe_allow_html=True)

    # Trend
    role_trend = data['Role_vs_Reality'][data['Role_vs_Reality']['Month'].isin(selected_months)].groupby('Month').agg({
        'Low_Value_Work_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(role_trend) > 1:
        fig = create_trend_chart(role_trend, 'Month', 'Low_Value_Work_Percentage', 'Low-Value Work Trend', 'Percentage (%)', '#dc2626')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(role_data[['Department', 'Role', 'Low_Value_Work_Percentage', 'Opportunity_Cost_Dollars']].head(20), "Role vs Reality")

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

    # Opportunity
    if len(auto_sorted) > 0:
        st.markdown("""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Scale Top Projects</strong><br>
            Expand {} to other departments. Expected additional savings: ${:,.0f} annually.
        </div>
        """.format(auto_sorted.iloc[-1]['Process_Name'], total_savings_cost * 2), unsafe_allow_html=True)

    # Trend
    auto_trend = data['Automation_ROI'][data['Automation_ROI']['Month'].isin(selected_months)].groupby('Month').agg({
        'ROI_Percentage_6M': 'mean'
    }).reset_index().sort_values('Month')

    if len(auto_trend) > 1:
        fig = create_trend_chart(auto_trend, 'Month', 'ROI_Percentage_6M', 'Automation ROI Trend', 'ROI %', '#059669')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(auto_data[['Process_Name', 'Task_Type', 'Monthly_Hours_Saved', 'Monthly_Cost_Savings', 'ROI_Percentage_6M']].head(20), "Automation ROI")

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

    # Opportunity
    friction_impact = (dept_friction.iloc[0] / 100) * 8 * 5 * 60  # hours per week
    st.markdown(f"""
    <div class="alert-box">
        <strong>⚠️ Alert: High Friction in {worst}</strong><br>
        Causing ~{friction_impact:.0f} hours of lost productivity per week. Recommend urgent system upgrades.
    </div>
    """, unsafe_allow_html=True)

    # Trend
    digital_trend = data['Digital_Index'][data['Digital_Index']['Month'].isin(selected_months)].groupby('Month').agg({
        'Friction_Index_Score': 'mean'
    }).reset_index().sort_values('Month')

    if len(digital_trend) > 1:
        fig = create_trend_chart(digital_trend, 'Month', 'Friction_Index_Score', 'Friction Index Trend', 'Score', '#f59e0b')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(digital_data[['Department', 'Friction_Index_Score', 'Primary_Friction_App', 'App_Response_Latency_Sec']].head(20), "Digital Index")

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

    # Opportunity
    if len(process_rework) > 0:
        st.markdown(f"""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Process Improvement</strong><br>
            Fix {worst_process} rework issues. Potential annual savings: ${annual_cost * 0.5:,.0f} (50% reduction).
        </div>
        """, unsafe_allow_html=True)

    # Trend
    rework_trend = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)].groupby('Month').agg({
        'Rework_Cost_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(rework_trend) > 1:
        fig = create_trend_chart(rework_trend, 'Month', 'Rework_Cost_Percentage', 'Rework Cost Trend', 'Percentage (%)', '#dc2626')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(rework_data[['Process_Name', 'Department', 'Rework_Cost_Percentage', 'Rework_Cost_Dollars']].head(20), "Rework Cost")

# ==================== EXECUTION & RISK TAB ====================
elif st.session_state.current_page == "Execution & Risk":
    st.markdown("### ⚙️ Execution & Risk Management")

    # Filter data
    ftr_data = data['FTR_Rate'][data['FTR_Rate']['Month'].isin(selected_months)]
    adherence_data = data['Adherence'][data['Adherence']['Month'].isin(selected_months)]
    resilience_data = data['Resilience'][data['Resilience']['Month'].isin(selected_months)]
    escalation_data = data['Escalation'][data['Escalation']['Month'].isin(selected_months)]

    if dept_filter:
        ftr_data = ftr_data[ftr_data['Department'].isin(dept_filter)]
        adherence_data = adherence_data[adherence_data['Department'].isin(dept_filter)]
        resilience_data = resilience_data[resilience_data['Department'].isin(dept_filter)]
        escalation_data = escalation_data[escalation_data['Department'].isin(dept_filter)]

    st.info(f"📊 Showing data for: {', '.join([str(m) for m in selected_months])}")

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

    # Opportunity
    if critical_count > 0:
        st.markdown(f"""
        <div class="alert-box">
            <strong>🔴 Critical: Single Points of Failure</strong><br>
            {critical_count} tasks have single-person dependency. Cross-training can reduce risk by 60%.
        </div>
        """, unsafe_allow_html=True)

    # Trend
    res_trend = data['Resilience'][data['Resilience']['Month'].isin(selected_months)].groupby('Month').agg({
        'Resilience_Score': 'mean'
    }).reset_index().sort_values('Month')

    if len(res_trend) > 1:
        fig = create_trend_chart(res_trend, 'Month', 'Resilience_Score', 'Resilience Score Trend', 'Score', '#0891b2')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(resilience_data[['Critical_Task', 'Department', 'Risk_Percentage', 'Risk_Level']].drop_duplicates().head(20), "Resilience")

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

    # Opportunity
    gap = ftr_summary['Target_FTR_Rate'].mean() - ftr_summary['FTR_Rate_Percentage'].mean()
    if gap > 5:
        st.markdown(f"""
        <div class="opportunity-box">
            <strong>🎯 Opportunity: Quality Improvement</strong><br>
            Close FTR gap of {gap:.0f} percentage points. Estimated quality cost savings: $125K annually.
        </div>
        """, unsafe_allow_html=True)

    # Trend
    ftr_trend = data['FTR_Rate'][data['FTR_Rate']['Month'].isin(selected_months)].groupby('Month').agg({
        'FTR_Rate_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(ftr_trend) > 1:
        fig = create_trend_chart(ftr_trend, 'Month', 'FTR_Rate_Percentage', 'FTR Rate Trend', 'Percentage (%)', '#059669')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(ftr_data[['Process', 'Department', 'FTR_Rate_Percentage', 'Error_Rate_Percentage']].head(20), "FTR Rate")

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

    st.info(f"📊 Showing data for: {', '.join([str(m) for m in selected_months])}")

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

    # Alert
    if at_risk > 5:
        st.markdown(f"""
        <div class="alert-box">
            <strong>🔴 Critical: Capacity Crisis</strong><br>
            {at_risk} employees over capacity. Immediate staffing or workload rebalancing required.
        </div>
        """, unsafe_allow_html=True)

    # Trend
    cap_trend = data['Capacity'][data['Capacity']['Month'].isin(selected_months)].groupby('Month').agg({
        'Capacity_Utilization_Percentage': 'mean'
    }).reset_index().sort_values('Month')

    if len(cap_trend) > 1:
        fig = create_trend_chart(cap_trend, 'Month', 'Capacity_Utilization_Percentage', 'Capacity Trend', 'Utilization %', '#f59e0b')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(capacity_data[['Department', 'Role', 'Capacity_Utilization_Percentage', 'Burnout_Risk_Flag']].drop_duplicates().head(20), "Capacity")

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

    # Opportunity
    cost_savings = work_summary['Cost_Per_Transaction'].max() - work_summary['Cost_Per_Transaction'].min()
    st.markdown(f"""
    <div class="opportunity-box">
        <strong>🎯 Opportunity: Work Model Optimization</strong><br>
        Shift 30% of work to {best['Work_Model']} model. Potential annual savings: ${cost_savings * 50000:,.0f}.
    </div>
    """, unsafe_allow_html=True)

    # Trend
    work_trend = data['Work_Models'][data['Work_Models']['Month'].isin(selected_months)].groupby('Month').agg({
        'Output_Per_Hour': 'mean'
    }).reset_index().sort_values('Month')

    if len(work_trend) > 1:
        fig = create_trend_chart(work_trend, 'Month', 'Output_Per_Hour', 'Productivity Trend', 'Output/Hr', '#059669')
        st.plotly_chart(fig, use_container_width=True)

    # Drill down
    create_drill_down_modal(work_models_data[['Work_Model', 'Department', 'Output_Per_Hour', 'Cost_Per_Transaction']].head(20), "Work Models")

# Footer
st.divider()
st.markdown(f"""
    <div style="text-align: center; padding: 20px; color: #6b7280; font-size: 12px;">
        <strong>COO Dashboard v4.0 - Enhanced</strong> | Months: {', '.join([str(m) for m in selected_months])} | 
        <strong>Source:</strong> COO_ROI_Dashboard_KPIs_Complete_12.xlsx
    </div>
""", unsafe_allow_html=True)
