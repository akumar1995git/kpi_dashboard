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

# ==================== CUSTOM CSS ====================
st.markdown("""
    <style>
    .main {
        padding: 0px;
    }

    /* L1 Top KPI Cards */
    .l1-kpi-card {
        background: linear-gradient(135deg, #1e40af 0%, #1e3a8a 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
    }

    .l1-value {
        font-size: 36px;
        font-weight: 800;
        margin: 10px 0;
    }

    .l1-label {
        font-size: 13px;
        opacity: 0.9;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .l1-subtitle {
        font-size: 11px;
        opacity: 0.8;
        margin-top: 8px;
    }

    /* Objective Cards */
    .objective-card {
        background: white;
        border-radius: 10px;
        border: 2px solid #e5e7eb;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }

    .objective-card:hover {
        border-color: #1e40af;
        box-shadow: 0 4px 8px rgba(30, 64, 175, 0.15);
    }

    .objective-title {
        font-size: 16px;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .objective-signal {
        font-size: 12px;
        color: #6b7280;
        font-style: italic;
        margin-bottom: 15px;
    }

    /* L2 Metric Boxes */
    .l2-metric {
        background: #f9fafb;
        border-left: 4px solid #1e40af;
        padding: 15px;
        border-radius: 6px;
        margin-bottom: 12px;
    }

    .l2-metric.cost {
        border-left-color: #ef4444;
    }

    .l2-metric.quality {
        border-left-color: #059669;
    }

    .l2-metric.efficiency {
        border-left-color: #f59e0b;
    }

    .l2-metric-title {
        font-size: 13px;
        font-weight: 600;
        color: #1f2937;
        margin-bottom: 8px;
    }

    .l2-metric-value {
        font-size: 24px;
        font-weight: 700;
        color: #1e40af;
    }

    .metric-trend {
        font-size: 11px;
        color: #6b7280;
        margin-top: 6px;
    }

    /* Export Section */
    .export-section {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 20px;
        border-radius: 10px;
        margin-top: 30px;
    }

    .export-header {
        font-size: 18px;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 14px;
        font-weight: 600;
    }

    /* Divider */
    .divider-custom {
        border: none;
        border-top: 2px solid #e5e7eb;
        margin: 30px 0;
    }
    </style>
""", unsafe_allow_html=True)

# ==================== LOAD DATA ====================
@st.cache_data
def load_excel_data():
    """Load all data from Excel file"""
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
        st.error("❌ Excel file not found: 'COO_ROI_Dashboard_KPIs_Complete_12.xlsx'")
        st.stop()

data = load_excel_data()

# Initialize session state
if 'current_objective' not in st.session_state:
    st.session_state.current_objective = "Cost & Efficiency"

# ==================== HEADER ====================
st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); 
                color: white; padding: 30px 20px; border-radius: 0; 
                margin: -25px -20px 25px -20px; text-align: center;">
        <h1 style="margin: 0; font-size: 32px;">📊 COO Operational Dashboard</h1>
        <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">Real-time KPI Monitoring & Insights</p>
    </div>
""", unsafe_allow_html=True)

# ==================== SIDEBAR FILTERS ====================
st.sidebar.markdown("## 🔧 Filters")

role_months = sorted(data['Role_vs_Reality']['Month'].unique())
selected_months = st.sidebar.multiselect(
    "Select Months",
    role_months,
    default=list(role_months),
    key="month_filter"
)

all_departments = sorted(
    list(set(list(data['Role_vs_Reality'].get('Department', []).unique()) + 
             list(data['Capacity'].get('Department', []).unique())))
)
selected_depts = st.sidebar.multiselect(
    "Select Departments",
    all_departments,
    default=all_departments,
    key="dept_filter"
)

dept_filter = selected_depts if len(selected_depts) > 0 else None

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ==================== DATA FILTERING HELPER ====================
def filter_data(df, month_col='Month', dept_col='Department'):
    """Filter dataframe by selected months and departments"""
    result = df[df[month_col].isin(selected_months)]
    if dept_filter and dept_col in result.columns:
        result = result[result[dept_col].isin(dept_filter)]
    return result

# ==================== L1: TOP KPIs ====================
st.markdown("### 📈 L1: Executive KPIs")

# Calculate metrics
role_data = filter_data(data['Role_vs_Reality'])
auto_data = filter_data(data['Automation_ROI'])
ftr_data = filter_data(data['FTR_Rate'])
capacity_data = filter_data(data['Capacity'])
work_data = filter_data(data['Work_Models'])

rework_pct = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)]['Rework_Cost_Percentage'].mean() if len(selected_months) > 0 else 0
auto_roi = auto_data['ROI_Percentage_6M'].mean() if len(auto_data) > 0 else 0
ftr_rate = ftr_data['FTR_Rate_Percentage'].mean() if len(ftr_data) > 0 else 0
avg_capacity = capacity_data['Capacity_Utilization_Percentage'].mean() if len(capacity_data) > 0 else 0
avg_output = work_data['Output_Per_Hour'].mean() if len(work_data) > 0 else 0

# L1 KPI Row
col1, col2, col3 = st.columns(3, gap="medium")

with col1:
    st.markdown(f"""
    <div class="l1-kpi-card">
        <div class="l1-label">💰 Rework Cost %</div>
        <div class="l1-value">{rework_pct:.1f}%</div>
        <div class="l1-subtitle">Cost & Efficiency Signal</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="l1-kpi-card">
        <div class="l1-label">✅ FTR Rate</div>
        <div class="l1-value">{ftr_rate:.0f}%</div>
        <div class="l1-subtitle">Quality Signal</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="l1-kpi-card">
        <div class="l1-label">⚡ Output/FTE</div>
        <div class="l1-value">{avg_output:.2f}</div>
        <div class="l1-subtitle">Productivity Signal</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==================== L2: OBJECTIVE CARDS ====================
st.markdown("### 🎯 L2: Objective Deep-Dives")

col1, col2, col3 = st.columns(3, gap="medium")

# -------- OBJECTIVE 1: COST & EFFICIENCY --------
with col1:
    st.markdown("""
    <div class="objective-card">
        <div class="objective-title">💡 Cost & Efficiency</div>
        <div class="objective-signal">Financial signal: Control costs & maximize output</div>
    """, unsafe_allow_html=True)

    # L2 Metrics
    rework_cost = data['Process_Rework'][data['Process_Rework']['Month'].isin(selected_months)]['Rework_Cost_Dollars'].sum()

    st.markdown(f"""
    <div class="l2-metric cost">
        <div class="l2-metric-title">💸 Rework Cost</div>
        <div class="l2-metric-value">${rework_cost:,.0f}</div>
        <div class="metric-trend">Monthly cost impact</div>
    </div>
    """, unsafe_allow_html=True)

    auto_value = auto_data['Monthly_Cost_Savings'].sum() if len(auto_data) > 0 else 0
    st.markdown(f"""
    <div class="l2-metric efficiency">
        <div class="l2-metric-title">🤖 Automation ROI</div>
        <div class="l2-metric-value">{auto_roi:.0f}%</div>
        <div class="metric-trend">{auto_value:,.0f} monthly savings</div>
    </div>
    """, unsafe_allow_html=True)

    digital_index = data['Digital_Index'][data['Digital_Index']['Month'].isin(selected_months)]['Friction_Index_Score'].mean()
    st.markdown(f"""
    <div class="l2-metric efficiency">
        <div class="l2-metric-title">📱 Digital Friction Index</div>
        <div class="l2-metric-value">{digital_index:.1f}</div>
        <div class="metric-trend">Lower is better</div>
    </div>
    """, unsafe_allow_html=True)

    low_value = role_data['Low_Value_Work_Percentage'].mean() if len(role_data) > 0 else 0
    st.markdown(f"""
    <div class="l2-metric cost">
        <div class="l2-metric-title">📊 Low-Value Work %</div>
        <div class="l2-metric-value">{low_value:.1f}%</div>
        <div class="metric-trend">Role optimization opportunity</div>
    </div>
    """, unsafe_allow_html=True)

    # Expandable details
    if st.checkbox("📋 View Cost & Efficiency Details", key="obj1_details"):
        st.markdown("#### Key Metrics by Department")

        dept_summary = role_data.groupby('Department').agg({
            'Low_Value_Work_Percentage': 'mean',
            'Opportunity_Cost_Dollars': 'sum'
        }).sort_values('Opportunity_Cost_Dollars', ascending=False)
        st.dataframe(dept_summary, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# -------- OBJECTIVE 2: EXECUTION & RESILIENCE --------
with col2:
    st.markdown("""
    <div class="objective-card">
        <div class="objective-title">⚙️ Execution & Resilience</div>
        <div class="objective-signal">Quality signal: Minimize errors & build resilience</div>
    """, unsafe_allow_html=True)

    resilience_score = data['Resilience'][data['Resilience']['Month'].isin(selected_months)]['Resilience_Score'].mean()
    st.markdown(f"""
    <div class="l2-metric quality">
        <div class="l2-metric-title">🛡️ Resilience Score</div>
        <div class="l2-metric-value">{resilience_score:.1f}/10</div>
        <div class="metric-trend">Process robustness</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="l2-metric quality">
        <div class="l2-metric-title">✅ FTR Rate</div>
        <div class="l2-metric-value">{ftr_rate:.1f}%</div>
        <div class="metric-trend">First-time accuracy</div>
    </div>
    """, unsafe_allow_html=True)

    adherence_rate = data['Adherence'][data['Adherence']['Month'].isin(selected_months)]['Adherence_Rate_Percentage'].mean()
    st.markdown(f"""
    <div class="l2-metric quality">
        <div class="l2-metric-title">📋 Process Adherence</div>
        <div class="l2-metric-value">{adherence_rate:.1f}%</div>
        <div class="metric-trend">Policy compliance</div>
    </div>
    """, unsafe_allow_html=True)

    escalations = data['Escalation'][data['Escalation']['Month'].isin(selected_months)]['Step_Exception_Count'].sum()
    st.markdown(f"""
    <div class="l2-metric cost">
        <div class="l2-metric-title">🚨 Escalations</div>
        <div class="l2-metric-value">{escalations:.0f}</div>
        <div class="metric-trend">Exception volume</div>
    </div>
    """, unsafe_allow_html=True)

    # Expandable details
    if st.checkbox("📋 View Execution & Resilience Details", key="obj2_details"):
        st.markdown("#### FTR by Process")

        ftr_summary = ftr_data.groupby('Process').agg({
            'FTR_Rate_Percentage': 'mean',
            'Error_Rate_Percentage': 'mean'
        }).sort_values('FTR_Rate_Percentage')
        st.dataframe(ftr_summary, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# -------- OBJECTIVE 3: WORKFORCE & PRODUCTIVITY --------
with col3:
    st.markdown("""
    <div class="objective-card">
        <div class="objective-title">👥 Workforce & Productivity</div>
        <div class="objective-signal">Efficiency signal: Optimize capacity & output</div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="l2-metric efficiency">
        <div class="l2-metric-title">⚡ Output/FTE</div>
        <div class="l2-metric-value">{avg_output:.2f}</div>
        <div class="metric-trend">Productivity per agent</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="l2-metric efficiency">
        <div class="l2-metric-title">📊 Capacity Utilization</div>
        <div class="l2-metric-value">{avg_capacity:.0f}%</div>
        <div class="metric-trend">Workload balance</div>
    </div>
    """, unsafe_allow_html=True)

    burnout_count = capacity_data[capacity_data['Burnout_Risk_Flag'] == 'Yes'].shape[0]
    st.markdown(f"""
    <div class="l2-metric cost">
        <div class="l2-metric-title">🔥 At-Risk Employees</div>
        <div class="l2-metric-value">{burnout_count}</div>
        <div class="metric-trend">Burnout risk count</div>
    </div>
    """, unsafe_allow_html=True)

    collab_hours = data['Collaboration'][data['Collaboration']['Month'].isin(selected_months)]['Collaboration_Tools_Time_Hours'].mean()
    st.markdown(f"""
    <div class="l2-metric efficiency">
        <div class="l2-metric-title">🗣️ Collab Hours</div>
        <div class="l2-metric-value">{collab_hours:.1f} hrs</div>
        <div class="metric-trend">Daily collaboration time</div>
    </div>
    """, unsafe_allow_html=True)

    # Expandable details
    if st.checkbox("📋 View Workforce & Productivity Details", key="obj3_details"):
        st.markdown("#### Capacity by Department")

        capacity_summary = capacity_data.groupby('Department').agg({
            'Capacity_Utilization_Percentage': 'mean'
        }).sort_values('Capacity_Utilization_Percentage', ascending=False)
        st.dataframe(capacity_summary, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==================== L3: DETAILED DATA & EXPORTS ====================
st.markdown("""<div class="divider-custom"></div>""", unsafe_allow_html=True)

st.markdown("""
<div class="export-section">
    <div class="export-header">📂 L3: Detailed Data & Exports</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📊 KPI Sheet Selector", "📥 Download & Export"])

# -------- TAB 1: KPI SELECTOR --------
with tab1:
    st.markdown("#### Select KPI Sheet to View")

    sheet_map = {
        "Role vs Reality": "Role_vs_Reality",
        "Automation ROI": "Automation_ROI",
        "Digital Workplace Index": "Digital_Index",
        "Process Rework Cost": "Process_Rework",
        "FTR Rate": "FTR_Rate",
        "Process Adherence": "Adherence",
        "Resilience Score": "Resilience",
        "Escalation Patterns": "Escalation",
        "Hidden Capacity & Burnout": "Capacity",
        "Capacity Model Accuracy": "Model_Accuracy",
        "Work Models Effectiveness": "Work_Models",
        "Collaboration Overload": "Collaboration",
    }

    selected_sheet = st.selectbox(
        "🔍 Choose KPI sheet:",
        list(sheet_map.keys()),
        index=0
    )

    key_name = sheet_map[selected_sheet]
    sheet_df = data[key_name].copy()

    # Filter by selected months
    if 'Month' in sheet_df.columns:
        sheet_df = sheet_df[sheet_df['Month'].isin(selected_months)]

    # Filter by departments if applicable
    if 'Department' in sheet_df.columns and dept_filter:
        sheet_df = sheet_df[sheet_df['Department'].isin(dept_filter)]

    st.info(f"📋 **{selected_sheet}** | {len(sheet_df)} records")

    # Show preview
    st.markdown("#### Data Preview (First 50 rows)")
    st.dataframe(sheet_df.head(50), use_container_width=True, hide_index=True)

# -------- TAB 2: DOWNLOAD & EXPORT --------
with tab2:
    st.markdown("#### 📥 Export Options")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Current View (Filtered)**")

        # Current view - combine all filtered data for overview
        all_filtered = pd.concat([
            filter_data(data['Role_vs_Reality']),
            filter_data(data['Automation_ROI']),
            filter_data(data['FTR_Rate']),
        ], ignore_index=True)

        csv_current = all_filtered.to_csv(index=False)
        st.download_button(
            label="📥 Download Current View (CSV)",
            data=csv_current,
            file_name="dashboard_current_view.csv",
            mime="text/csv",
            key="dl_current_view"
        )

    with col2:
        st.markdown("**Selected KPI Sheet**")

        # Download selected sheet
        selected_sheet = st.selectbox(
            "Choose sheet to export:",
            list(sheet_map.keys()),
            index=0,
            key="export_sheet_select"
        )

        key_name = sheet_map[selected_sheet]
        export_df = data[key_name].copy()

        # Filter
        if 'Month' in export_df.columns:
            export_df = export_df[export_df['Month'].isin(selected_months)]
        if 'Department' in export_df.columns and dept_filter:
            export_df = export_df[export_df['Department'].isin(dept_filter)]

        csv_export = export_df.to_csv(index=False)
        st.download_button(
            label=f"📥 Download {selected_sheet} (CSV)",
            data=csv_export,
            file_name=f"{key_name}.csv",
            mime="text/csv",
            key=f"dl_{key_name}"
        )

    st.divider()

    # Excel export for all sheets
    st.markdown("**Export All Sheets as Excel**")

    if st.button("⚙️ Prepare Excel Export (All Sheets)"):
        st.info("💡 This will package all KPI sheets. Processing...")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in data.items():
                filtered_df = df[df['Month'].isin(selected_months)] if 'Month' in df.columns else df
                if 'Department' in filtered_df.columns and dept_filter:
                    filtered_df = filtered_df[filtered_df['Department'].isin(dept_filter)]
                filtered_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        output.seek(0)

        st.download_button(
            label="📥 Download All Sheets (Excel)",
            data=output.getvalue(),
            file_name="COO_Dashboard_All_KPIs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_all_excel"
        )

st.markdown("</div>", unsafe_allow_html=True)

# ==================== FOOTER ====================
st.divider()
st.markdown(f"""
    <div style="text-align: center; padding: 15px; color: #6b7280; font-size: 11px;">
        <strong>COO Dashboard v6.0 - Hierarchical View</strong> | 
        {len(selected_months)} months | {len(dept_filter) if dept_filter else len(all_departments)} departments | 
        Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
    </div>
""", unsafe_allow_html=True)
