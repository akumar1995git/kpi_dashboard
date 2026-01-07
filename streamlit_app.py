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

    /* Sub-objective Box with inline sparkline */
    .subobjective-box {
        background: #f9fafb;
        border-left: 4px solid #1e40af;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 10px;
        display: grid;
        grid-template-columns: 1fr auto;
        gap: 12px;
        align-items: center;
        cursor: pointer;
        transition: all 0.2s ease;
    }

    .subobjective-box:hover {
        background: #f3f4f6;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }

    .subobjective-box.cost { border-left-color: #ef4444; }
    .subobjective-box.quality { border-left-color: #059669; }
    .subobjective-box.efficiency { border-left-color: #f59e0b; }

    .subobjective-info {
        flex-grow: 1;
    }

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

    .sparkline-container {
        width: 90px;
        height: 40px;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    /* Navigation buttons */
    .nav-container {
        display: flex;
        gap: 8px;
        margin-bottom: 20px;
        flex-wrap: wrap;
    }

    .nav-button {
        padding: 8px 14px;
        border-radius: 6px;
        border: 1px solid #d1d5db;
        background: white;
        cursor: pointer;
        font-size: 12px;
        font-weight: 500;
        transition: all 0.2s ease;
    }

    .nav-button:hover {
        background: #f3f4f6;
        border-color: #1e40af;
    }

    .insights-box {
        background: #fff7ed;
        border-left: 5px solid #ea580c;
        padding: 10px 16px;
        border-radius: 8px;
        font-size: 12px;
        margin-bottom: 8px;
    }

    .recommendation-box {
        background: #fee2e2;
        border-left: 5px solid #ef4444;
        padding: 10px 16px;
        border-radius: 8px;
        font-size: 12px;
        margin-bottom: 8px;
    }

    .trend-badge {
        font-size: 10px;
        border-radius: 999px;
        padding: 3px 8px;
        display: inline-block;
        margin-left: 6px;
    }

    .trend-up {
        background: #dcfce7;
        color: #166534;
    }

    .trend-down {
        background: #fee2e2;
        color: #991b1b;
    }

    .trend-flat {
        background: #e5e7eb;
        color: #374151;
    }

    .metric-caption {
        font-size: 11px;
        color: #6b7280;
        margin-top: 2px;
    }

    .section-divider {
        margin-top: 10px;
        margin-bottom: 4px;
        border-top: 1px solid #e5e7eb;
    }

    .header-subtitle {
        font-size: 12px;
        color: #6b7280;
        margin-top: -4px;
        margin-bottom: 10px;
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
        st.error("File not found: 'COO_ROI_Dashboard_KPIs_Complete_12.xlsx'")
        st.stop()

data = load_excel_data()

# ==================== SESSION STATE ====================
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'main'

# ==================== SIDEBAR FILTERS ====================
st.sidebar.markdown("## Filters")

role_months = sorted(data['Role_vs_Reality']['Month'].unique())
selected_months = st.sidebar.multiselect(
    "Select Months",
    role_months,
    default=list(role_months),
)

all_departments = sorted(
    list(
        set(
            list(data['Role_vs_Reality'].get('Department', []).unique())
            + list(data['Capacity'].get('Department', []).unique())
        )
    )
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

def get_latest_month_data(df, month_col='Month'):
    """Return only latest month rows from a dataframe."""
    if df.empty or month_col not in df.columns:
        return df
    latest = df[month_col].max()
    return df[df[month_col] == latest]

def create_trend_chart(df, x_col, y_col, title, color='#1e40af', height=280):
    """Create a trend line chart"""
    if len(df) < 2:
        return None

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df[x_col],
        y=df[y_col],
        mode='lines+markers',
        line=dict(color=color, width=3),
        marker=dict(size=8),
        fill='tozeroy',
        fillcolor=f'rgba(30, 64, 175, 0.1)',
        name=y_col
    ))

    fig.update_layout(
        title=title,
        height=height,
        margin=dict(l=0, r=0, t=30, b=0),
        showlegend=True,
        plot_bgcolor="rgba(0,0,0,0)",
        hovermode='x unified',
        yaxis=dict(rangemode='nonnegative')
    )

    return fig

def create_sparkline(df, x_col, y_col, color='#1e40af'):
    """Create a compact sparkline chart for inline display"""
    if len(df) < 2:
        return None

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df[x_col],
        y=df[y_col],
        mode='lines',
        line=dict(color=color, width=2.5),
        fill='tozeroy',
        fillcolor=f'rgba({int(color[1:3], 16)}, {int(color[3:5], 16)}, {int(color[5:7], 16)}, 0.15)',
        hoverinfo='y',
        name=y_col
    ))

    fig.update_layout(
        margin=dict(l=0, r=0, t=0, b=0),
        height=40,
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)"
    )

    return fig

def round_value(val, mode='whole'):
    if pd.isna(val):
        return 0
    if mode == 'percentage':
        return float(val)
    if mode == 'currency':
        return float(val)
    if mode == 'hours':
        return float(val)
    return float(val)

def format_trend(delta):
    if delta > 0.1:
        return "↑", "trend-up"
    elif delta < -0.1:
        return "↓", "trend-down"
    else:
        return "→", "trend-flat"

def show_navigation():
    st.markdown(
        """
        <div class="nav-container">
            <button class="nav-button" onclick="window.location.reload()">Overview</button>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==================== MAIN PAGE ====================
if st.session_state.current_page == 'main':
    st.title("COO Operational Dashboard")
    st.markdown(
        "<p class='header-subtitle'>Executive KPI Management with Visual Analytics</p>",
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.markdown("### Objective Summary")
        obj_col1, obj_col2 = st.columns([1, 1])

        with obj_col1:
            rework_data = filter_data(data['Process_Rework'])
            auto_data = filter_data(data['Automation_ROI'])
            digital_data = filter_data(data['Digital_Index'])
            role_data = filter_data(data['Role_vs_Reality'])

            rework_pct = rework_data['Rework_Cost_Percentage'].mean() if len(rework_data) > 0 else 0
            auto_roi = auto_data['ROI_Percentage_6M'].mean() if len(auto_data) > 0 else 0
            friction = digital_data['Friction_Index_Score'].mean() if len(digital_data) > 0 else 0
            role_reality = role_data['Low_Value_Work_Percentage'].mean() if len(role_data) > 0 else 0

            def get_month_over_month_change(df, value_col, month_col='Month'):
                if len(df) < 2:
                    return 0, 0
                grouped = df.groupby(month_col)[value_col].mean().reset_index().sort_values(month_col)
                if len(grouped) < 2:
                    return grouped[value_col].iloc[-1], 0
                latest = grouped[value_col].iloc[-1]
                prev = grouped[value_col].iloc[-2]
                delta = latest - prev
                return latest, delta

            rework_val, rework_change = get_month_over_month_change(rework_data, 'Rework_Cost_Percentage')
            auto_val, auto_change = get_month_over_month_change(auto_data, 'ROI_Percentage_6M')
            friction_val, friction_change = get_month_over_month_change(digital_data, 'Friction_Index_Score')
            role_val, role_change = get_month_over_month_change(role_data, 'Low_Value_Work_Percentage')

            st.markdown("#### Cost & Efficiency")
            icon, cls = format_trend(-rework_change)
            st.metric("Rework Cost %", f"{rework_val:.1f}%", delta=f"{-rework_change:.1f} pp")
            st.caption(f"Rework trending {icon}", help="Month-over-month change")

            st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

            icon, cls = format_trend(auto_change)
            st.metric("Automation ROI (6M)", f"{auto_val:.0f}%", delta=f"{auto_change:.1f} pp")

        with obj_col2:
            st.markdown("#### Digital Friction & Role Reality")
            icon, cls = format_trend(friction_change)
            st.metric("Digital Friction Index", f"{friction_val:.1f}", delta=f"{friction_change:.1f}")

            icon, cls = format_trend(role_change)
            st.metric("Low-Value Work %", f"{role_val:.1f}%", delta=f"{role_change:.1f} pp")

    with col2:
        st.markdown("### Digital Workplace Index")
        digital_data = filter_data(data['Digital_Index'])
        if len(digital_data) > 0:
            latest_score = digital_data.sort_values('Month').groupby('Department')['Digital_Workplace_Index_Score'].last()
            fig = go.Figure(data=[
                go.Indicator(
                    mode="gauge+number",
                    value=latest_score.mean(),
                    title={'text': "Overall Index"},
                    gauge={'axis': {'range': [0, 100]}}
                )
            ])
            fig.update_layout(height=260)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ===== Immediate Action Insights (TOP SECTION) =====
    rework_data = filter_data(data['Process_Rework'])
    auto_data = filter_data(data['Automation_ROI'])
    digital_data = filter_data(data['Digital_Index'])
    role_data = filter_data(data['Role_vs_Reality'])
    work_data = filter_data(data['Work_Models'])

    st.markdown("### Action Insights")
    col_action1, col_action2 = st.columns([1, 1])

    with col_action1:
        st.markdown("**Immediate Attention Required:**")
        if len(role_data) > 0:
            # Use only latest month to avoid duplication across months
            role_data_latest = get_latest_month_data(role_data, month_col='Month')
            top_low_value = role_data_latest.nlargest(
                5,
                'Opportunity_Cost_Dollars'
            )[['Employee_ID', 'Role', 'Low_Value_Work_Percentage', 'Opportunity_Cost_Dollars']]
            for _, row in top_low_value.iterrows():
                st.markdown(
                    f'<div class="insights-box">'
                    f'{row["Employee_ID"]} ({row["Role"]}): '
                    f'{row["Low_Value_Work_Percentage"]:.1f}% low-value work - '
                    f'${row["Opportunity_Cost_Dollars"]:,.0f}/month'
                    f'</div>',
                    unsafe_allow_html=True
                )

    with col_action2:
        st.markdown("**Top Automation Opportunities:**")
        if len(auto_data) > 0:
            top_auto = auto_data.nlargest(
                5,
                'ROI_Percentage_6M'
            )[['Process_Name', 'Time_Savings_Hours'] if 'Time_Savings_Hours' in auto_data.columns
              else ['Process_Name', 'Monthly_Hours_Saved']]
            time_col = 'Time_Savings_Hours' if 'Time_Savings_Hours' in auto_data.columns else 'Monthly_Hours_Saved'
            for _, row in top_auto.iterrows():
                st.markdown(
                    f'<div class="recommendation-box">'
                    f'{row["Process_Name"]}: {row[time_col]:.0f} hours/month potential savings'
                    f'</div>',
                    unsafe_allow_html=True
                )

    st.divider()
    st.markdown("---")
    st.markdown("#### Detailed Data & Export")

    tabs = st.tabs(["Rework Cost", "Automation ROI", "Digital Index", "Work Models", "Role Analysis"])

    with tabs[0]:
        st.dataframe(rework_data.head(100), use_container_width=True, hide_index=True)
        csv = rework_data.to_csv(index=False)
        st.download_button("Download Rework Data (CSV)", data=csv,
                           file_name="rework_data.csv", mime="text/csv", key="dl_rework")

    with tabs[1]:
        st.dataframe(auto_data.head(100), use_container_width=True, hide_index=True)
        csv = auto_data.to_csv(index=False)
        st.download_button("Download Automation Data (CSV)", data=csv,
                           file_name="automation_data.csv", mime="text/csv", key="dl_auto")

    with tabs[2]:
        st.dataframe(digital_data.head(100), use_container_width=True, hide_index=True)
        csv = digital_data.to_csv(index=False)
        st.download_button("Download Digital Index (CSV)", data=csv,
                           file_name="digital_index.csv", mime="text/csv", key="dl_digital")

    with tabs[3]:
        st.dataframe(work_data.head(100), use_container_width=True, hide_index=True)
        csv = work_data.to_csv(index=False)
        st.download_button("Download Work Models (CSV)", data=csv,
                           file_name="work_models.csv", mime="text/csv", key="dl_work")

    with tabs[4]:
        st.dataframe(role_data.head(100), use_container_width=True, hide_index=True)
        csv = role_data.to_csv(index=False)
        st.download_button("Download Role Analysis (CSV)", data=csv,
                           file_name="role_analysis.csv", mime="text/csv", key="dl_role")

# ==================== DETAIL PAGE 2: EXECUTION & RESILIENCE ====================
elif st.session_state.current_page == 'execution_resilience':
    show_navigation()
    st.markdown("### Execution & Resilience - Deep Dive")

    # KPI data
    ftr_data = filter_data(data['FTR_Rate'])
    adherence_data = filter_data(data['Adherence'])
    resilience_data = filter_data(data['Resilience'])
    escalation_data = filter_data(data['Escalation'])

    st.markdown("#### Detailed Metrics with Trends & Analysis")

    # ROW 1: FTR Rate
    st.markdown("**First-Time-Right (FTR) Rate**")
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.markdown("**KPI Card**")
        ftr_rate = ftr_data['FTR_Rate_Percentage'].mean() if len(ftr_data) > 0 else 0
        st.metric(label="FTR Rate", value=f"{ftr_rate:.1f}%", delta="+1.2%")

    with col2:
        st.markdown("**FTR Trend**")
        if len(ftr_data) > 0:
            ftr_trend = ftr_data.groupby('Month').agg({'FTR_Rate_Percentage': 'mean'}).reset_index().sort_values('Month')
            fig = create_trend_chart(ftr_trend, 'Month', 'FTR_Rate_Percentage', 'FTR Trend', '#16a34a', height=280)
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)

    with col3:
        st.markdown("**By Process**")
        if len(ftr_data) > 0:
            ftr_by_proc = ftr_data.groupby('Process_Name').agg(
                {'FTR_Rate_Percentage': 'mean'}
            ).sort_values('FTR_Rate_Percentage', ascending=True).head(6)
            fig = go.Figure(data=[
                go.Bar(
                    y=ftr_by_proc.index,
                    x=ftr_by_proc['FTR_Rate_Percentage'],
                    orientation='h',
                    marker_color='#16a34a',
                    text=[f"{x:.1f}%" for x in ftr_by_proc['FTR_Rate_Percentage']],
                    textposition='outside',
                    name='FTR Rate'
                )
            ])
            fig.update_layout(
                height=280,
                showlegend=True,
                plot_bgcolor="rgba(0,0,0,0)",
                hovermode='y unified'
            )
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ROW 2: Adherence & Resilience
    st.markdown("**Process Adherence & Resilience**")
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.markdown("**Adherence KPI**")
        adherence = adherence_data['Adherence_Rate_Percentage'].mean() if len(adherence_data) > 0 else 0
        st.metric(label="Adherence Rate", value=f"{adherence:.1f}%", delta="+0.8%")

    with col2:
        st.markdown("**Resilience KPI**")
        resilience = resilience_data['Resilience_Score'].mean() if len(resilience_data) > 0 else 0
        st.metric(label="Resilience Score", value=f"{resilience:.1f}", delta="+0.5")

    with col3:
        st.markdown("**Departments Below Target**")
        if len(adherence_data) > 0:
            below_target = adherence_data[adherence_data['Adherence_Rate_Percentage'] < 90]
            by_dept = below_target.groupby('Department').agg(
                {'Adherence_Rate_Percentage': 'mean'}
            ).sort_values('Adherence_Rate_Percentage').head(5)
            for dept, row in by_dept.iterrows():
                st.markdown(
                    f'<div class="insights-box">'
                    f'{dept}: Adherence {row["Adherence_Rate_Percentage"]:.1f}% - Coaching required'
                    f'</div>',
                    unsafe_allow_html=True
                )

    st.divider()

    # ROW 3: Resilience & Escalation Insights
    st.markdown("**Execution Risks & Escalations**")
    col_action1, col_action2 = st.columns([1, 1])

    with col_action1:
        st.markdown("**Resilience Risks by Process:**")
        if len(resilience_data) > 0:
            low_res = resilience_data.sort_values('Resilience_Score').head(5)
            for _, row in low_res.iterrows():
                if 'Process_Name' in row:
                    st.markdown(
                        f'<div class="recommendation-box">'
                        f'{row["Process_Name"]} ({row["Department"]}): '
                        f'Resilience Score {row["Resilience_Score"]:.1f} - '
                        f'Workload balancing needed'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        f'<div class="recommendation-box">'
                        f'{row["Department"]}: Resilience Score {row["Resilience_Score"]:.1f} - '
                        f'Workload balancing needed'
                        f'</div>',
                        unsafe_allow_html=True
                    )

    with col_action2:
        st.markdown("**Escalation Risk Hotspots:**")
        if len(escalation_data) > 0:
            # latest month only
            escalation_data_latest = get_latest_month_data(escalation_data, month_col='Month')
            top_escalations = (
                escalation_data_latest.nlargest(5, 'Step_Exception_Count')[['Process', 'Step_Exception_Count']]
                if 'Process' in escalation_data_latest.columns
                else escalation_data_latest.nlargest(5, 'Step_Exception_Count')
            )
            for _, row in top_escalations.iterrows():
                process_name = row['Process'] if 'Process' in escalation_data_latest.columns else 'Unknown Process'
                count = row['Step_Exception_Count']
                st.markdown(
                    f'<div class="insights-box">'
                    f'{process_name}: {int(count)} exceptions - Root cause analysis needed'
                    f'</div>',
                    unsafe_allow_html=True
                )

    st.divider()

    st.markdown("#### Detailed Data & Export")

    tabs = st.tabs(["FTR Rate", "Adherence", "Resilience", "Escalations"])

    with tabs[0]:
        st.dataframe(ftr_data.head(100), use_container_width=True, hide_index=True)
        csv = ftr_data.to_csv(index=False)
        st.download_button("Download FTR Data (CSV)", data=csv,
                           file_name="ftr_data.csv", mime="text/csv", key="dl_ftr")

    with tabs[1]:
        st.dataframe(adherence_data.head(100), use_container_width=True, hide_index=True)
        csv = adherence_data.to_csv(index=False)
        st.download_button("Download Adherence Data (CSV)", data=csv,
                           file_name="adherence_data.csv", mime="text/csv", key="dl_adherence")

    with tabs[2]:
        st.dataframe(resilience_data.head(100), use_container_width=True, hide_index=True)
        csv = resilience_data.to_csv(index=False)
        st.download_button("Download Resilience Data (CSV)", data=csv,
                           file_name="resilience_data.csv", mime="text/csv", key="dl_resilience")

    with tabs[3]:
        st.dataframe(escalation_data.head(100), use_container_width=True, hide_index=True)
        csv = escalation_data.to_csv(index=False)
        st.download_button("Download Escalation Data (CSV)", data=csv,
                           file_name="escalation_data.csv", mime="text/csv", key="dl_escalation")

# ==================== DETAIL PAGE 1: COST & PRODUCTIVITY (Rework & Automation) ====================
elif st.session_state.current_page == 'cost_productivity':
    show_navigation()
    st.markdown("### Cost & Productivity - Deep Dive")

    rework_data = filter_data(data['Process_Rework'])
    auto_data = filter_data(data['Automation_ROI'])
    digital_data = filter_data(data['Digital_Index'])
    role_data = filter_data(data['Role_vs_Reality'])
    work_data = filter_data(data['Work_Models'])

    st.markdown("#### Detailed Metrics with Trends & Analysis")

    # ROW 1: Rework Cost Analysis
    st.markdown("**Rework Cost Analysis**")
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.markdown("**KPI Card**")
        rework_pct = rework_data['Rework_Cost_Percentage'].mean()
        rework_dollars = rework_data['Rework_Cost_Dollars'].sum()
        st.metric(label="Rework Cost %", value=f"{round_value(rework_pct, 'percentage'):.1f}%", delta="-0.3%")
        st.metric(label="Total Rework $", value=f"${round_value(rework_dollars, 'currency'):,.0f}")

    with col2:
        st.markdown("**By Process (Cost & %)**")
        if len(rework_data) > 0:
            process_rework = rework_data.groupby('Process_Name').agg({
                'Rework_Cost_Dollars': 'sum',
                'Rework_Cost_Percentage': 'mean'
            }).sort_values('Rework_Cost_Dollars', ascending=False).head(6)

            fig = go.Figure(data=[
                go.Bar(
                    y=process_rework.index,
                    x=process_rework['Rework_Cost_Dollars'],
                    orientation='h',
                    marker=dict(
                        color=process_rework['Rework_Cost_Dollars'],
                        colorscale='Reds',
                        showscale=False,
                        line=dict(width=0)
                    ),
                    name='Cost ($)',
                    text=[f"${x:,.0f}<br>({p:.1f}%)"
                          for x, p in zip(process_rework['Rework_Cost_Dollars'],
                                          process_rework['Rework_Cost_Percentage'])],
                    textposition='outside'
                )
            ])
            fig.update_layout(
                height=280,
                showlegend=True,
                plot_bgcolor="rgba(0,0,0,0)",
                hovermode='y unified'
            )
            st.plotly_chart(fig, use_container_width=True)

    with col3:
        st.markdown("**By Department**")
        if len(rework_data) > 0:
            dept_rework = rework_data.groupby('Department').agg({
                'Rework_Cost_Dollars': 'sum',
                'Rework_Cost_Percentage': 'mean'
            }).sort_values('Rework_Cost_Dollars', ascending=False)

            fig = go.Figure(data=[
                go.Bar(
                    y=dept_rework.index,
                    x=dept_rework['Rework_Cost_Dollars'],
                    orientation='h',
                    marker=dict(
                        color=dept_rework['Rework_Cost_Dollars'],
                        colorscale='Reds',
                        showscale=False,
                        line=dict(width=0)
                    ),
                    name='Cost ($)',
                    text=[f"${x:,.0f}<br>({p:.1f}%)"
                          for x, p in zip(dept_rework['Rework_Cost_Dollars'],
                                          dept_rework['Rework_Cost_Percentage'])],
                    textposition='outside'
                )
            ])
            fig.update_layout(
                height=280,
                showlegend=True,
                plot_bgcolor="rgba(0,0,0,0)",
                hovermode='y unified'
            )
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ROW 2: Automation ROI Analysis
    st.markdown("**Automation ROI Potential**")
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.markdown("**KPI Card**")
        auto_roi = auto_data['ROI_Percentage_6M'].mean()
        time_savings = 0
        if 'Time_Savings_Hours' in auto_data.columns:
            time_savings = auto_data['Time_Savings_Hours'].sum()
        elif 'Monthly_Hours_Saved' in auto_data.columns:
            time_savings = auto_data['Monthly_Hours_Saved'].sum()

        st.metric(label="Automation ROI", value=f"{round_value(auto_roi, 'whole'):.0f}%", delta="+4.5%")
        st.metric(label="Time Savings", value=f"{round_value(time_savings, 'hours'):,.1f} hrs", delta="+450 hrs")

    with col2:
        st.markdown("**ROI Trend**")
        if len(auto_data) > 0:
            auto_trend = auto_data.groupby('Month').agg({'ROI_Percentage_6M': 'mean'}).reset_index().sort_values('Month')
            if len(auto_trend) > 1:
                fig = create_trend_chart(auto_trend, 'Month', 'ROI_Percentage_6M', 'ROI Trend', '#059669', height=280)
                if fig is not None:
                    st.plotly_chart(fig, use_container_width=True)

    with col3:
        st.markdown("**Top Automation Candidates**")
        if len(auto_data) > 0:
            top_auto = auto_data.nlargest(
                5,
                'ROI_Percentage_6M'
            )[['Process_Name', 'ROI_Percentage_6M', 'Time_Savings_Hours']
              if 'Time_Savings_Hours' in auto_data.columns
              else ['Process_Name', 'ROI_Percentage_6M', 'Monthly_Hours_Saved']]
            for _, row in top_auto.iterrows():
                hours_col = 'Time_Savings_Hours' if 'Time_Savings_Hours' in row else 'Monthly_Hours_Saved'
                st.markdown(
                    f'<div class="recommendation-box">'
                    f'{row["Process_Name"]}: ROI {row["ROI_Percentage_6M"]:.0f}% - '
                    f'{row[hours_col]:.0f} hrs/month'
                    f'</div>',
                    unsafe_allow_html=True
                )

    st.divider()

    # ROW 3: Digital Friction & Work Models
    st.markdown("**Digital Work Models & Friction**")
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.markdown("**Digital Friction Index**")
        friction = digital_data['Friction_Index_Score'].mean() if len(digital_data) > 0 else 0
        st.metric(label="Friction Score", value=f"{friction:.1f}", delta="-0.2")

    with col2:
        st.markdown("**Work Model Effectiveness**")
        if len(work_data) > 0 and 'Effectiveness_Score' in work_data.columns:
            work_effect = work_data.groupby('Work_Model_Type').agg(
                {'Effectiveness_Score': 'mean'}
            ).sort_values('Effectiveness_Score', ascending=False)
            fig = go.Figure(data=[
                go.Bar(
                    x=work_effect.index,
                    y=work_effect['Effectiveness_Score'],
                    marker_color='#1e40af',
                    text=[f"{x:.1f}" for x in work_effect['Effectiveness_Score']],
                    textposition='outside',
                    name='Effectiveness'
                )
            ])
            fig.update_layout(
                height=260,
                showlegend=True,
                plot_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig, use_container_width=True)

    with col3:
        st.markdown("**Clusters with High Friction**")
        if len(digital_data) > 0:
            high_friction = digital_data.sort_values('Friction_Index_Score', ascending=False).head(5)
            for _, row in high_friction.iterrows():
                st.markdown(
                    f'<div class="insights-box">'
                    f'{row["Department"]}: Friction Score {row["Friction_Index_Score"]:.1f} - '
                    f'Requires digital workflow redesign'
                    f'</div>',
                    unsafe_allow_html=True
                )
