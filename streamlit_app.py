import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide")

# ========================
# Load Excel Data
# ========================
@st.cache_data
def load_data():
    file_path = "Updated_18_KPI_Dashboard.xlsx"
    if not os.path.exists(file_path):
        st.error(f"File '{file_path}' not found. Please upload it.")
        st.stop()
    xls = pd.ExcelFile(file_path)
    sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
    return sheets

data_sheets = load_data()

# ========================
# Sidebar
# ========================
st.sidebar.header("🔍 Filters")

sheet_name = st.sidebar.selectbox("Select Department/Sheet", list(data_sheets.keys()))
df = data_sheets[sheet_name]

# Validate data
if "Employee_ID" not in df.columns or "Reporting_Period" not in df.columns:
    st.error("The sheet must have 'Employee_ID' and 'Reporting_Period' columns.")
    st.stop()

df["Reporting_Period"] = pd.to_datetime(df["Reporting_Period"])

employees = sorted(df["Employee_ID"].unique())
selected_employee = st.sidebar.multiselect("Select Employees", employees, default=employees)
filtered_df = df[df["Employee_ID"].isin(selected_employee)]

metric_cols = [col for col in df.columns if col not in ["Employee_ID", "Reporting_Period"]]
selected_metric = st.sidebar.selectbox("Select KPI Metric", metric_cols)

# ========================
# Dashboard Header
# ========================
st.markdown(
    """
    <h2 style='text-align:center; color:#1E3A8A;'>
    📊 Employee KPI Performance Dashboard
    </h2>
    """,
    unsafe_allow_html=True
)

# ========================
# KPI Summary Cards
# ========================
latest_period = filtered_df["Reporting_Period"].max()
previous_period = filtered_df["Reporting_Period"].sort_values().unique()[-2] if len(filtered_df["Reporting_Period"].unique()) > 1 else None

current_df = filtered_df[filtered_df["Reporting_Period"] == latest_period]
previous_df = filtered_df[filtered_df["Reporting_Period"] == previous_period] if previous_period else None

summary = []
for metric in metric_cols:
    current_avg = current_df[metric].mean()
    prev_avg = previous_df[metric].mean() if previous_df is not None else current_avg
    change = current_avg - prev_avg
    direction = "🟢▲ Improved" if change > 0 else "🔴▼ Declined" if change < 0 else "🟡 No Change"
    summary.append((metric, round(current_avg, 2), direction))

st.markdown("### 📈 Summary Overview")
cols = st.columns(3)
for i, (metric, avg, direction) in enumerate(summary):
    with cols[i % 3]:
        st.markdown(
            f"""
            <div style="background-color:#F9FAFB;padding:20px;border-radius:15px;
            box-shadow:1px 2px 5px rgba(0,0,0,0.1);margin-bottom:10px;">
            <h4 style="color:#2563EB;">{metric}</h4>
            <h2 style="margin:0;">{avg}</h2>
            <p style="color:#10B981;">{direction}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

# ========================
# Main KPI Chart
# ========================
st.markdown(f"### 📊 {selected_metric} over Time")

chart_type = "bar" if any(word in selected_metric.lower() for word in ["count", "num", "total", "hours", "rating"]) else "line"

if chart_type == "line":
    fig = px.line(filtered_df, x="Reporting_Period", y=selected_metric, color="Employee_ID",
                  markers=True, title=f"{selected_metric} Trend")
else:
    fig = px.bar(filtered_df, x="Reporting_Period", y=selected_metric, color="Employee_ID",
                 barmode="group", title=f"{selected_metric} Comparison")

fig.update_layout(
    xaxis_title="Reporting Period",
    yaxis_title=selected_metric,
    template="plotly_white",
    hovermode="x unified",
    height=450
)
st.plotly_chart(fig, use_container_width=True)

# ========================
# Top Employees
# ========================
st.markdown("### 🏆 Top Performing Employees")
avg_scores = filtered_df.groupby("Employee_ID")[metric_cols].mean().reset_index()
avg_scores["Overall_Performance"] = avg_scores[metric_cols].mean(axis=1)
top_employees = avg_scores.nlargest(5, "Overall_Performance")

fig2 = px.bar(
    top_employees,
    x="Employee_ID",
    y="Overall_Performance",
    title="Top 5 Employees by Average KPI",
    text_auto=True,
    color="Overall_Performance",
    color_continuous_scale="Blues"
)
fig2.update_layout(template="plotly_white", height=400)
st.plotly_chart(fig2, use_container_width=True)
