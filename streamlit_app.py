
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

st.set_page_config(page_title="Employee KPI Dashboard", layout="wide")
st.markdown("<h1 style='text-align:left'>📊 Employee KPI Dashboard</h1>", unsafe_allow_html=True)
st.write("Auto-loaded from bundled Excel: **Updated_18_KPI_Dashboard.xlsx**")

@st.cache_data
def load_data(path="Updated_18_KPI_Dashboard.xlsx"):
    df = pd.read_excel(path)
    # parse reporting period if exists
    for col in df.columns:
        if 'date' in col.lower() or 'period' in col.lower():
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            except Exception:
                pass
    return df

df = load_data()

# Detect columns
EMP_COL = "Employee_ID" if "Employee_ID" in df.columns else df.columns[0]
DATE_COL = next((c for c in df.columns if 'date' in c.lower() or 'period' in c.lower()), None)
numeric_cols = df.select_dtypes(include='number').columns.tolist()

# Heuristic employee-related metrics (from data)
EMP_METRICS = ['Time_Low_Value_Tasks_Hours', 'Total_Work_Time_Hours', 'Cost_Per_Hour']

# Sidebar filters
st.sidebar.header("Filters")
employees = ["All"] + sorted(df[EMP_COL].dropna().unique().astype(str).tolist())
selected_emp = st.sidebar.selectbox("Select Employee", employees, index=0)
if DATE_COL and pd.api.types.is_datetime64_any_dtype(df[DATE_COL]):
    min_date = df[DATE_COL].min().date()
    max_date = df[DATE_COL].max().date()
    date_range = st.sidebar.date_input("Date range", value=(min_date, max_date))
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        df = df[(df[DATE_COL] >= start) & (df[DATE_COL] <= end)]

if selected_emp != "All":
    df = df[df[EMP_COL].astype(str) == str(selected_emp)]

# Top KPI cards
st.subheader("Key Employee Metrics")
metrics = EMP_METRICS if EMP_METRICS else numeric_cols[:4]
cols = st.columns(min(4, len(metrics)))
for i, m in enumerate(metrics[:4]):
    col = cols[i]
    if m in df.columns:
        val = df[m].dropna()
        if not val.empty:
            latest = val.iloc[-1]
            first = val.iloc[0]
            delta = latest - first
            delta_str = f"{{delta:+.2f}}".format(delta=delta)
            col.metric(label=m.replace('_', ' '), value=f"{{latest:,.2f}}".format(latest=latest), delta=delta_str)
        else:
            col.metric(label=m.replace('_', ' '), value="N/A", delta="N/A")
    else:
        col.write('')


# Trend chart for a selected metric
if metrics:
    st.subheader("Trend / Time Series")
    available_metrics = [m for m in metrics if m in df.columns]
    if available_metrics:
        metric_choice = st.selectbox("Select metric to view trend", available_metrics, index=0)
        if DATE_COL and metric_choice in df.columns:
            fig = px.line(df.sort_values(by=DATE_COL), x=DATE_COL, y=metric_choice, markers=True, title=f"{{metric_choice}} over time")
            st.plotly_chart(fig, use_container_width=True)
        else:
            # fallback: show bar per employee (if all employees)
            if selected_emp == "All" and metric_choice in df.columns:
                agg = df.groupby(EMP_COL)[metric_choice].mean().reset_index().sort_values(metric_choice, ascending=False)
                fig = px.bar(agg, x=EMP_COL, y=metric_choice, title=f"Average {{metric_choice}} by Employee", text_auto=True)
                st.plotly_chart(fig, use_container_width=True)

# Comparison across employees (top N)
st.subheader("Employee Comparison")
comp_metric = st.selectbox("Metric for comparison", [m for m in metrics if m in df.columns], index=0)
top_n = st.sidebar.slider("Show top N employees", 3, min(10, max(3, df[EMP_COL].nunique())), 10)
if selected_emp == "All" and comp_metric in df.columns:
    comp_df = df.groupby(EMP_COL)[comp_metric].mean().reset_index().sort_values(comp_metric, ascending=False).head(top_n)
    fig = px.bar(comp_df, x=EMP_COL, y=comp_metric, title=f"Top {{top_n}} Employees by {{comp_metric}}", text_auto=True)
    st.plotly_chart(fig, use_container_width=True)

# Correlation heatmap for numeric employee metrics
st.subheader("Correlation between numeric metrics")
numeric_for_corr = [c for c in numeric_cols if c in df.columns]
if len(numeric_for_corr) >= 2:
    corr = df[numeric_for_corr].corr()
    fig = px.imshow(corr, text_auto=True, aspect='auto', title='Correlation Matrix')
    st.plotly_chart(fig, use_container_width=True)
else:
    st.write("Not enough numeric columns for correlation.")

# Data and download
st.subheader("Data Preview & Download")
st.dataframe(df.reset_index(drop=True).head(200), use_container_width=True)
csv = df.to_csv(index=False).encode('utf-8')
st.download_button("⬇️ Download filtered data", csv, "employee_filtered_data.csv", "text/csv")
