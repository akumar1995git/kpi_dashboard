import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

st.set_page_config(page_title="Employee KPI Dashboard v6", layout="wide")

st.markdown("<h1 style='text-align:left'>📊 Employee KPI Dashboard v6</h1>", unsafe_allow_html=True)

@st.cache_data
def load_all_sheets(path="Updated_18_KPI_Dashboard.xlsx"):
    xls = pd.ExcelFile(path)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        rp_cols = [c for c in df.columns if "report" in str(c).lower() or "period" in str(c).lower() or "date" in str(c).lower()]
        if rp_cols:
            col = rp_cols[0]
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce")
            except Exception:
                pass
        sheets[name] = df
    return sheets

sheets = load_all_sheets()

# Sidebar controls
st.sidebar.header("Controls")
sheet_name = st.sidebar.selectbox("Select sheet (department/team)", list(sheets.keys()))
df = sheets[sheet_name].copy()

# Detect key columns
EMP_COL = "Employee_ID" if "Employee_ID" in df.columns else next((c for c in df.columns if "employee" in str(c).lower()), df.columns[0])
RP_COL = next((c for c in df.columns if "report" in str(c).lower() or "period" in str(c).lower() or "date" in str(c).lower()), None)

if RP_COL and not pd.api.types.is_datetime64_any_dtype(df[RP_COL]):
    df[RP_COL] = pd.to_datetime(df[RP_COL], errors="coerce")

numeric_cols = df.select_dtypes(include="number").columns.tolist()
numeric_cols = [c for c in numeric_cols if c not in [EMP_COL, RP_COL]]

# Sidebar: employee multi-select
employees = sorted(df[EMP_COL].dropna().astype(str).unique().tolist())
selected_employees = st.sidebar.multiselect("Select Employees", employees, default=employees[:5])

if selected_employees:
    df = df[df[EMP_COL].astype(str).isin(selected_employees)]

# Sidebar: date filter
if RP_COL:
    min_date, max_date = df[RP_COL].min(), df[RP_COL].max()
    date_range = st.sidebar.date_input("Reporting Period", value=(min_date, max_date))
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        df = df[(df[RP_COL] >= start) & (df[RP_COL] <= end)]

# Aggregation
agg_choice = st.sidebar.selectbox("Aggregation for KPIs", ["mean", "sum"])

# Summary Section
st.markdown("## Summary Metrics")

if RP_COL:
    df["_month"] = df[RP_COL].dt.to_period("M").dt.to_timestamp()
    months = sorted(df["_month"].dropna().unique())
    latest_month = months[-1] if months else None
    prev_month = months[-2] if len(months) >= 2 else None
else:
    df["_month"], latest_month, prev_month = None, None, None

chosen_metrics = st.multiselect("Select metrics to summarize", numeric_cols, default=numeric_cols[:5])
cols = st.columns(max(1, len(chosen_metrics)))

def agg_series(data, metric, agg):
    if data.empty:
        return np.nan
    return round(data[metric].mean(skipna=True), 2) if agg == "mean" else round(data[metric].sum(skipna=True), 2)

if latest_month is not None:
    latest_df = df[df["_month"] == latest_month]
    prev_df = df[df["_month"] == prev_month] if prev_month is not None else pd.DataFrame(columns=df.columns)

    for i, m in enumerate(chosen_metrics):
        cur_val = agg_series(latest_df, m, agg_choice)
        prev_val = agg_series(prev_df, m, agg_choice) if not prev_df.empty else np.nan
        delta_display = "N/A"
        if pd.notna(prev_val) and prev_val != 0:
            change = round((cur_val - prev_val) / abs(prev_val) * 100, 2)
            delta_display = f"🟢▲ {abs(change)}%" if change > 0 else f"🔴▼ {abs(change)}%"
        cols[i].metric(m.replace("_", " "), f"{cur_val:,.2f}", delta_display)
else:
    for i, m in enumerate(chosen_metrics):
        val = agg_series(df, m, agg_choice)
        cols[i].metric(m.replace("_", " "), f"{val:,.2f}")

st.markdown("---")

# Monthly Trend
st.markdown("## Monthly Trends")
metrics_for_trend = st.multiselect("Select metrics to plot", numeric_cols, default=chosen_metrics[:2])
if RP_COL and metrics_for_trend:
    trend_df = df.groupby("_month")[metrics_for_trend].agg(agg_choice).reset_index()
    trend_df = trend_df.round(2)
    if not trend_df.empty:
        fig = px.bar(
            trend_df.melt(id_vars="_month", var_name="Metric", value_name="Value"),
            x="_month", y="Value", color="Metric", barmode="group",
            title="Monthly KPI Trends"
        )
        fig.update_layout(xaxis_title="Month", yaxis_title="Value", hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Top Employees Chart
st.markdown("## Top Employees")
comp_metric = st.selectbox("Select metric for comparison", numeric_cols)
max_employees = min(50, len(df[EMP_COL].unique()))
top_n = st.slider("Number of top employees to display", min_value=3, max_value=max_employees, value=min(5, max_employees))
if comp_metric:
    comp_df = df.groupby(EMP_COL)[comp_metric].agg(agg_choice).reset_index()
    comp_df = comp_df.sort_values(comp_metric, ascending=False).head(top_n)
    comp_df[comp_metric] = comp_df[comp_metric].round(2)
    fig = px.bar(
        comp_df,
        x=EMP_COL, y=comp_metric, text=comp_metric,
        title=f"Top {top_n} Employees by {comp_metric} ({agg_choice})"
    )
    fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig.update_layout(xaxis_title="Employee", yaxis_title=comp_metric, uniformtext_minsize=8, uniformtext_mode='hide')
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Data Preview
st.markdown("## Data Preview & Download")
st.dataframe(df.round(2).head(200), use_container_width=True)
csv = df.round(2).to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download filtered data", csv, file_name=f"{sheet_name}_filtered.csv", mime="text/csv")
