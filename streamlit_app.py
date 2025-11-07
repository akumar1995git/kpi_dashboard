import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# ========== PAGE CONFIG ==========
st.set_page_config(page_title="KPI Dashboard v3", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
        /* Global style */
        body {font-family: 'Inter', sans-serif;}
        .block-container {padding-top: 1rem; padding-bottom: 1rem;}

        /* KPI Cards */
        div[data-testid="metric-container"] {
            background: #f8f9fa;
            border: 1px solid #ddd;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 1px 1px 4px rgba(0,0,0,0.05);
        }
        div[data-testid="metric-container"] > label {
            font-weight: 600;
        }

        /* Section Titles */
        h2, h3 {
            border-left: 4px solid #4F8BF9;
            padding-left: 8px;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("<h1 style='text-align:left'>📊 Employee KPI Dashboard</h1>", unsafe_allow_html=True)

# ========== LOAD ALL SHEETS ==========
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
            df[col] = pd.to_datetime(df[col], errors="coerce")
        sheets[name] = df
    return sheets

sheets = load_all_sheets()

# ========== SIDEBAR ==========
st.sidebar.header("🔧 Filters & Settings")
sheet_name = st.sidebar.selectbox("Select Department Sheet", list(sheets.keys()))
df = sheets[sheet_name].copy()

EMP_COL = "Employee_ID" if "Employee_ID" in df.columns else next((c for c in df.columns if "employee" in str(c).lower()), df.columns[0])
RP_COL = next((c for c in df.columns if "report" in str(c).lower() or "period" in str(c).lower() or "date" in str(c).lower()), None)

# Ensure Reporting_Period is datetime
if RP_COL and not pd.api.types.is_datetime64_any_dtype(df[RP_COL]):
    df[RP_COL] = pd.to_datetime(df[RP_COL], errors="coerce")

numeric_cols = df.select_dtypes(include="number").columns.tolist()
numeric_cols = [c for c in numeric_cols if c not in [EMP_COL, RP_COL]]

# Sidebar filters
employees = ["All"] + sorted(df[EMP_COL].dropna().astype(str).unique().tolist())
selected_employee = st.sidebar.selectbox("Select Employee", employees, index=0)

if RP_COL:
    min_date, max_date = df[RP_COL].min().date(), df[RP_COL].max().date()
    date_range = st.sidebar.date_input("Reporting Period", (min_date, max_date))
    if len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        df = df[(df[RP_COL] >= start) & (df[RP_COL] <= end)]

agg_choice = st.sidebar.selectbox("Aggregation", ("mean", "sum"))

if selected_employee != "All":
    df = df[df[EMP_COL].astype(str) == str(selected_employee)]

# ========== SUMMARY SECTION ==========
st.markdown("## 📈 Summary Metrics")

if RP_COL:
    df["_month"] = df[RP_COL].dt.to_period("M").dt.to_timestamp()
    months = sorted(df["_month"].dropna().unique())
    latest_month = months[-1] if months else None
    prev_month = months[-2] if len(months) >= 2 else None
else:
    df["_month"] = None
    latest_month = prev_month = None

default_metrics = numeric_cols[:6]
chosen_metrics = st.multiselect("Select metrics to summarize", numeric_cols, default=default_metrics)

def agg_val(data, col):
    return data[col].mean() if agg_choice == "mean" else data[col].sum()

if latest_month is not None:
    latest_df = df[df["_month"] == latest_month]
    prev_df = df[df["_month"] == prev_month] if prev_month is not None else pd.DataFrame(columns=df.columns)

cols = st.columns(len(chosen_metrics) if chosen_metrics else 1)
for i, m in enumerate(chosen_metrics):
    cur_val = agg_val(latest_df, m) if latest_month is not None else agg_val(df, m)
    prev_val = agg_val(prev_df, m) if prev_month is not None else np.nan
    label = m.replace("_", " ")

    # Determine improvement text
    if not np.isnan(prev_val) and prev_val != 0:
        change = cur_val - prev_val
        improved = change > 0
        change_text = "▲ Improved" if improved else "▼ Declined"
        color = "green" if improved else "red"
        delta_html = f"<span style='color:{color};font-weight:bold'>{change_text}</span>"
    else:
        delta_html = "<span style='color:gray'>No prior data</span>"

    val_display = f"{cur_val:,.2f}" if not pd.isna(cur_val) else "N/A"
    cols[i].markdown(f"<h4>{label}</h4><h2>{val_display}</h2>{delta_html}", unsafe_allow_html=True)

st.markdown("---")

# ========== MAIN TREND CHART ==========
st.markdown("## 📊 Monthly Trends by Metric")

metrics_for_trend = st.multiselect("Select metrics to visualize", numeric_cols, default=chosen_metrics[:2])
if RP_COL and metrics_for_trend:
    df_trend = df.groupby("_month")[metrics_for_trend].agg(agg_choice).reset_index().sort_values("_month")

    for m in metrics_for_trend:
        chart_type = "bar" if any(k in m.lower() for k in ["count", "total", "num", "hour", "score"]) else "line"
        title = f"Trend for {m.replace('_',' ')} ({agg_choice})"

        if chart_type == "bar":
            fig = px.bar(df_trend, x="_month", y=m, text_auto=True, title=title)
        else:
            fig = px.line(df_trend, x="_month", y=m, markers=True, title=title)

        fig.update_layout(xaxis_title="Month", yaxis_title=m, hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Please select a metric and ensure Reporting_Period is valid.")

st.markdown("---")

# ========== METRIC DISTRIBUTION ==========
st.markdown("## 📦 Metric Distribution Across Employees")

dist_metric = st.selectbox("Select metric for distribution", numeric_cols)
if dist_metric:
    fig = px.box(df, x=EMP_COL, y=dist_metric, points="all", title=f"Distribution of {dist_metric} across employees")
    fig.update_layout(xaxis={'categoryorder':'total descending'})
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ========== DATA PREVIEW ==========
st.markdown("## 📄 Data Preview & Download")
st.dataframe(df.head(200), use_container_width=True)
csv = df.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download filtered data", csv, file_name=f"{sheet_name}_filtered.csv", mime="text/csv")
