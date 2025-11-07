import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

st.set_page_config(page_title="Employee KPI Dashboard v2", layout="wide", initial_sidebar_state="expanded")
st.markdown("<h1 style='text-align:left'>📊 Employee KPI Dashboard</h1>", unsafe_allow_html=True)
st.write("This app auto-loads **Updated_18_KPI_Dashboard.xlsx** from the repository and assumes each sheet has the same structure (Employee_ID, Reporting_Period, KPI columns).")

@st.cache_data
def load_all_sheets(path="Updated_18_KPI_Dashboard.xlsx"):
    # Read all sheets into a dict of DataFrames
    xls = pd.ExcelFile(path)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        # Normalize column names (strip)
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        # Try parse Reporting_Period to datetime if exists
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

# Sidebar: sheet selector and global filters
st.sidebar.header("Dataset & Filters")
sheet_name = st.sidebar.selectbox("Select sheet (department/team)", list(sheets.keys()))
df = sheets[sheet_name].copy()

# Identify special columns
EMP_COL = "Employee_ID" if "Employee_ID" in df.columns else next((c for c in df.columns if "employee" in str(c).lower()), df.columns[0])
RP_COL = next((c for c in df.columns if "report" in str(c).lower() or "period" in str(c).lower() or "date" in str(c).lower()), None)

# Ensure Reporting_Period is datetime-like
if RP_COL is not None and not pd.api.types.is_datetime64_any_dtype(df[RP_COL]):
    try:
        df[RP_COL] = pd.to_datetime(df[RP_COL], errors="coerce")
    except Exception:
        pass

# Numeric KPI columns (exclude EMP_COL and RP_COL)
numeric_cols = df.select_dtypes(include="number").columns.tolist()
numeric_cols = [c for c in numeric_cols if c not in [EMP_COL, RP_COL]]

# If no numeric_cols found, try to coerce some columns to numeric
if not numeric_cols:
    for c in df.columns:
        if c not in [EMP_COL, RP_COL]:
            coerced = pd.to_numeric(df[c], errors="coerce")
            if coerced.notna().sum() > 0:
                df[c] = coerced
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    numeric_cols = [c for c in numeric_cols if c not in [EMP_COL, RP_COL]]

# Sidebar: employee selector
employees = ["All"] + sorted(df[EMP_COL].dropna().astype(str).unique().tolist())
selected_employee = st.sidebar.selectbox("Select Employee", employees, index=0)

# Sidebar: date range if RP_COL exists
if RP_COL and pd.api.types.is_datetime64_any_dtype(df[RP_COL]):
    min_date = df[RP_COL].min().date()
    max_date = df[RP_COL].max().date()
    date_range = st.sidebar.date_input("Reporting period range", value=(min_date, max_date))
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        df = df[(df[RP_COL] >= start) & (df[RP_COL] <= end)]
else:
    date_range = None

# Sidebar: aggregation choice for summary (mean or sum)
agg_choice = st.sidebar.selectbox("Aggregation for KPI summaries", ("mean", "sum"))

# Filter by employee if selected
if selected_employee != "All":
    df = df[df[EMP_COL].astype(str) == str(selected_employee)]

st.markdown("## Summary")

# Compute month period (YYYY-MM) if RP_COL exists
if RP_COL and pd.api.types.is_datetime64_any_dtype(df[RP_COL]):
    df["_month"] = df[RP_COL].dt.to_period("M").dt.to_timestamp()
    # Get latest month and previous month in the filtered data
    months = sorted(df["_month"].dropna().unique())
    if months:
        latest_month = months[-1]
        prev_month = months[-2] if len(months) >= 2 else None
    else:
        latest_month = prev_month = None
else:
    df["_month"] = None
    latest_month = prev_month = None

# Choose which metrics to include in summary (employee-related)
default_metrics = numeric_cols[:6]  # show up to 6 by default
chosen_metrics = st.multiselect("Choose metrics to appear in summary", numeric_cols, default=default_metrics)

# Helper aggregation
def agg_series(data, metric, agg):
    if agg == "mean":
        return data[metric].mean(skipna=True)
    else:
        return data[metric].sum(skipna=True)

# Build summary cards: current vs previous month
cards = []
if latest_month is not None:
    latest_df = df[df["_month"] == latest_month]
    prev_df = df[df["_month"] == prev_month] if prev_month is not None else pd.DataFrame(columns=df.columns)
    cols = st.columns(max(1, len(chosen_metrics)))
    for i, m in enumerate(chosen_metrics):
        cur_val = agg_series(latest_df, m, agg_choice) if not latest_df.empty else np.nan
        prev_val = agg_series(prev_df, m, agg_choice) if not prev_df.empty else np.nan
        # compute percent change safely
        if pd.isna(prev_val) or prev_val == 0 or pd.isna(cur_val):
            pct_change = None
        else:
            pct_change = (cur_val - prev_val) / abs(prev_val) * 100
        # choose delta display and color
        if pct_change is None:
            delta_display = "N/A"
            delta_color = None
        else:
            arrow = "▲" if pct_change > 0 else "▼" if pct_change < 0 else "—"
            delta_display = f"{arrow} {abs(pct_change):.1f}%"
            delta_color = "green" if pct_change > 0 else "red" if pct_change < 0 else None
        label = f"{m.replace('_',' ')} ({agg_choice})"
        val_display = f"{cur_val:,.2f}" if not pd.isna(cur_val) else "N/A"
        # Use st.metric with delta if available
        try:
            cols[i].metric(label, val_display, delta_display)
        except Exception:
            cols[i].write(f"**{label}**: {val_display}  {delta_display}")
else:
    st.info("No Reporting_Period detected or no monthly data available — summary cards will show overall aggregates.")
    cols = st.columns(max(1, len(chosen_metrics)))
    for i, m in enumerate(chosen_metrics):
        overall = agg_series(df, m, agg_choice)
        cols[i].metric(m.replace('_',' '), f"{overall:,.2f}")

st.markdown("---")

# First main chart: Monthly trends for selected metrics (more explanatory)
st.markdown("## Monthly Trends")
metrics_for_trend = st.multiselect("Select metrics to plot over time (multi-select)", numeric_cols, default=chosen_metrics[:2] if chosen_metrics else numeric_cols[:2])
if RP_COL and pd.api.types.is_datetime64_any_dtype(df[RP_COL]) and metrics_for_trend:
    # Aggregate per month
    trend_df = df.groupby("_month")[metrics_for_trend].agg(agg_choice).reset_index().sort_values("_month")
    if trend_df.empty:
        st.warning("No monthly data available for selected filters.")
    else:
        # Melt for plotting multiple lines
        trend_melt = trend_df.melt(id_vars=["_month"], value_vars=metrics_for_trend, var_name="Metric", value_name="Value")
        fig = px.line(trend_melt, x="_month", y="Value", color="Metric", markers=True,
                      title="Monthly trend (aggregated) — hover to see values and month-to-month change")
        # Add percent change annotations between months for the first metric selected to make it easy to understand
        first_metric = metrics_for_trend[0]
        if first_metric in trend_df.columns and len(trend_df) >= 2:
            # compute pct change column
            pct = trend_df[first_metric].pct_change()*100
            # show as secondary bar chart under the lines using a secondary y-axis
            fig2 = go.Figure()
            for m in metrics_for_trend:
                fig2.add_trace(go.Scatter(x=trend_df["_month"], y=trend_df[m], mode="lines+markers", name=str(m)))
            fig2.update_layout(title=f"Monthly trends (primary) and % change for {first_metric} (secondary)",
                               xaxis_title="Month", yaxis_title="Value", hovermode="x unified")
            # Add bar for percent change scaled into secondary y-axis
            fig2.add_trace(go.Bar(x=trend_df["_month"], y=pct, name=f"% change {first_metric}", yaxis="y2", opacity=0.4))
            # Create secondary y axis
            fig2.update_layout(yaxis2=dict(title="% change", overlaying="y", side="right", showgrid=False))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Monthly trend requires Reporting_Period to be parsed as dates and at least one metric selected.")

st.markdown("---")

# Employee comparison chart (top N)
st.markdown("## Employee Comparison (Top N)")
comp_metric = st.selectbox("Metric for comparison across employees", numeric_cols, index=0 if numeric_cols else None)
top_n = st.slider("Top N employees to show", min_value=3, max_value=min(50, max(3, int(df[EMP_COL].nunique()))), value=min(10, max(3, int(df[EMP_COL].nunique()))))
if comp_metric:
    comp_df = df.groupby(EMP_COL)[comp_metric].agg(agg_choice).reset_index().sort_values(comp_metric, ascending=False).head(top_n)
    fig = px.bar(comp_df, x=EMP_COL, y=comp_metric, text=comp_metric, title=f"Top {top_n} employees by {comp_metric} ({agg_choice})")
    fig.update_layout(xaxis={'categoryorder':'total descending'})
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Correlation heatmap for numeric employee metrics
st.markdown("## Correlation Matrix (numeric metrics)")
if len(numeric_cols) >= 2:
    corr_df = df[numeric_cols].corr()
    fig = px.imshow(corr_df, text_auto=True, aspect="auto", title="Correlation between numeric KPIs")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.write("Not enough numeric KPI columns for correlation.")

st.markdown("---")

# Data preview and download
st.markdown("## Data Preview & Download")
st.dataframe(df.reset_index(drop=True).head(300), use_container_width=True)
csv = df.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download filtered data (CSV)", csv, file_name=f"{sheet_name}_filtered_data.csv", mime="text/csv")
