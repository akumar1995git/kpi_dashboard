import streamlit as st
import pandas as pd
import plotly.express as px

# ----------------------
# CONFIGURATION
# ----------------------
st.set_page_config(page_title="KPI Dashboard", layout="wide")
st.title("📊 KPI Dashboard")

@st.cache_data
def load_data(uploaded_file=None):
    if uploaded_file is not None:
        file_name = uploaded_file.name.lower()
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)
    else:
        # Default: load from local Excel/CSV file in the same directory
        try:
            df = pd.read_excel("Updated_18_KPI_Metrics.xlsx")
        except Exception:
            df = pd.read_csv("Updated_18_KPI_Metrics.csv")
    
    # Try to parse potential date columns
    for col in df.columns:
        try:
            df[col] = pd.to_datetime(df[col], errors='ignore')
        except Exception:
            pass
    return df

# File uploader
uploaded_file = st.sidebar.file_uploader("📁 Upload your KPI Excel or CSV file", type=["csv", "xlsx", "xls"])
df = load_data(uploaded_file)

# Identify column types
numeric_cols = df.select_dtypes(include='number').columns.tolist()
cat_cols = df.select_dtypes(exclude='number').columns.tolist()

# Sidebar Filters
st.sidebar.header("🔍 Filters")
date_col = None
for c in df.columns:
    if 'date' in c.lower() or 'period' in c.lower():
        date_col = c
        break

if date_col and pd.api.types.is_datetime64_any_dtype(df[date_col]):
    min_date, max_date = df[date_col].min(), df[date_col].max()
    date_range = st.sidebar.date_input("Select Date Range", [min_date, max_date])
    if isinstance(date_range, list) and len(date_range) == 2:
        df = df[(df[date_col] >= pd.to_datetime(date_range[0])) & (df[date_col] <= pd.to_datetime(date_range[1]))]

# Category filter
if cat_cols:
    category_col = st.sidebar.selectbox("Select Category Column", cat_cols)
    selected_category = st.sidebar.multiselect(
        "Filter by Category", df[category_col].dropna().unique(),
        default=df[category_col].dropna().unique()[:5]
    )
    df = df[df[category_col].isin(selected_category)]

# ----------------------
# KPI METRICS
# ----------------------
st.subheader("📈 Key Metrics")
if numeric_cols:
    top_kpis = numeric_cols[:4]
    kpi_cols = st.columns(len(top_kpis))
    for i, k in enumerate(top_kpis):
        value = df[k].mean()
        delta = df[k].iloc[-1] - df[k].iloc[0] if len(df) > 1 else 0
        kpi_cols[i].metric(label=k, value=f"{value:,.2f}", delta=f"{delta:,.2f}")

# ----------------------
# TREND CHART
# ----------------------
if date_col and numeric_cols:
    st.subheader("📉 KPI Trend Over Time")
    metric_to_plot = st.selectbox("Select Metric", numeric_cols)
    fig = px.line(
        df, x=date_col, y=metric_to_plot,
        color=category_col if cat_cols else None,
        markers=True, title=f"{metric_to_plot} Over Time"
    )
    fig.update_layout(hovermode="x unified", template="plotly_white")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------
# CORRELATION HEATMAP
# ----------------------
if len(numeric_cols) > 2:
    st.subheader("📊 Correlation Heatmap")
    corr = df[numeric_cols].corr()
    fig = px.imshow(corr, text_auto=True, color_continuous_scale="RdBu_r", title="Correlation Matrix")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------
# CATEGORY COMPOSITION
# ----------------------
if cat_cols and numeric_cols:
    st.subheader("🧩 Composition by Category")
    metric_for_pie = st.selectbox("Select Metric for Pie Chart", numeric_cols)
    category_col = st.selectbox("Select Category for Pie Chart", cat_cols)
    pie_data = df.groupby(category_col)[metric_for_pie].mean().reset_index()
    fig = px.pie(pie_data, values=metric_for_pie, names=category_col, title=f"{metric_for_pie} by {category_col}")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------
# DATA TABLE
# ----------------------
st.subheader("📋 Data Preview")
st.dataframe(df.head(100))

# Download filtered data
csv = df.to_csv(index=False).encode('utf-8')
st.download_button("⬇️ Download Filtered Data", csv, "filtered_data.csv", "text/csv")
