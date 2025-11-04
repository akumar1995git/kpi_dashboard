import streamlit as st

# Detect the theme mode via JavaScript injected element or user button
# Streamlit does not provide an official API, so use colors that are visible in both

# Neutral background that works well on both modes
st.markdown(
    """
    <style>
    /* Override background colors for app container */
    .stApp, .main, .block-container {
        background-color: #f0f2f6 !important;
        color: #0e1117 !important;
    }

    /* Metric labels and values with both light and dark-friendly colors */
    .stMetricLabel, .stMetricValue {
        color: #0e1117 !important;  /* Dark enough for light mode, visible in dark mode */
    }

    /* Tab buttons colored for visibility in both light/dark mode */
    .stTabs [data-baseweb="tab-list"] button {
        background-color: #1f77b4 !important;  /* softer blue */
        color: white !important;
    }
    .stTabs [data-baseweb="tab-list"] button:focus, 
    .stTabs [data-baseweb="tab-list"] button:hover {
        background-color: #155d8b !important;
        color: white !important;
    }

    /* Line chart background - use white so dark mode charts stand out */
    .plotly-graph-div {
        background-color: white !important;
    }

    /* Sidebar background for uniformity */
    .css-1d391kg {
        background-color: #e3e8f0 !important;
        color: #0e1117 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.write("This styling improves visibility for both dark and light Streamlit themes.")
