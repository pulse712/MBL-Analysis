import streamlit as st

st.set_page_config(page_title="MLB Daily Betting Report", page_icon="⚾", layout="wide")

pg = st.navigation([
    st.Page("pages/1_Daily_Report.py",   title="Daily Report",    icon="⚾"),
    st.Page("pages/2_Scatter_Analysis.py", title="Scatter Analysis", icon="📊"),
])
pg.run()
