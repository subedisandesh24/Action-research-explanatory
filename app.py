# app.py
import streamlit as st

st.set_page_config(
    page_title="AgriStats Report Generator",
    page_icon="🌾",
    layout="wide"
)

# Sidebar Configuration Control Panel
st.sidebar.markdown("# ⚙️ Control Panel")
design_type = st.sidebar.selectbox(
    "Select Experimental Design",
    [
        "1. RCBD Single Factor (One Year/Season)",
        "2. RCBD Single Factor (2-Year Combined Analysis)",
        "3. RCBD Two-Factor Factorial"
    ]
)

# Route to the appropriate module based on selection
if design_type == "1. RCBD Single Factor (One Year/Season)":
    from modules import rcbd_single_factor_1year
    rcbd_single_factor_1year.show_module()
    
elif design_type == "2. RCBD Single Factor (2-Year Combined Analysis)":
    from modules import rcbd_single_factor_2year
    rcbd_single_factor_2year.show_module()
    
elif design_type == "3. RCBD Two-Factor Factorial":
    from modules import rcbd_two_factor
    rcbd_two_factor.show_module()

# Footer Information
st.sidebar.markdown("---")
st.sidebar.markdown(
    """
    Developed for Q1 Journal Submissions.
    
    Supported Analyses:
    * Summarized Tables (APA Formatting)
    * Raw Replication Data (Linear ANOVA)
    * Compact Letter Displays (CLD)
    """
)
