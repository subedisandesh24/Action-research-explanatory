import streamlit as st

# Configure the Streamlit page layout as the absolute first command
st.set_page_config(
    page_title="RCBD Statistical Analysis Web Portal",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Sidebar Navigation Menu
st.sidebar.title("Design Navigation")
analysis_option = st.sidebar.selectbox(
    "Select Experimental Design Type:",
    options=[
        "RCBD Single Factor (1 Year)",
        "RCBD Single Factor (2 Years / Pooled)",
        "RCBD Two Factor"
    ]
)

# Module Import Validation
try:
    from modules import rcbd_single_factor_1year
    from modules import rcbd_single_factor_2year
    from modules import rcbd_two_factor
except ImportError as e:
    st.error(
        f"Critical Import Error: {e}. Please ensure that the 'modules' directory "
        "and its files (__init__.py, rcbd_single_factor_1year.py, rcbd_single_factor_2year.py, "
        "rcbd_two_factor.py) are correctly placed."
    )
    st.stop()

st.title("RCBD Statistical Analysis Web Portal")
st.markdown("---")

# Orchestration and Routing Logic
if analysis_option == "RCBD Single Factor (1 Year)":
    rcbd_single_factor_1year.show_module()

elif analysis_option == "RCBD Single Factor (2 Years / Pooled)":
    rcbd_single_factor_2year.show_module()

elif analysis_option == "RCBD Two Factor":
    rcbd_two_factor.show_module()
