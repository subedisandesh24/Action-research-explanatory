import streamlit as st
import sys
import importlib

# Set page configuration must remain the first Streamlit command in the script
st.set_page_config(
    page_title="RCBD Statistical Analysis Web Portal",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Force Streamlit to hot-reload nested submodule edits from memory
submodules_to_reload = [
    "modules.rcbd_single_factor_1year",
    "modules.rcbd_single_factor_2year",
    "modules.rcbd_two_factor"
]
for mod in submodules_to_reload:
    if mod in sys.modules:
        try:
            importlib.reload(sys.modules[mod])
        except Exception as e:
            pass  # Suppress initial reload errors before full import takes place

# Safe directory-aware imports from the 'modules' folder
try:
    from modules import rcbd_single_factor_1year
    from modules import rcbd_single_factor_2year
    from modules import rcbd_two_factor
except ImportError as e:
    st.error(f"Import Error: Could not load modules. Details: {e}")
    st.stop()

# --- Main Page Layout ---
st.title("RCBD Statistical Analysis Web Portal")
st.write("An analytical suite to process, format, and explain Randomized Complete Block Design (RCBD) experiments.")

# Sidebar Selection
design_type = st.sidebar.selectbox(
    "Select Experimental Design:",
    options=[
        "RCBD Single-Factor (1 Year)",
        "RCBD Single-Factor (2 Years)",
        "RCBD Two-Factor (Factorial)"
    ],
    key="app_design_selection"
)

st.sidebar.markdown("---")
st.sidebar.info(
    "Ensure your input data matches the correct layout. "
    "Check the specific notes inside each module."
)

# --- Routing Controller Logic ---
if design_type == "RCBD Single-Factor (1 Year)":
    if hasattr(rcbd_single_factor_1year, "show_module"):
        try:
            rcbd_single_factor_1year.show_module()
        except Exception as e:
            st.error(f"An error occurred in 1-Year analysis: {e}")
    else:
        st.warning("The single-factor 1-year module has not exposed a 'show_module()' function.")

elif design_type == "RCBD Single-Factor (2 Years)":
    if hasattr(rcbd_single_factor_2year, "show_module"):
        try:
            rcbd_single_factor_2year.show_module()
        except Exception as e:
            st.error(f"An error occurred in 2-Year analysis: {e}")
    else:
        st.warning("The single-factor 2-year module has not exposed a 'show_module()' function.")

elif design_type == "RCBD Two-Factor (Factorial)":
    if hasattr(rcbd_two_factor, "show_module"):
        try:
            rcbd_two_factor.show_module()
        except Exception as e:
            st.error(f"An error occurred in Two-Factor analysis: {e}")
    else:
        st.warning("The two-factor module has not exposed a 'show_module()' function.")
