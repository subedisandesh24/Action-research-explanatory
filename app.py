import streamlit as st

# Set page configuration at the very top
st.set_page_config(
    page_title="RCBD Statistical Analysis Web Portal",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Safe relative imports from the nested 'modules' directory
try:
    from modules import rcbd_single_factor_1year
    from modules import rcbd_single_factor_2year
    from modules import rcbd_two_factor
except ImportError as e:
    st.error(f"Import Error: Could not load modules. Details: {e}")
    st.stop()

# --- Main Application Layout ---
st.title("RCBD Statistical Analysis Web Portal")
st.write("An analytical suite to process, format, and explain Randomized Complete Block Design (RCBD) experiments.")

# Sidebar Selection - Corrected 'choices' to 'options'
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
    try:
        rcbd_single_factor_1year.show_module()
    except AttributeError:
        st.warning("The single-factor 1-year module has not exposed a 'show_module()' function.")
    except Exception as e:
        st.error(f"An error occurred in 1-Year analysis: {e}")

elif design_type == "RCBD Single-Factor (2 Years)":
    try:
        rcbd_single_factor_2year.show_module()
    except AttributeError:
        st.warning("The single-factor 2-year module has not exposed a 'show_module()' function.")
    except Exception as e:
        st.error(f"An error occurred in 2-Year analysis: {e}")

elif design_type == "RCBD Two-Factor (Factorial)":
    try:
        rcbd_two_factor.show_module()
    except AttributeError:
        st.warning("The two-factor module has not exposed a 'show_module()' function.")
    except Exception as e:
        st.error(f"An error occurred in Two-Factor analysis: {e}")
