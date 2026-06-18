import streamlit as st
st.write("Testing imports...")

try:
    from excel_fixer import fix_excel, get_parking_codes_from_pnl
    st.success("excel_fixer imports OK")
except Exception as e:
    st.error(f"excel_fixer failed: {e}")
