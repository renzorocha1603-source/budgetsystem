import streamlit as st
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from mistralai import Mistral
import io
import re

st.set_page_config(page_title="Budget System", layout="wide")

# Mistral setup
MISTRAL_API_KEY = st.secrets.get("MISTRAL_API_KEY", "")
client = Mistral(api_key=MISTRAL_API_KEY) if MISTRAL_API_KEY else None

st.title("📊 Budget System - Excel Automation")
st.markdown("Upload Excel template + PDF report → Auto-fix Excel")

# File uploads
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Excel Template (.xlsx)", type=["xlsx"])
with col2:
    pdf_file = st.file_uploader("PDF Report (.pdf)", type=["pdf"])

if excel_file and pdf_file:
    # Extract from PDF
    with pdfplumber.open(pdf_file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()
    
    # Find numbers
    numbers = re.findall(r'\$?([\d,]+\.?\d*)', text)
    total_revenue = float(numbers[0].replace(',', '')) if numbers else 0
    
    st.subheader("📊 Extracted Data")
    st.metric("Total Revenue", f"${total_revenue:,.2f}")
    
    # Configuration
    col1, col2, col3 = st.columns(3)
    with col1:
        parking_name = st.selectbox("Parking", ["CMO142", "CMO143", "CMO144"])
    with col2:
        parking_type = st.selectbox("Type", ["SC (School)", "RG (Tourism)"])
    with col3:
        supervisor_hours = st.number_input("Supervisor Hours/Day", 0, 24, 1)
    
    if st.button("🔧 Fix Excel Template", type="primary"):
        # Load and modify Excel
        wb = load_workbook(io.BytesIO(excel_file.read()))
        ws = wb.active
        
        # Update cells
        ws['K26'] = total_revenue
        
        # Seasonal adjustments
        multipliers = [1.2, 1.2, 1.2, 1.1, 0.8, 0.8, 0.8, 0.8, 1.1, 1.1, 1.2, 1.2]
        base = total_revenue / 12
        for i, m in enumerate(multipliers[:12]):
            try:
                ws[f'K{20+i}'] = base * m
            except:
                pass
        
        # Save to bytes
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("✅ Excel file fixed successfully!")
        st.download_button(
            label="📥 Download Fixed Excel",
            data=output,
            file_name=f"{parking_name}_budget_fixed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Ask Mistral
        if client:
            st.subheader("🤖 Ask Mistral AI")
            question = st.text_input("Question about your budget:")
            if question:
                response = client.chat.complete(
                    model="mistral-small-latest",
                    messages=[{"role": "user", "content": question}]
                )
                st.info(response.choices[0].message.content)
