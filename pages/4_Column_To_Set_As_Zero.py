import pandas as pd
import streamlit as st
import base64
from io import BytesIO

def process_excel(file):
    df = pd.read_excel(file)
    
    columns_to_zero = [
        "Number of days stay in ICU", "ICU per day charges",
        "Hospital Bill - Accommodation Charges (ICU) - Claimed", "Hospital Bill - Accommodation Charges (ICU) - Payable",
        "ICU deduction amount", "Number of days stay in Non ICU", "Non ICU per day charges",
        "Accommodation Charges (Non ICU) - Claimed", "Accommodation Charges (Non ICU) - Payable",
        "Non ICU deduction amount", "Consultation Charges - Claimed", "Consultation Charges - Payable",
        "Consultation deducted amount", "Surgeon Charges - Claimed", "Surgeon Charges - Payable",
        "Surgeon Charges - Deducted amount", "Operation Theatre Charges - Claimed",
        "Operation Theatre Charges - Payable", "Operation Theatre Charges - Deducted amount",
        "Anesthetist Charges - Claimed", "Anesthetist Charges - Payable", "Anesthetist Charges - Deduction amount",
        "Anesthesia Charge - Claimed", "Anesthesia Charge - Payable", "Anesthesia Charge - Deduction amount",
        "Ward Consumables Charges - claimed", "Ward Consumables Charges - Payable", "Ward Consumables Charges - Deduction amount",
        "Medicine Charges - Claimed", "Medicine Charges - Payable", "Medicine Charges - Deducted amount",
        "Investigation charges - Claimed", "Investigation charges - Payable", "Investigation charges - Deducted amount",
        "Reg-Service charges - Claimed", "Reg-Service charges - Payable", "Reg-Service charges - Deducted amount",
        "Ambulance Charges - Claimed", "Ambulance Charges - Payable", "Ambulance Charges - Deduction Amount",
        "Total Pre hospitalization charges", "Total Post hospitalization charges"
    ]
    
    for col in columns_to_zero:
        if col in df.columns:
            df[col] = 0
    
    return df

def get_table_download_link(df, filename="processed_output.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Download Processed File</a>'
    return href

# Streamlit UI Enhancements
st.set_page_config(page_title="Excel Processor", layout="wide")
st.markdown("""
    <style>
        .main { background-color: #f5f5f5; }
        .stButton>button { border-radius: 8px; background-color: #007bff; color: white; }
        .stFileUploader { border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Excel Column To Set As Zero")
st.markdown("( For payment Annexure Only )")
st.markdown("Upload an Excel file to reset specific column values to 0 and download the modified file.")

uploaded_file = st.file_uploader("📂 Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing file..."):
        df_processed = process_excel(uploaded_file)
    
    st.success("✅ Processing complete!")
    st.subheader("🔍 Preview of modified data:")
    st.dataframe(df_processed.head(10))
    
    st.markdown(get_table_download_link(df_processed), unsafe_allow_html=True)
