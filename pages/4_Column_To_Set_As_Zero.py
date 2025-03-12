import pandas as pd
import streamlit as st
import base64
from io import BytesIO

def process_excel_step1(file):
    # Read the Excel file as text to avoid auto-formatting issues
    df = pd.read_excel(file, dtype=str)
    
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
            df[col] = "0"  # Reset as text "0" to prevent auto-formatting in Excel
    
    return df

def process_excel_step2(df):
    # Check that all required columns exist before processing
    required_columns = [
        "Miscellaneous charges - Claimed", "Total claimed amount", "Assessed Claim Amount",
        "Hospital Discount", "Miscellaneous charges - Payable", "Miscellaneous charges - Deducted Amount",
        "Deducted Amount"
    ]
    
    if all(col in df.columns for col in required_columns):
        # Fill blank cells in "Deducted Amount" with "0"
        df["Deducted Amount"] = df["Deducted Amount"].fillna("0")
        df["Deducted Amount"].replace("", "0", inplace=True)
        
        # Copy the value from "Total claimed amount" to "Miscellaneous charges - Claimed"
        df["Miscellaneous charges - Claimed"] = df["Total claimed amount"]
        
        # Calculate "Miscellaneous charges - Payable"
        payable = df["Assessed Claim Amount"].astype(float) + df["Hospital Discount"].astype(float)
        # Format so that integers don't show as decimals (e.g., "100" instead of "100.0")
        df["Miscellaneous charges - Payable"] = payable.apply(lambda x: str(int(x)) if x.is_integer() else str(x))
        
        # Copy the value from "Deducted Amount" to "Miscellaneous charges - Deducted Amount"
        df["Miscellaneous charges - Deducted Amount"] = df["Deducted Amount"]
        
        # Perform a validation check: (Claimed - Deducted) should equal Payable
        df["Validation Check"] = (
            df["Miscellaneous charges - Claimed"].astype(float) -
            df["Miscellaneous charges - Deducted Amount"].astype(float)
            == df["Miscellaneous charges - Payable"].astype(float)
        )
        
        return df, df["Validation Check"].all()
    
    return df, False

def get_table_download_link(df, filename="processed_output.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Download Processed File</a>'
    return href

# Streamlit UI setup
st.set_page_config(page_title="Excel Processor", layout="wide")
st.title("📊 Excel Column To Set As Zero")
st.markdown("( For payment Annexure Only )")
st.markdown("Upload an Excel file to reset specific column values to 0 and download the modified file.")

uploaded_file = st.file_uploader("📂 Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing Step 1..."):
        df_step1 = process_excel_step1(uploaded_file)
    
    st.success("✅ Step 1 Complete!")
    st.subheader("🔍 Preview of modified data:")
    st.dataframe(df_step1.head(10))
    
    st.markdown(get_table_download_link(df_step1, "step1_output.xlsx"), unsafe_allow_html=True)
    
    if st.button("Proceed To Step 2"):
        with st.spinner("Processing Step 2..."):
            df_final, validation_success = process_excel_step2(df_step1)
        
        if validation_success:
            st.success("✅ Data Match Success!")
        else:
            st.error("❌ Data Match Failed. Manually crosscheck processed excel.")
        
        st.subheader("🔍 Final Processed Data Preview:")
        st.dataframe(df_final.head(10))
        
        st.markdown(get_table_download_link(df_final, "final_processed_output.xlsx"), unsafe_allow_html=True)
