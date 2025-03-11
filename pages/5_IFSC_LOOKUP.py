import os
import streamlit as st
import requests
import re
import pandas as pd
import time
from io import StringIO, BytesIO

# Ensure xlsxwriter is installed
os.system("pip install xlsxwriter")

# Now import it
import xlsxwriter

# --- Helper Functions ---

def fetch_ifsc_details(ifsc_code):
    """Fetch IFSC details from Razorpay API."""
    url = f"https://ifsc.razorpay.com/{ifsc_code}"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.json()
    except (requests.RequestException, ValueError):
        return None

def extract_pincode(address):
    """Extract the last occurring 6-digit number from the address as the pincode."""
    if not address:
        return ""
    address_without_spaces = address.replace(" ", "")
    pincode_matches = re.findall(r"\d{6}", address_without_spaces)
    return pincode_matches[-1] if pincode_matches else ""

def process_ifsc_codes(ifsc_codes):
    """Process a list of IFSC codes and fetch details for each."""
    results = []
    for code in ifsc_codes:
        code = code.strip()
        if code:
            details = fetch_ifsc_details(code)
            if details:
                details["PINCODE"] = extract_pincode(details.get("ADDRESS", ""))
                # Ensure expected keys are present
                details.setdefault("BANK", "")
                details.setdefault("BRANCH", "")
                details.setdefault("CITY", "")
                details.setdefault("DISTRICT", "")
                details.setdefault("STATE", "")
                details["ERROR"] = ""
                results.append(details)
            else:
                results.append({
                    "BANK": "",
                    "IFSC": code,
                    "BRANCH": "",
                    "ADDRESS": "",
                    "CITY": "",
                    "DISTRICT": "",
                    "STATE": "",
                    "PINCODE": "",
                    "ERROR": "Invalid IFSC Code"
                })
    return results

# --- Streamlit App UI ---

st.set_page_config(page_title="IFSC Code Lookup", layout="wide")
st.title("IFSC Code Lookup")
st.markdown("Created using [Razorpay API](https://ifsc.razorpay.com/)")

# Use a form for input so that the whole page doesn’t rerun with every change
with st.form(key='ifsc_form'):
    ifsc_input = st.text_area("Enter IFSC Codes (comma‐separated)",
                              placeholder="E.g., SBIN0005943, HDFC0000123, ICIC0000001")
    remove_duplicates = st.checkbox("Exclude duplicate results", value=False)
    submitted = st.form_submit_button("Search")

if submitted:
    if not ifsc_input.strip():
        st.error("Please enter at least one IFSC code.")
    else:
        # Split the input string into individual codes
        codes = [code.strip() for code in ifsc_input.split(",") if code.strip()]
        if remove_duplicates:
            # Remove duplicates while preserving order
            codes = list(dict.fromkeys(codes))
        
        total = len(codes)
        st.info(f"Processing {total} IFSC code{'s' if total > 1 else ''}...")

        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, code in enumerate(codes):
            status_text.text(f"Processing {i+1} of {total}...")
            data = fetch_ifsc_details(code)
            if data:
                data["PINCODE"] = extract_pincode(data.get("ADDRESS", ""))
                data.setdefault("BANK", "")
                data.setdefault("BRANCH", "")
                data.setdefault("CITY", "")
                data.setdefault("DISTRICT", "")
                data.setdefault("STATE", "")
                data["ERROR"] = ""
                results.append(data)
            else:
                results.append({
                    "BANK": "",
                    "IFSC": code,
                    "BRANCH": "",
                    "ADDRESS": "",
                    "CITY": "",
                    "DISTRICT": "",
                    "STATE": "",
                    "PINCODE": "",
                    "ERROR": "Invalid IFSC Code"
                })
            progress_bar.progress(int(((i + 1) / total) * 100))
            time.sleep(0.1)  # small delay for visual effect

        status_text.text("Processing complete.")

        # Convert results to a DataFrame with a fixed column order
        df = pd.DataFrame(results)
        expected_cols = ["BANK", "IFSC", "BRANCH", "ADDRESS", "CITY", "DISTRICT", "STATE", "PINCODE", "ERROR"]
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ""
        df = df[expected_cols]

        st.subheader("Bank Details")
        st.dataframe(df, use_container_width=True)

        # --- Export Options ---

        # CSV export
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        st.download_button(label="Export to CSV",
                           data=csv_data,
                           file_name="ifsc_details.csv",
                           mime="text/csv")

        # Excel export using an in-memory bytes buffer
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="IFSC Details")
        st.download_button(label="Export to Excel",
                           data=excel_buffer.getvalue(),
                           file_name="ifsc_details.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Provide tab-separated text for copy-paste
        tsv_text = df.to_csv(sep="\t", index=False)
        st.text_area("Copy to Clipboard", value=tsv_text, height=200,
                     help="Select all and copy to clipboard")
