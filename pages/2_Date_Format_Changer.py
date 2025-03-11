import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# Function to check and process individual cell values
def process_value(val):
    if isinstance(val, str):
        val_str = val.strip()
        # First, try to match a date-time pattern (dd mm yyyy hh:mm) with any separator
        dt_pattern = r'^\s*(\d{1,2})\D+(\d{1,2})\D+(\d{4})\D+(\d{1,2}):(\d{2})\s*$'
        m_dt = re.match(dt_pattern, val_str)
        if m_dt:
            day, month, year, hour, minute = m_dt.groups()
            try:
                dt_obj = datetime(int(year), int(month), int(day), int(hour), int(minute))
                return dt_obj.strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass
        # Next, try to match a date-only pattern (dd mm yyyy) with any separator
        date_pattern = r'^\s*(\d{1,2})\D+(\d{1,2})\D+(\d{4})\s*$'
        m_date = re.match(date_pattern, val_str)
        if m_date:
            day, month, year = m_date.groups()
            try:
                date_obj = datetime(int(year), int(month), int(day))
                return date_obj.strftime("%Y-%m-%d")
            except Exception:
                pass
    # Return the original value if no date/datetime is recognized
    return val

# Apply the cell processing function to an entire dataframe
def process_dataframe(df):
    return df.applymap(process_value)

def main():
    st.set_page_config(page_title="Excel Date Formatter", layout="wide")
    st.title("Excel Date Formatter")
    st.write("Upload an Excel file. The app will search for dates in the formats 'dd mm yyyy' or 'dd mm yyyy hh:mm' (using any non-digit as a separator) and reformat them to 'yyyy-mm-dd' or 'yyyy-mm-dd hh:mm' respectively. Other cells remain unchanged.")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Read the Excel file (by default, the first sheet is read)
            df = pd.read_excel(uploaded_file)
            st.subheader("Original Data")
            st.dataframe(df)
            
            # Process the dataframe
            df_processed = process_dataframe(df)
            st.subheader("Processed Data")
            st.dataframe(df_processed)
            
            # Save the processed dataframe to an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_processed.to_excel(writer, index=False)
            processed_data = output.getvalue()
            
            st.download_button(
                label="Download Processed Excel File",
                data=processed_data,
                file_name="processed_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error processing the file: {e}")

if __name__ == "__main__":
    main()
