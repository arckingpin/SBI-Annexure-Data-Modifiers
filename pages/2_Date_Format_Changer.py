import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# Function to check and process individual cell values
def process_value(val):
    if isinstance(val, str):
        val_str = val.strip()
        # Match date-time pattern (dd mm yyyy hh:mm) with any separator
        dt_pattern = r'^\s*(\d{1,2})\D+(\d{1,2})\D+(\d{4})\D+(\d{1,2}):(\d{2})\s*$'
        m_dt = re.match(dt_pattern, val_str)
        if m_dt:
            day, month, year, hour, minute = m_dt.groups()
            try:
                dt_obj = datetime(int(year), int(month), int(day), int(hour), int(minute))
                return dt_obj.strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass
        # Match date-only pattern (dd mm yyyy) with any separator
        date_pattern = r'^\s*(\d{1,2})\D+(\d{1,2})\D+(\d{4})\s*$'
        m_date = re.match(date_pattern, val_str)
        if m_date:
            day, month, year = m_date.groups()
            try:
                date_obj = datetime(int(year), int(month), int(day))
                return date_obj.strftime("%Y-%m-%d")
            except Exception:
                pass
    return val  # Return the original value if no date/datetime is recognized

# Apply processing function to entire DataFrame
def process_dataframe(df):
    return df.applymap(process_value)

# Remove time from date-time values in a column
def remove_time_from_column(df, column):
    df[column] = df[column].apply(lambda x: x.split(" ")[0] if isinstance(x, str) and " " in x else x)
    return df

def main():
    st.set_page_config(page_title="Excel Date Formatter", layout="wide")
    st.title("Excel Date Formatter")
    st.write("Upload an Excel file. The app detects dates in the formats 'dd mm yyyy' or 'dd mm yyyy hh:mm' and converts them to 'yyyy-mm-dd' or 'yyyy-mm-dd hh:mm' respectively.")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.subheader("Original Data")
            st.dataframe(df)

            # Process the dataframe
            df_processed = process_dataframe(df)
            st.subheader("Processed Data")
            processed_data_container = st.container()
            with processed_data_container:
                st.dataframe(df_processed)

            # Identify columns containing date and time
            datetime_columns = [col for col in df_processed.columns if df_processed[col].astype(str).str.contains(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", regex=True).any()]

            if datetime_columns:
                st.subheader("Navigate to Date-Time Columns")
                for col in datetime_columns:
                    if st.button(f"Jump to '{col}'"):
                        st.write(f"Displaying column: {col}")
                        st.dataframe(df_processed[[col]])

                # Allow user to remove time from selected column
                st.subheader("Remove Time from Selected Column")
                selected_column = st.selectbox("Select a column to remove time:", datetime_columns)
                if st.button("Remove Time"):
                    df_processed = remove_time_from_column(df_processed, selected_column)
                    st.success(f"Time removed from column: {selected_column}")
                    st.experimental_rerun()

            # Save processed file for download
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
