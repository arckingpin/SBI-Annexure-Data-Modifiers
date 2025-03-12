import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# Function to convert text-based date/time values
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
    return val  # Return original value if no date/datetime is recognized

# Apply processing function to entire DataFrame
def process_dataframe(df):
    return df.applymap(process_value)

# Remove time from date-time values in a column and convert to text.
# If a cell is blank, leave it blank.
def remove_time_from_column(df, column):
    def remove_time_value(val):
        # If the cell is NaN or an empty string, leave it unchanged.
        if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
            return ""
        if isinstance(val, str) and " " in val:
            return val.split(" ")[0]
        return val
    if pd.api.types.is_datetime64_any_dtype(df[column]):
        # For datetime64 values, convert using strftime if not NaT, else blank.
        df[column] = df[column].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else "")
    else:
        df[column] = df[column].apply(remove_time_value)
    df[column] = df[column].astype(str)  # Ensure output is text
    return df

def main():
    st.set_page_config(page_title="Excel Date Formatter", layout="wide")
    st.title("Excel Date Formatter")
    st.write("Upload an Excel file. The app detects date-time values and allows you to remove the time part if needed.")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Read the Excel file as text to prevent Excel auto-formatting issues.
            df = pd.read_excel(uploaded_file, dtype=str)
            st.subheader("Original Data")
            st.dataframe(df)

            # Process the data (for text-based date formats) and store in session state.
            if "df_processed" not in st.session_state:
                st.session_state.df_processed = process_dataframe(df)
            
            processed_df = st.session_state.df_processed
            st.subheader("Processed Data")
            st.dataframe(processed_df)

            # Identify columns containing datetime values (either datetime64 or text-based with time)
            datetime_columns = [
                col for col in processed_df.columns 
                if pd.api.types.is_datetime64_any_dtype(processed_df[col]) or
                processed_df[col].astype(str).str.contains(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", regex=True).any()
            ]

            if datetime_columns:
                st.subheader("Navigate to Date-Time Columns")
                for col in datetime_columns:
                    if st.button(f"Jump to '{col}'", key=f"jump_{col}"):
                        st.write(f"Displaying column: {col}")
                        st.dataframe(processed_df[[col]])

                # Option to remove time from a selected date-time column.
                st.subheader("Remove Time from Selected Column")
                selected_column = st.selectbox("Select a column to remove time:", datetime_columns, key="remove_select")
                if st.button("Remove Time", key="remove_button"):
                    st.session_state.df_processed = remove_time_from_column(processed_df.copy(), selected_column)
                    st.success(f"Time successfully removed from column: {selected_column}")
                    st.subheader("Updated Processed Data")
                    st.dataframe(st.session_state.df_processed)

            # Save the processed data to an Excel file with text formatting.
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                text_format = workbook.add_format({'num_format': '@'})  # Force text format for all columns.
                
                df_processed = st.session_state.df_processed
                sheet_name = "Processed Data"
                df_processed.to_excel(writer, index=False, sheet_name=sheet_name)
                worksheet = writer.sheets[sheet_name]
                
                # Apply text format to all columns.
                for col_num, _ in enumerate(df_processed.columns):
                    worksheet.set_column(col_num, col_num, None, text_format)
                    
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