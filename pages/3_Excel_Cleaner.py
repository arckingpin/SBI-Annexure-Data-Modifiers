import streamlit as st
import openpyxl
from io import BytesIO
import time

st.set_page_config(page_title="Excel Cleaner", layout="centered")

st.title("Excel Cleaner")

# File uploader for Excel file (.xlsx)
uploaded_file = st.file_uploader("Select an Excel file", type=["xlsx"])

# Text input for output file name (without extension)
output_file_name = st.text_input("Enter a name for the processed file", "cleaned_file")

if st.button("Start Cleaning"):
    if uploaded_file is None:
        st.warning("Please upload an Excel file!")
    elif not output_file_name:
        st.warning("Please enter a name for the processed file!")
    else:
        try:
            # Load workbook using openpyxl from the uploaded file (BytesIO)
            workbook = openpyxl.load_workbook(uploaded_file)
            sheet = workbook.active

            # Calculate total number of cells for progress tracking
            total_cells = sheet.max_row * sheet.max_column
            processed_cells = 0

            progress_bar = st.progress(0)
            progress_text = st.empty()

            # Iterate through all cells in the active sheet
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        cell.value = cell.value.strip()  # Remove extra spaces from strings
                    elif cell.value is None:
                        cell.value = ""  # Replace None with empty string

                    processed_cells += 1
                    progress = processed_cells / total_cells
                    progress_bar.progress(int(progress * 100))
                    progress_text.text(f"Progress: {int(progress * 100)}%")

                    # Simulate a small delay (for demonstration)
                    time.sleep(0.001)

            # Save the cleaned workbook to a BytesIO stream
            output = BytesIO()
            workbook.save(output)
            output.seek(0)

            st.success("File cleaned successfully!")

            # Generate an output file name and provide a download button
            original_name = uploaded_file.name.replace(".xlsx", "")
            output_filename = f"{original_name}_{output_file_name}.xlsx"
            st.download_button(
                label="Download Cleaned File",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
