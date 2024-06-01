import streamlit as st
import pandas as pd
from openpyxl import Workbook
import os

# Create a Streamlit app
st.title("Excel Sheet Generator")

# Get user input for column headings and data
column_headings = st.text_input("Enter column headings (comma separated):")
column_data = st.text_input("Enter column data (comma separated):")

# Create a button to generate the Excel sheet
generate_button = st.button("Generate Excel Sheet")

if generate_button:
    # Create a Pandas dataframe from the user input
    df = pd.DataFrame([column_data.split(",")], columns=column_headings.split(","))

    # Create an Excel file and write the dataframe to it
    wb = Workbook()
    ws = wb.active
    for r in df.values:
        ws.append(r)
    wb.save("generated_file.xlsx")

    # Display the generated Excel file in Streamlit
    with open("generated_file.xlsx", "rb") as f:
        st.download_button("Download Excel Sheet", f, "generated_file.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Remove the generated file from the server
    os.remove("generated_file.xlsx")

    # Check the file size and display an error message if it exceeds 100MB
    file_size = os.path.getsize("generated_file.xlsx")
    if file_size > 100 * 1024 * 1024:
        st.error("Error: File size exceeds 100MB. Please reduce the amount of data.")
        os.remove("generated_file.xlsx")
        st.stop()

# Add a limit to the app so that it doesn't exceed 100MB
@st.cache(ttl=60)  # Cache for 1 minute
def check_file_size():
    if os.path.exists("generated_file.xlsx"):
        file_size = os.path.getsize("generated_file.xlsx")
        if file_size > 100 * 1024 * 1024:
            st.error("Error: File size exceeds 100MB. Please reduce the amount of data.")
            os.remove("generated_file.xlsx")
            st.stop()

check_file_size()