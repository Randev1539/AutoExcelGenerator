import streamlit as st
import pandas as pd
import os
import tempfile

def generate_excel_sheet_from_data(column_headers, data):
    try:
        # Create a Pandas DataFrame from the user input
        df = pd.DataFrame([x.split(',') for x in data.split(';')], columns=column_headers.split(','))
    except Exception as e:
        raise ValueError(f"Error processing input data: {e}")
    
    try:
        # Create a temporary directory to store the generated Excel file
        temp_dir = tempfile.TemporaryDirectory()
        file_path = os.path.join(temp_dir.name, 'output.xlsx')

        # Write the DataFrame to the Excel file
        df.to_excel(file_path, index=False)
    except Exception as e:
        raise IOError(f"Error creating Excel file: {e}")

    return file_path, temp_dir

def generate_excel_sheet_from_csv(csv_file):
    try:
        # Read the CSV file into a Pandas DataFrame
        df = pd.read_csv(csv_file)
    except Exception as e:
        raise ValueError(f"Error reading CSV file: {e}")
    
    try:
        # Create a temporary directory to store the generated Excel file
        temp_dir = tempfile.TemporaryDirectory()
        file_path = os.path.join(temp_dir.name, 'output.xlsx')

        # Write the DataFrame to the Excel file
        df.to_excel(file_path, index=False)
    except Exception as e:
        raise IOError(f"Error creating Excel file: {e}")

    return file_path, temp_dir

def main():
    st.title("Generate Excel Sheet")
    st.write("Enter column headers and data separated by commas, or upload a CSV file")

    option = st.selectbox("Choose input method", ("Manual Entry", "Upload CSV"))

    if option == "Manual Entry":
        column_headers = st.text_input("Column Headers (e.g. Name,Age,Address)")
        data = st.text_area("Data (e.g. John,25,New York; Alice,30,London; ...)\n    Separate each row with a semicolon (;)", height=300)

        if st.button("Generate Excel Sheet"):
            if not column_headers or not data:
                st.error("Please provide both column headers and data.")
                return

            try:
                file_path, temp_dir = generate_excel_sheet_from_data(column_headers, data)

                # Display the generated Excel file
                st.write("Generated Excel Sheet:")
                st.markdown(f"**File size:** {os.path.getsize(file_path)/1024:.2f} KB")

                with open(file_path, "rb") as file:
                    st.download_button(
                        label="Download Excel Sheet",
                        data=file,
                        file_name="output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Cleanup: The temporary file will be removed after download
                temp_dir.cleanup()

            except ValueError as ve:
                st.error(f"Data Processing Error: {ve}")
            except IOError as ioe:
                st.error(f"File Creation Error: {ioe}")
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

    elif option == "Upload CSV":
        uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

        if uploaded_file is not None:
            try:
                file_path, temp_dir = generate_excel_sheet_from_csv(uploaded_file)

                # Display the generated Excel file
                st.write("Generated Excel Sheet:")
                st.markdown(f"**File size:** {os.path.getsize(file_path)/1024:.2f} KB")

                with open(file_path, "rb") as file:
                    st.download_button(
                        label="Download Excel Sheet",
                        data=file,
                        file_name="output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Cleanup: The temporary file will be removed after download
                temp_dir.cleanup()

            except ValueError as ve:
                st.error(f"CSV Reading Error: {ve}")
            except IOError as ioe:
                st.error(f"File Creation Error: {ioe}")
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
