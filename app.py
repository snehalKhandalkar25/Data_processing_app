import pandas as pd
import re
from datetime import datetime
import streamlit as st

# Streamlit app title
st.title("Excel Data Processing and Display")

# File uploader to select the Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Load the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file)

    # Step 1: Ensure 'Emp Name' column has three or more letters
    df = df[df['Emp Name'].apply(lambda x: isinstance(x, str) and len(x.strip()) >= 3)]

    # Step 2: Check 'DOJ' column for proper date format and convert to date only
    def check_date_format(date_value):
        if pd.isna(date_value):
            return False
        if isinstance(date_value, datetime):
            return True
        try:
            # Try to convert to datetime with common formats
            pd.to_datetime(date_value, errors='raise', format='%Y-%m-%d')
            return True
        except (ValueError, TypeError):
            try:
                # Check for other common formats with dayfirst
                pd.to_datetime(date_value, errors='raise', dayfirst=True)
                return True
            except (ValueError, TypeError):
                return False

    # Filter rows with proper date in 'DOJ'
    df = df[df['DOJ'].apply(check_date_format)]

    # Convert 'DOJ' to date format and keep only date (YYYY-MM-DD)
    df['DOJ'] = pd.to_datetime(df['DOJ']).dt.date

    # Step 3: Replace blank cells in 'OP balance' with zero
    df['Op balance'] = df['Op balance'].fillna(0)

    # Step 4: Find all columns containing "Leaves Credited" and set all their cells to 2
    columns_with_leaves_credited = [col for col in df.columns if 'Leaves Credited' in col]
    df[columns_with_leaves_credited] = 2

    # Step 5: Remove special characters from all columns
    df = df.applymap(lambda x: re.sub(r'[^A-Za-z0-9\s]+', '', str(x)) if isinstance(x, str) else x)

    # Step 6: Delete the last three columns
    df = df.iloc[:, :-3]

    # Step 7: Replace any empty cells with zero
    df = df.fillna(0)

    # Display the first few rows of the processed DataFrame in Streamlit
    st.subheader("Processed DataFrame")
    st.dataframe(df)

    # Download button for the processed Excel file
    output_file_path = 'updated_' + uploaded_file.name
    df.to_excel(output_file_path, index=False)
    with open(output_file_path, "rb") as file:
        btn = st.download_button(
            label="Download Processed Excel File",
            data=file,
            file_name=output_file_path,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload an Excel file to process.")
