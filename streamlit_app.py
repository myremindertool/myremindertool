import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime

# Clean and process extracted data
def clean_data(df):
    if df.empty:
        return df

    df = df[df.iloc[:, 0] != df.columns[0]]  # Remove repeated headers
    df = df.replace(r'\*', '', regex=True)  # Remove '*'
    df.columns = [re.sub(r'[^a-zA-Z ]', '', str(col)) for col in df.columns]  # Clean headers

    if 'Card Number' in df.columns:
        df['Expiry Date'] = df['Card Number'].str.extract(r'(\d{2}/\d{2}/\d{4})')
        df['Card Number'] = df['Card Number'].str.replace(r'(\d{2}/\d{2}/\d{4})', '', regex=True)

    if 'Person Name' in df.columns:
        df['Person Name'] = df['Person Name'].str.replace(r'[^a-zA-Z ]+', '', regex=True)

    for col in df.columns:
        if col != 'Expiry Date':
            df[col] = df[col].str.replace('/', '', regex=False)

    df = df.map(lambda x: re.sub(r'[^a-zA-Z0-9 /]+', '', str(x)).replace('\n', ' ') if pd.notnull(x) else x)
    df = df.dropna(how='all')
    return df

# Extract tables from uploaded PDF
def extract_tables(pdf_file):
    all_data = []
    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        progress = st.progress(0)
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                all_data.extend(table)
            progress.progress((i + 1) / total_pages)

    if not all_data:
        st.warning("No tables found in the PDF.")
        return pd.DataFrame()

    df = pd.DataFrame(all_data)
    df.columns = df.iloc[0]
    df = df[1:]
    return clean_data(df)

# Streamlit UI
st.set_page_config(page_title="MOL PDF to Excel", layout="centered")
st.title("📄 MOL PDF to Excel Converter")
st.write("Developed by JP - Upload your MOL PDF and download cleaned Excel data.")

uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Extracting and cleaning data..."):
        extracted_data = extract_tables(uploaded_file)

    if not extracted_data.empty:
        st.success("Data extracted successfully!")
        st.dataframe(extracted_data.head())

        # Download button for Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            extracted_data.to_excel(writer, index=False, sheet_name='MOL Data')
        buffer.seek(0)

        st.download_button(
            label="📥 Download Excel File",
            data=buffer,
            file_name=f"MOL_Extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No valid data extracted from PDF.")
