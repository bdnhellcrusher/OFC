# pdf_to_excel_converter.py
import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from datetime import datetime, timedelta

def extract_table_data_from_text(text, current_date):
    data = []

    lines = text.split('\n')
    st.write("Text extracted from PDF:")
    st.write(lines)

    for line in lines:
        line = line.strip()
        if not line:
            continue

        date_match = re.match(r'^\d{2}/\d{2}/\d{4}', line)
        if date_match:
            current_date = date_match.group(0)
            current_date = pd.to_datetime(current_date, format='%d/%m/%Y').strftime('%d/%m/%Y')
            continue

        if not current_date:
            continue

        user_id_match = re.search(r'\b[A-Za-z0-9]{4,}\b', line)
        punch_time_match = re.search(r'\d{2}:\d{2}:\d{2}', line)
        io_type_match = re.search(r'\bIN\b|\bOUT\b', line)

        user_id = user_id_match.group(0).strip() if user_id_match else ''
        punch_time = punch_time_match.group(0).strip() if punch_time_match else ''
        io_type = io_type_match.group(0).strip() if io_type_match else ''

        if user_id:
            name_start = line.find(user_id) + len(user_id)
            name_end = line.find(punch_time) if punch_time else len(line)
            name = line[name_start:name_end].strip()
            name = re.sub(r'\bIN\b|\bOUT\b', '', name).strip()

            if punch_time:
                data.append([current_date, user_id, name, punch_time, io_type])

    return data, current_date

def pdf_to_excel(pdf_file):
    all_data = []
    current_date = None

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            st.write(f"Text from page {page.page_number}:")
            st.write(text)

            if text:
                page_data, current_date = extract_table_data_from_text(text, current_date)
                if page_data:
                    all_data.extend(page_data)
            else:
                st.write("No text extracted from page.")

    if all_data:
        result_df = pd.DataFrame(all_data, columns=['Date', 'User ID', 'Name', 'Punch Time', 'I/O Type'])
        return result_df
    else:
        return pd.DataFrame()

