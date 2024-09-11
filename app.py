# app.py
import streamlit as st
import pandas as pd
from pdf_to_excel_converter import pdf_to_excel
from morning_shift_calculator import process_all_sheets as process_morning_shifts
from night_shift_calculator import process_all_sheets as process_night_shifts
from io import BytesIO

def process_and_organize_data(excel_data):
    """Process and organize the Excel data"""
    # Load the Excel data into a DataFrame
    df = pd.read_excel(BytesIO(excel_data), sheet_name='Sheet1')
    
    # Organize the data
    date_col, punch_time_col, io_type_col, user_id_col, name_col = identify_columns(df)

    if all([date_col, punch_time_col, io_type_col, user_id_col, name_col]):
        organized_df = process_shift_data(df, date_col, punch_time_col, io_type_col, user_id_col, name_col)
        return organized_df
    else:
        st.error("The uploaded file does not contain the required columns.")
        return pd.DataFrame()

def identify_columns(df):
    """Identify relevant columns in the DataFrame"""
    import re
    date_col = next((col for col in df.columns if re.search(r'\bdate\b', col, re.IGNORECASE)), None)
    punch_time_col = next((col for col in df.columns if re.search(r'\bpunch\s*time\b', col, re.IGNORECASE)), None)
    io_type_col = next((col for col in df.columns if re.search(r'\bi\s*/\s*o\s*type\b', col, re.IGNORECASE)), None)
    user_id_col = next((col for col in df.columns if re.search(r'\buser\s*id\b', col, re.IGNORECASE)), None)
    name_col = next((col for col in df.columns if re.search(r'\bname\b', col, re.IGNORECASE)), None)
    return date_col, punch_time_col, io_type_col, user_id_col, name_col

def process_shift_data(df, date_col, punch_time_col, io_type_col, user_id_col, name_col):
    """Process shift data and return organized DataFrame"""
    from datetime import datetime, timedelta

    df['DateTime'] = pd.to_datetime(df[date_col].astype(str) + ' ' + df[punch_time_col].astype(str), format="%d/%m/%Y %H:%M:%S", errors='coerce')
    df.dropna(subset=['DateTime'], inplace=True)

    evening_start_time = datetime.strptime('17:00:00', '%H:%M:%S').time()
    night_end_time = datetime.strptime('02:15:00', '%H:%M:%S').time()

    all_data = []
    df = df.sort_values(by=[user_id_col, 'DateTime'])

    for user, user_df in df.groupby(user_id_col):
        user_data = []
        current_shift = []
        previous_logout_date = None

        for _, row in user_df.iterrows():
            current_time = row['DateTime'].time()
            current_date = row['DateTime'].date()

            if current_time >= evening_start_time:
                if previous_logout_date and current_date > previous_logout_date:
                    next_day_data = user_df[user_df['DateTime'].dt.date == current_date]
                    if not next_day_data.empty:
                        previous_row = next_day_data.iloc[0]
                        current_shift.append(previous_row)
                    previous_logout_date = None

                if current_time <= night_end_time:
                    current_date += timedelta(days=1)
                row[date_col] = current_date.strftime('%d/%m/%Y')
                current_shift.append(row)

            else:
                if current_shift:
                    user_data.append(current_shift)
                current_shift = [row]

            if row[io_type_col] == 'OUT':
                previous_logout_date = row['DateTime'].date()

        if current_shift:
            user_data.append(current_shift)

        for shift in user_data:
            shift_start_date = shift[0]['DateTime'].date()
            shift_end_date = shift[-1]['DateTime'].date()

            for i, record in enumerate(shift):
                if i > 0 and shift[i-1][io_type_col] == 'OUT' and record[io_type_col] == 'IN':
                    if shift[i-1]['DateTime'].date() != record['DateTime'].date():
                        shift_end_date = record['DateTime'].date()

                shift_date = shift_start_date if record[date_col] == shift[0][date_col] else shift_end_date

                all_data.append({
                    'Date': shift_date.strftime('%d/%m/%Y'),
                    'User ID': user,
                    'Name': record[name_col],
                    'Punch Time': record[punch_time_col],
                    'I/O Type': record[io_type_col],
                    'Shift Start': shift_start_date,
                    'Shift End': shift_end_date
                })

    final_df = pd.DataFrame(all_data)
    return final_df

def main():
    st.title("Employee Shift Calculation Application")

    st.sidebar.header("Upload Your File")
    uploaded_file = st.sidebar.file_uploader("Upload PDF file", type=["pdf"])

    if uploaded_file:
        st.sidebar.write("Processing PDF...")
        data_df = pdf_to_excel(uploaded_file)
        
        if not data_df.empty:
            # Convert DataFrame to Excel and store in session_state
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                data_df.to_excel(writer, sheet_name='Sheet1', index=False)
            excel_buffer.seek(0)
            st.session_state.excel_data = excel_buffer.getvalue()
            st.write("PDF converted to Excel successfully!")

            # Provide download button for the converted file
            st.download_button(
                label="Download Converted Excel File",
                data=st.session_state.excel_data,
                file_name="converted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Automatically process and organize the data
            organized_df = process_and_organize_data(st.session_state.excel_data)

            if not organized_df.empty:
                st.write("Organized Data by Day and User:")
                st.dataframe(organized_df)

                # Convert organized data to Excel
                organized_excel_buffer = BytesIO()
                with pd.ExcelWriter(organized_excel_buffer, engine='xlsxwriter') as writer:
                    organized_df.to_excel(writer, sheet_name='OrganizedData', index=False)
                organized_excel_buffer.seek(0)
                st.session_state.organized_excel_data = organized_excel_buffer.getvalue()

                # Allow user to select shift type
                shift_type = st.sidebar.selectbox("Select Shift Type", ["None", "Morning", "Night"])

                if shift_type != "None":
                    st.sidebar.write(f"Processing {shift_type} shifts...")

                    # Save organized data to a temporary file
                    temp_file_buffer = BytesIO(st.session_state.organized_excel_data)
                    temp_file_buffer.seek(0)

                    if shift_type == "Morning":
                        results = process_morning_shifts(temp_file_buffer)
                    elif shift_type == "Night":
                        results = process_night_shifts(temp_file_buffer)

                    for sheet_name, result in results.items():
                        st.write(f"Results for {sheet_name}:")
                        st.download_button(
                            label=f"Download Results for {sheet_name}",
                            data=result,
                            file_name=f"{sheet_name}_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.write("No data was organized based on the provided criteria.")
        else:
            st.write("No data extracted from PDF.")

if __name__ == "__main__":
    main()
