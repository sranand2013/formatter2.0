import streamlit as st
import pandas as pd
import pdfplumber
import io
import zipfile
import re
from datetime import datetime

st.set_page_config(page_title="CMS Attendance PDF to Excel Batch Converter", layout="wide")

def extract_pdf_data(pdf_file):
    """Extracts tabular data and merges wrap-around pages horizontally by USN."""
    student_dict = {}
    all_subjects_ordered = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
                
            header_rows = []
            data_rows = []
            
            # Separate headers from student data on the current page
            for row in table:
                clean_row = [str(cell).replace('\n', ' ').strip() if cell is not None else '' for cell in row]
                # Identify valid student rows by matching the USN pattern
                if clean_row[0] and re.match(r'^[0-9]{2}[A-Z]{4}[0-9]{4}$', clean_row[0].strip()):
                    data_rows.append(clean_row)
                elif not data_rows: 
                    header_rows.append(clean_row)
                    
            # Dynamically determine subjects for THIS specific page
            page_subjects = []
            if header_rows:
                num_cols = len(header_rows[-1])
                # Each subject in the PDF spans 4 columns (C, A, R, %)
                num_page_subjects = (num_cols - 2) // 4
                
                for i in range(num_page_subjects):
                    subject_text_parts = []
                    start_col = 2 + i * 4
                    end_col = start_col + 4
                    
                    for h_row in header_rows:
                        for c_idx in range(start_col, min(end_col, len(h_row))):
                            text = h_row[c_idx].strip()
                            if text and text.upper() not in ('C', 'R', 'A', '%', 'C R A %', 'USN', 'STUDENTS NAME'):
                                if not subject_text_parts or subject_text_parts[-1] != text:
                                    subject_text_parts.append(text)
                                    
                    raw_title = " ".join(subject_text_parts).strip()
                    raw_title = re.sub(r'\s+', ' ', raw_title)
                    
                    # Extract "CODE - NAME" if a code pattern is detected at the start
                    match = re.match(r'^([A-Z0-9\-]{6,20})\s+(.*)$', raw_title, re.IGNORECASE)
                    
                    if match:
                        subject_title = f"{match.group(1)} {match.group(2)}"
                    else:
                        subject_title = raw_title
                        
                    subject_title = subject_title if subject_title else f"Subject_P{page.page_number}_{i+1}"
                    page_subjects.append(subject_title)
                    
                    # Maintain the global horizontal column order
                    if subject_title not in all_subjects_ordered:
                        all_subjects_ordered.append(subject_title)

            # Merge the page's data horizontally into the master dictionary
            for row in data_rows:
                usn = row[0]
                name = row[1]
                
                if usn not in student_dict:
                    student_dict[usn] = {'Name': name, 'Subjects': {}}
                    
                for i, subj_name in enumerate(page_subjects):
                    start_col = 2 + i * 4
                    if start_col + 1 < len(row):
                        c_val = row[start_col] if row[start_col] else '0'
                        a_val = row[start_col + 1] if row[start_col + 1] else '0'
                        student_dict[usn]['Subjects'][subj_name] = {'C': c_val, 'A': a_val}

    return student_dict, all_subjects_ordered

def format_to_excel(student_dict, all_subjects):
    """Structures the merged dictionary data into a single-row-per-student Excel format."""
    if not student_dict:
        return None
        
    # Building flat header rows to avoid pandas MultiIndex export limitations
    top_header = ['S No', 'Reg/USN', 'Student Name', 'Section']
    sub_header = ['', '', '', '']
    
    for sub in all_subjects:
        top_header.extend([sub, '', ''])
        sub_header.extend(['C', 'A', '%'])
        
    top_header.extend(['Overall %', 'Total Classes Conducted', 'Total Classes Attended', 'Signature'])
    sub_header.extend(['', '', '', ''])
    
    final_data = [top_header, sub_header]
    
    for i, (usn, data) in enumerate(student_dict.items()):
        formatted_row = [i + 1, usn, data['Name'], 'A'] # S No, USN, Name, Section
        
        total_c = 0.0
        total_a = 0.0
        
        for sub in all_subjects:
            subj_data = data['Subjects'].get(sub, {'C': '0', 'A': '0'})
            c_str = subj_data['C'].strip()
            a_str = subj_data['A'].strip()
            
            c_str = '0' if c_str.lower() == 'none' or c_str == '' else c_str
            a_str = '0' if a_str.lower() == 'none' or a_str == '' else a_str
            
            c_val = float(c_str) if c_str.replace('.', '', 1).isdigit() else 0
            a_val = float(a_str) if a_str.replace('.', '', 1).isdigit() else 0
            
            # Recalculate pure percentage ignoring R
            pure_pct = round((a_val / c_val * 100), 2) if c_val > 0 else 0
            
            formatted_row.extend([c_str, a_str, pure_pct])
            total_c += c_val
            total_a += a_val
            
        # Overall Summary Calculation
        overall_pct = round((total_a / total_c * 100), 2) if total_c > 0 else 0
        formatted_row.extend([overall_pct, total_c, total_a, ''])
        final_data.append(formatted_row)
        
    return pd.DataFrame(final_data)

def create_excel_buffer(df, original_filename):
    """Writes the flat DataFrame to Excel, matching the exact internal metadata format."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Section - A')
        
        title_format = workbook.add_format({'bold': True, 'align': 'left', 'font_size': 11})
        worksheet.write(0, 0, 'JAIN (DEEMED-TO-BE UNIVERSITY) CENTER FOR MANAGEMENT STUDIES', title_format)
        worksheet.write(1, 0, f'Programme: BBA (Converted from {original_filename})')
        worksheet.write(2, 0, f'Generated Date Time: {datetime.now().strftime("%d-%m-%Y %H:%M:%S")}')
        
        # Write the dataframe without index or headers (as they are built into the rows)
        df.to_excel(writer, sheet_name='Section - A', startrow=3, index=False, header=False)
        
    output.seek(0)
    return output

# --- UI Layout ---
st.title("📄 Batch PDF to Excel Attendance Converter")
st.markdown("Upload multiple Lalbagh Campus PDF attendance reports. The system will extract the data, merge wrapped subjects horizontally, calculate pure attendance (ignoring Remedial classes), format it to match the CMS multi-tier Excel template, and package the files into a downloadable ZIP archive.")

uploaded_files = st.file_uploader("Upload PDF Reports", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Convert and Export as ZIP"):
        with st.spinner("Processing PDFs and generating Excel files..."):
            zip_buffer = io.BytesIO()
            valid_files_processed = 0
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for pdf_file in uploaded_files:
                    student_dict, extracted_subjects = extract_pdf_data(pdf_file)
                    
                    if student_dict:
                        df_formatted = format_to_excel(student_dict, extracted_subjects)
                        excel_buffer = create_excel_buffer(df_formatted, pdf_file.name)
                        
                        excel_filename = re.sub(r'\.pdf$', '.xlsx', pdf_file.name, flags=re.IGNORECASE)
                        zip_file.writestr(excel_filename, excel_buffer.getvalue())
                        valid_files_processed += 1
            
            if valid_files_processed > 0:
                zip_buffer.seek(0)
                st.success(f"Successfully converted {valid_files_processed} files.")
                st.download_button(
                    label="📥 Download Converted Excels (ZIP)",
                    data=zip_buffer,
                    file_name=f"CMS_Attendance_Converted_{datetime.now().strftime('%d%m%Y_%H%M')}.zip",
                    mime="application/zip"
                )
            else:
                st.error("No valid student data could be extracted from the uploaded PDFs. Please ensure they match the Lalbagh Campus format.")