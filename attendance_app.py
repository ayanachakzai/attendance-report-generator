"""
ATTENDANCE REPORT GENERATOR - WEB APP (Full Version)
====================================================
Works on Streamlit Cloud with PDF support via LibreOffice
"""

import streamlit as st
import pandas as pd
from docx import Document
import os
import subprocess
import tempfile
import zipfile
from io import BytesIO
import shutil

st.set_page_config(
    page_title="Attendance Report Generator",
    page_icon="📊",
    layout="wide"
)

def check_libreoffice():
    """Check if LibreOffice is installed"""
    # Check common LibreOffice paths
    possible_paths = [
        '/usr/bin/soffice',  # Linux/Streamlit Cloud
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',  # Mac
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe'  # Windows
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # Try finding it in PATH
    if shutil.which('soffice'):
        return shutil.which('soffice')
    
    return None

def generate_reports(df, template_file, output_format, group_by, libreoffice_path):
    """Generate reports for all students"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        
        total_students = len(df)
        
        for index, row in df.iterrows():
            
            # Update progress
            progress = int((index + 1) / total_students * 100)
            
            student_name = f"{row['Name']} {row['Surname']}"
            bnu_id = str(int(row['BNU ID']))
            campus = row['Campus']
            attendance = row['LIVE']
            
            if group_by:
                student_group = row['Group Ref']
            
            attendance_percent = attendance * 100
            
            if attendance_percent >= 80:
                attendance_category = "Excellent attendance"
            elif attendance_percent >= 70:
                attendance_category = "Very good attendance"
            elif attendance_percent >= 60:
                attendance_category = "Good attendance"
            else:
                attendance_category = "Attendance could be better"
            
            # Create document from template
            doc = Document(template_file)
            
            table = doc.tables[0]
            table.rows[4].cells[1].text = student_name
            table.rows[5].cells[1].text = bnu_id
            table.rows[6].cells[1].text = campus
            
            attendance_table = doc.tables[1]
            attendance_table.rows[1].cells[1].text = ""
            attendance_table.rows[2].cells[1].text = ""
            attendance_table.rows[3].cells[1].text = ""
            attendance_table.rows[4].cells[1].text = ""
            
            if attendance_category == "Excellent attendance":
                attendance_table.rows[1].cells[1].text = "Yes"
            elif attendance_category == "Very good attendance":
                attendance_table.rows[2].cells[1].text = "Yes"
            elif attendance_category == "Good attendance":
                attendance_table.rows[3].cells[1].text = "Yes"
            else:
                attendance_table.rows[4].cells[1].text = "Yes"
            
            # Determine file path in zip
            if group_by:
                base_path = f"{student_group}/{bnu_id}_{row['Surname']}_{row['Name']}_Attendance_Report"
            else:
                base_path = f"{bnu_id}_{row['Surname']}_{row['Name']}_Attendance_Report"
            
            # Save as DOCX or convert to PDF
            if output_format == "DOCX":
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp:
                    doc.save(temp.name)
                    zip_file.write(temp.name, f"{base_path}.docx")
                    os.unlink(temp.name)
            
            else:  # PDF
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                    doc.save(temp_docx.name)
                    
                    # Create temporary output directory
                    temp_dir = tempfile.mkdtemp()
                    
                    try:
                        # Convert to PDF using LibreOffice
                        command = [
                            libreoffice_path,
                            '--headless',
                            '--convert-to',
                            'pdf',
                            '--outdir',
                            temp_dir,
                            temp_docx.name
                        ]
                        
                        subprocess.run(
                            command,
                            capture_output=True,
                            text=True,
                            check=True,
                            timeout=30
                        )
                        
                        # Get PDF filename
                        pdf_name = os.path.splitext(os.path.basename(temp_docx.name))[0] + '.pdf'
                        pdf_path = os.path.join(temp_dir, pdf_name)
                        
                        # Add PDF to zip
                        if os.path.exists(pdf_path):
                            zip_file.write(pdf_path, f"{base_path}.pdf")
                            os.unlink(pdf_path)
                        
                    finally:
                        # Cleanup
                        os.unlink(temp_docx.name)
                        if os.path.exists(temp_dir):
                            shutil.rmtree(temp_dir)
    
    zip_buffer.seek(0)
    return zip_buffer

# Check LibreOffice availability
libreoffice_path = check_libreoffice()
has_libreoffice = libreoffice_path is not None

st.title("📊 Attendance Report Generator")
st.markdown("Generate professional attendance reports for your students")

# Show status
if has_libreoffice:
    st.success("✅ PDF conversion available!")
else:
    st.info("ℹ️ **Cloud Version:** DOCX only. For PDF support, add `packages.txt` file (see instructions in sidebar).")

st.divider()

st.header("📁 Step 1: Upload Files")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader(
        "Upload Student Attendance Excel",
        type=['xlsx', 'xls'],
        help="Upload your Excel file with student attendance data"
    )

with col2:
    template_file = st.file_uploader(
        "Upload Report Template (DOCX)",
        type=['docx'],
        help="Upload the Word document template for reports"
    )

st.divider()

st.header("⚙️ Step 2: Configure Options")

col3, col4 = st.columns(2)

with col3:
    if has_libreoffice:
        output_format = st.selectbox(
            "Output Format",
            ["PDF", "DOCX"],
            help="Choose whether to generate PDF or DOCX files"
        )
    else:
        output_format = "DOCX"
        st.info("📄 Output: DOCX (PDF requires LibreOffice)")

with col4:
    group_by = st.checkbox(
        "Organize by Groups",
        value=True,
        help="Organize reports into folders by student groups"
    )

st.divider()

st.header("🚀 Step 3: Generate Reports")

if st.button("Generate Reports", type="primary", use_container_width=True):
    
    if excel_file and template_file:
        
        if output_format == "PDF" and not has_libreoffice:
            st.error("❌ PDF conversion requires LibreOffice. Please use DOCX format or add packages.txt file.")
        else:
            try:
                with st.spinner("Reading attendance data..."):
                    df = pd.read_excel(excel_file, skiprows=1)
                    df.columns = df.columns.str.strip()
                    df = df.dropna(subset=['BNU ID'])
                    df['Name'] = df['Name'].str.strip()
                    df['Surname'] = df['Surname'].str.strip()
                
                st.info(f"✅ Loaded {len(df)} students")
                
                if group_by:
                    groups = df['Group Ref'].unique()
                    st.info(f"📂 Found {len(groups)} groups")
                
                with st.spinner(f"Generating {len(df)} {output_format} reports... This may take a few minutes."):
                    progress_bar = st.progress(0)
                    zip_buffer = generate_reports(df, template_file, output_format, group_by, libreoffice_path)
                    progress_bar.progress(100)
                
                st.success(f"🎉 Successfully generated {len(df)} reports!")
                
                st.download_button(
                    label=f"⬇️ Download All Reports ({output_format})",
                    data=zip_buffer,
                    file_name=f"attendance_reports_{output_format.lower()}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
    
    else:
        st.warning("⚠️ Please upload both files before generating reports")

with st.sidebar:
    st.header("📖 Instructions")
    
    st.markdown("""
    ### How to Use:
    
    1. **Rename the files*
        - Rename your Excel file to `student_attendance.xlsx`
        - Rename your template file to `template.docx`
    
    2. **Upload Files**
       - Student attendance Excel file (student_attendance.xlsx)
       - Report template (DOCX) (template.docx)
    
    3. **Choose Options**
       - PDF or DOCX output
       - Organize by groups or not
    
    4. **Generate**
       - Click "Generate Reports"
       - Wait for processing
       - Download ZIP file
    
    
    ### Requirements:
    
    - **Rename the sheet file to "student_attendance.xlsx" and the template file to "template.docx" for the program to work.
    - **Excel format:** Columns for Name, Surname, BNU ID, Campus, LIVE, Group Ref
    
    ### Support:
    
    Contact Ayan (ayan.achakzai@magnacartacollege.ac.uk) for help!
    """)