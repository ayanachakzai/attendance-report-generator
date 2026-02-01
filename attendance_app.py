"""
ATTENDANCE REPORT GENERATOR - WEB APP
======================================
A web-based interface for generating student attendance reports.

Built with Streamlit - run with: streamlit run attendance_app.py

Author: Ayan
Date: January 2026
"""

import streamlit as st
import pandas as pd
from docx import Document
import os
import subprocess
import tempfile
import zipfile
from io import BytesIO

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================

# WHAT IS st.set_page_config()?
# - Configures the web page appearance
# - Must be the FIRST Streamlit command
st.set_page_config(
    page_title="Attendance Report Generator",  # Browser tab title
    page_icon="📊",  # Browser tab icon (emoji)
    layout="wide"  # Use full page width
)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def check_libreoffice():
    """
    Check if LibreOffice is installed (needed for PDF conversion).
    
    WHAT IS A DOCSTRING?
    - Text in triple quotes after function definition
    - Explains what the function does
    - Good practice for documentation
    """
    libreoffice_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return os.path.exists(libreoffice_path)

def generate_reports(df, template_file, output_format, group_by):
    """
    Generate attendance reports for all students.
    
    Parameters:
    - df: DataFrame with student data
    - template_file: Uploaded template file
    - output_format: 'PDF' or 'DOCX'
    - group_by: Whether to organize by groups
    
    Returns:
    - BytesIO object containing zip file of all reports
    """
    
    # WHAT IS BytesIO()?
    # - Creates a file-like object in MEMORY (not on disk)
    # - We use it to create a zip file without saving to disk
    # - Can be downloaded directly by the user
    zip_buffer = BytesIO()
    
    # WHAT IS zipfile.ZipFile()?
    # - Creates a ZIP archive
    # - 'w' means write mode (create new zip)
    # - zipfile.ZIP_DEFLATED compresses the files
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        
        # Get unique groups if needed
        if group_by:
            groups = df['Group Ref'].unique()
        else:
            groups = [None]  # Single "group" for all students
        
        # Process each student
        for index, row in df.iterrows():
            
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
                # WHAT IS f-string WITH VARIABLE?
                # - Creates path like "Group 1/student_report.docx"
                # - Organizes files by group inside the zip
                zip_path = f"{student_group}/{bnu_id}_{row['Surname']}_{row['Name']}_Attendance_Report"
            else:
                zip_path = f"{bnu_id}_{row['Surname']}_{row['Name']}_Attendance_Report"
            
            # Save as DOCX or convert to PDF
            if output_format == "DOCX":
                # Save directly to zip as DOCX
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp:
                    doc.save(temp.name)
                    # WHAT IS zip_file.write()?
                    # - Adds a file to the zip archive
                    # - First argument: file on disk to add
                    # - Second argument: path inside the zip
                    zip_file.write(temp.name, f"{zip_path}.docx")
                    os.unlink(temp.name)
            
            else:  # PDF
                # Convert to PDF first
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                    doc.save(temp_docx.name)
                    
                    # Create temporary output directory
                    temp_dir = tempfile.mkdtemp()
                    
                    # Convert to PDF using LibreOffice
                    command = [
                        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                        '--headless',
                        '--convert-to',
                        'pdf',
                        '--outdir',
                        temp_dir,
                        temp_docx.name
                    ]
                    
                    subprocess.run(command, capture_output=True, text=True, check=True)
                    
                    # Get PDF filename
                    pdf_name = os.path.splitext(os.path.basename(temp_docx.name))[0] + '.pdf'
                    pdf_path = os.path.join(temp_dir, pdf_name)
                    
                    # Add PDF to zip
                    if os.path.exists(pdf_path):
                        zip_file.write(pdf_path, f"{zip_path}.pdf")
                        os.unlink(pdf_path)
                    
                    # Cleanup
                    os.unlink(temp_docx.name)
                    os.rmdir(temp_dir)
    
    # WHAT IS zip_buffer.seek(0)?
    # - Resets the "read position" to the beginning of the file
    # - Like rewinding a tape
    # - Necessary before reading/downloading the zip
    zip_buffer.seek(0)
    return zip_buffer

# ============================================================================
# MAIN APP INTERFACE
# ============================================================================

# WHAT IS st.title()?
# - Creates a large heading on the page
# - Uses markdown syntax
st.title("📊 Attendance Report Generator")

# WHAT IS st.markdown()?
# - Displays formatted text
# - Supports markdown (bold, italic, links, etc.)
st.markdown("Generate professional attendance reports for your students")

# WHAT IS st.divider()?
# - Creates a horizontal line separator
# - Makes the UI cleaner
st.divider()

# ============================================================================
# STEP 1: FILE UPLOADS
# ============================================================================

st.header("📁 Step 1: Upload Files")

# WHAT IS st.columns()?
# - Creates side-by-side columns on the page
# - [1, 1] means two equal-width columns
col1, col2 = st.columns(2)

with col1:
    # WHAT IS st.file_uploader()?
    # - Creates a file upload widget
    # - Returns the uploaded file object
    # - type=[...] restricts file types
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

# ============================================================================
# STEP 2: OPTIONS
# ============================================================================

st.header("⚙️ Step 2: Configure Options")

col3, col4 = st.columns(2)

with col3:
    # WHAT IS st.selectbox()?
    # - Creates a dropdown menu
    # - User can select one option
    # - Returns the selected value
    output_format = st.selectbox(
        "Output Format",
        ["PDF", "DOCX"],
        help="Choose whether to generate PDF or DOCX files"
    )

with col4:
    # WHAT IS st.checkbox()?
    # - Creates a checkbox (on/off toggle)
    # - Returns True if checked, False if not
    # - value=True makes it checked by default
    group_by = st.checkbox(
        "Organize by Groups",
        value=True,
        help="Organize reports into folders by student groups"
    )

st.divider()

# ============================================================================
# STEP 3: GENERATE REPORTS
# ============================================================================

st.header("🚀 Step 3: Generate Reports")

# WHAT IS st.button()?
# - Creates a clickable button
# - Returns True when clicked, False otherwise
# - type="primary" makes it blue/highlighted
if st.button("Generate Reports", type="primary", use_container_width=True):
    
    # Check if files are uploaded
    # WHAT IS 'if excel_file and template_file:'?
    # - Checks if BOTH files are uploaded
    # - 'and' is a LOGICAL OPERATOR (both must be True)
    if excel_file and template_file:
        
        # Check PDF requirements
        if output_format == "PDF" and not check_libreoffice():
            # WHAT IS st.error()?
            # - Displays a red error message
            # - icon adds an emoji to the message
            st.error(
                "❌ LibreOffice is not installed! PDF conversion requires LibreOffice. "
                "Please install it from https://www.libreoffice.org/download/download/"
            )
        else:
            try:
                # WHAT IS st.spinner()?
                # - Shows a loading animation
                # - Displays text while code runs
                # - Automatically disappears when done
                with st.spinner("Reading attendance data..."):
                    # Read Excel file
                    df = pd.read_excel(excel_file, skiprows=1)
                    
                    # Clean column names
                    df.columns = df.columns.str.strip()
                    
                    # Remove empty rows
                    df = df.dropna(subset=['BNU ID'])
                    
                    # Clean data
                    df['Name'] = df['Name'].str.strip()
                    df['Surname'] = df['Surname'].str.strip()
                
                # Show summary
                # WHAT IS st.info()?
                # - Displays a blue information message
                st.info(f"✅ Loaded {len(df)} students")
                
                if group_by:
                    groups = df['Group Ref'].unique()
                    st.info(f"📂 Found {len(groups)} groups")
                
                # Generate reports
                with st.spinner(f"Generating {len(df)} {output_format} reports... This may take a few minutes."):
                    # WHAT IS st.progress()?
                    # - Creates a progress bar
                    # - Updates as work progresses
                    progress_bar = st.progress(0)
                    
                    # For simplicity, we'll generate all at once
                    # (In production, you'd update progress for each student)
                    zip_buffer = generate_reports(df, template_file, output_format, group_by)
                    
                    progress_bar.progress(100)
                
                # Success message
                # WHAT IS st.success()?
                # - Displays a green success message
                st.success(f"🎉 Successfully generated {len(df)} reports!")
                
                # Download button
                # WHAT IS st.download_button()?
                # - Creates a button that downloads a file
                # - data: the file content to download
                # - file_name: what to name the downloaded file
                # - mime: file type for browser
                st.download_button(
                    label=f"⬇️ Download All Reports ({output_format})",
                    data=zip_buffer,
                    file_name=f"attendance_reports_{output_format.lower()}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )
                
            except Exception as e:
                # WHAT IS 'except Exception as e:'?
                # - Catches any error that occurs
                # - 'e' contains the error details
                st.error(f"❌ Error: {str(e)}")
                # WHAT IS st.exception(e)?
                # - Displays detailed error information
                # - Useful for debugging
                st.exception(e)
    
    else:
        # WHAT IS st.warning()?
        # - Displays an orange warning message
        st.warning("⚠️ Please upload both files before generating reports")

# ============================================================================
# SIDEBAR: INSTRUCTIONS
# ============================================================================

# WHAT IS st.sidebar?
# - Creates a sidebar on the left side of the page
# - Good for instructions, settings, etc.
with st.sidebar:
    st.header("📖 Instructions")
    
    st.markdown("""
    ### How to Use:
    
    1. **Upload Files**
       - Student attendance Excel file
       - Report template (DOCX)
    
    2. **Choose Options**
       - PDF or DOCX output
       - Organize by groups or not
    
    3. **Generate**
       - Click "Generate Reports"
       - Wait for processing
       - Download ZIP file
    
    ### Requirements:
    
    - **For PDF:** LibreOffice must be installed
    - **Excel format:** Should have columns for Name, Surname, BNU ID, Campus, LIVE, Group Ref
    
    ### Support:
    
    Contact Ayan for help!
    """)

# ============================================================================
# END OF APP
# ============================================================================

# HOW TO RUN THIS APP:
# ====================
# 
# 1. Install Streamlit:
#    pip install streamlit
# 
# 2. Run the app:
#    streamlit run attendance_app.py
# 
# 3. Open browser to:
#    http://localhost:8501
# 
# STREAMLIT BASICS:
# =================
# 
# DISPLAY WIDGETS:
# - st.title() - Large heading
# - st.header() - Medium heading
# - st.text() - Plain text
# - st.markdown() - Formatted text
# 
# INPUT WIDGETS:
# - st.button() - Clickable button
# - st.file_uploader() - File upload
# - st.selectbox() - Dropdown menu
# - st.checkbox() - On/off toggle
# - st.slider() - Number slider
# - st.text_input() - Text box
# 
# FEEDBACK WIDGETS:
# - st.success() - Green success message
# - st.error() - Red error message
# - st.warning() - Orange warning
# - st.info() - Blue information
# - st.spinner() - Loading animation
# - st.progress() - Progress bar
# 
# LAYOUT WIDGETS:
# - st.columns() - Side-by-side sections
# - st.sidebar - Left sidebar
# - st.divider() - Horizontal line
# - st.container() - Group widgets
# 
# DOWNLOAD:
# - st.download_button() - Download file
