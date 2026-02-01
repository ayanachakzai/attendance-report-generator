"""
DOCX TO PDF CONVERTER (LIBREOFFICE VERSION)
============================================
This script converts all .docx files to PDF using LibreOffice.
Use this if you DON'T have Microsoft Word installed.

Author: Ayan
Date: January 2026
"""

# ============================================================================
# SECTION 1: IMPORTING LIBRARIES
# ============================================================================

import os
import glob
import subprocess
# WHAT IS 'subprocess'?
# - A library for running external programs/commands
# - We'll use it to run LibreOffice from Python


# ============================================================================
# SECTION 2: CONFIGURATION
# ============================================================================

INPUT_DIR = '/Users/ayanachakzai/Downloads/attendance report/attendance_reports'

OUTPUT_DIR = '/Users/ayanachakzai/Downloads/attendance report/attendance_reports_pdf'

# Path to LibreOffice on Mac
# WHAT IS THIS PATH?
# - Default installation location of LibreOffice on Mac
# - If you installed LibreOffice elsewhere, change this path
LIBREOFFICE_PATH = '/Applications/LibreOffice.app/Contents/MacOS/soffice'


# ============================================================================
# SECTION 3: CHECK IF LIBREOFFICE IS INSTALLED
# ============================================================================

if not os.path.exists(LIBREOFFICE_PATH):
    # WHAT IS 'os.path.exists()'?
    # - A FUNCTION that checks if a file/folder exists
    # - Returns True if exists, False if not
    
    print("❌ ERROR: LibreOffice not found!")
    print(f"   Expected location: {LIBREOFFICE_PATH}")
    print("\n📥 Please install LibreOffice:")
    print("   1. Download from: https://www.libreoffice.org/download/download/")
    print("   2. Install it like any other Mac app")
    print("   3. Run this script again")
    exit(1)


# ============================================================================
# SECTION 4: CREATE OUTPUT FOLDER
# ============================================================================

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================================
# SECTION 5: FIND ALL DOCX FILES
# ============================================================================

print("🔍 Looking for .docx files...\n")

docx_pattern = os.path.join(INPUT_DIR, '*.docx')
docx_files = glob.glob(docx_pattern)

print(f"✅ Found {len(docx_files)} DOCX files to convert\n")

if not docx_files:
    print("❌ No .docx files found!")
    print(f"   Make sure files exist in: {INPUT_DIR}")
    exit(1)


# ============================================================================
# SECTION 6: CONVERT EACH DOCX TO PDF
# ============================================================================

print("🔄 Starting conversion with LibreOffice...\n")

for index, file_path in enumerate(docx_files):
    
    filename = os.path.basename(file_path)
    
    print(f"📄 Converting ({index + 1}/{len(docx_files)}): {filename}")
    
    try:
        # WHAT IS 'subprocess.run()'?
        # - A FUNCTION that runs external commands/programs
        # - Like typing commands in Terminal, but from Python
        
        # Build the LibreOffice command
        # WHAT IS THIS COMMAND?
        # - Tells LibreOffice to:
        #   --headless = run without opening a window
        #   --convert-to pdf = convert to PDF format
        #   --outdir = where to save the PDF
        #   file_path = the file to convert
        command = [
            LIBREOFFICE_PATH,
            '--headless',
            '--convert-to',
            'pdf',
            '--outdir',
            OUTPUT_DIR,
            file_path
        ]
        # WHAT IS A LIST OF STRINGS?
        # - Each item is one part of the command
        # - subprocess.run() will join them with spaces
        
        # Run the command
        result = subprocess.run(
            command,
            capture_output=True,
            # WHAT IS 'capture_output=True'?
            # - Captures what the program prints
            # - Stores it so we can check for errors
            
            text=True,
            # WHAT IS 'text=True'?
            # - Returns output as TEXT (string)
            # - Not as raw bytes
            
            check=True
            # WHAT IS 'check=True'?
            # - If command fails, raises an error
            # - Lets us catch it in the except block
        )
        
        print(f"   ✅ Saved to: {OUTPUT_DIR}\n")
        
    except subprocess.CalledProcessError as e:
        # WHAT IS 'CalledProcessError'?
        # - An error that occurs when external command fails
        # - Contains details about what went wrong
        
        print(f"   ❌ Error converting {filename}")
        print(f"   Error details: {e}\n")
        
    except Exception as e:
        # Catch any other errors
        print(f"   ❌ Unexpected error: {e}\n")


# ============================================================================
# SECTION 7: COMPLETION MESSAGE
# ============================================================================

print("=" * 50)
print("✅ CONVERSION COMPLETE!")
print(f"📁 PDF files saved in: {OUTPUT_DIR}")
print("=" * 50)


# ============================================================================
# END OF SCRIPT
# ============================================================================

# COMPARISON: docx2pdf vs LibreOffice
# ===================================
# 
# docx2pdf (Microsoft Word):
# - Pros: Fast, preserves formatting perfectly
# - Cons: Requires Microsoft Word (paid software)
# 
# LibreOffice:
# - Pros: Free, open-source, works without Word
# - Cons: Slightly slower, formatting might differ slightly
# 
# WHICH SHOULD YOU USE?
# - If you have Word → use convert_docx_to_pdf.py
# - If you don't have Word → use this script
# 
# SUBPROCESS EXPLAINED:
# ====================
# subprocess.run(['program', 'arg1', 'arg2'])
# 
# Is like typing in Terminal:
# program arg1 arg2
# 
# Example:
# subprocess.run(['ls', '-la', '/Users'])
# = ls -la /Users
# 
# Our LibreOffice command:
# soffice --headless --convert-to pdf --outdir /path/to/output file.docx
