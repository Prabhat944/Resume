#!/usr/bin/env python3
"""
Script to convert DOCX resume to PDF
"""

import sys
import os

def convert_docx_to_pdf():
    """Convert DOCX to PDF using available method"""
    docx_file = 'Prabhat_Kumar_Resume_TwoColumn.docx'
    pdf_file = 'Prabhat_Kumar_Resume_TwoColumn.pdf'
    
    if not os.path.exists(docx_file):
        print(f"Error: {docx_file} not found!")
        return False
    
    # Try method 1: docx2pdf (requires Microsoft Word on macOS)
    try:
        from docx2pdf import convert
        print("Converting DOCX to PDF using docx2pdf...")
        convert(docx_file, pdf_file)
        if os.path.exists(pdf_file):
            print(f"✓ Successfully created: {pdf_file}")
            return True
    except Exception as e:
        print(f"docx2pdf failed: {e}")
    
    # Try method 2: Using subprocess with LibreOffice (if available)
    try:
        import subprocess
        # Try different LibreOffice paths
        libreoffice_paths = ['/Applications/LibreOffice.app/Contents/MacOS/soffice', 
                            'libreoffice', 'soffice']
        
        for lo_path in libreoffice_paths:
            try:
                result = subprocess.run(
                    [lo_path, '--headless', '--convert-to', 'pdf', 
                     '--outdir', '.', docx_file],
                    capture_output=True,
                    timeout=30
                )
                if result.returncode == 0 and os.path.exists(pdf_file):
                    print(f"✓ Successfully created: {pdf_file}")
                    return True
            except (FileNotFoundError, subprocess.TimeoutExpired):
                continue
    except Exception as e:
        print(f"LibreOffice conversion failed: {e}")
    
    print("\n" + "="*60)
    print("PDF conversion failed. Please try one of these options:")
    print("="*60)
    print("\nOption 1: Use Microsoft Word")
    print("  - Open the DOCX file in Microsoft Word")
    print("  - Go to File > Export > PDF")
    print("  - Save as PDF")
    print("\nOption 2: Use online converter")
    print("  - Visit: https://www.ilovepdf.com/word-to-pdf")
    print("  - Upload the DOCX file and convert")
    print("\nOption 3: Install LibreOffice")
    print("  - Download from: https://www.libreoffice.org/")
    print("  - Then run this script again")
    print("="*60)
    
    return False

if __name__ == '__main__':
    convert_docx_to_pdf()
