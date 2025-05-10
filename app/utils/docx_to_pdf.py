# utils/docx_to_pdf.py
from docx2pdf import convert
import shutil
import os

def convert_docx_to_pdf(input_path: str, output_path: str):
    # docx2pdf only works on Windows with paths, or needs Word installed
    # This approach assumes you're running on Windows
    # If Linux, consider calling libreoffice via subprocess

    temp_dir = os.path.dirname(output_path)
    convert(input_path, temp_dir)
    
    # The PDF will have same name but .pdf extension
    generated_pdf = os.path.join(temp_dir, os.path.basename(input_path).replace('.docx', '.pdf'))
    
    if os.path.exists(generated_pdf):
        shutil.move(generated_pdf, output_path)
    else:
        raise Exception("Failed to convert DOCX to PDF")