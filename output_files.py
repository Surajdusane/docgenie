import docx2txt
from fpdf import FPDF
import os
from pathlib import Path
import zipfile

def convert_docx_to_pdf(docx_path, pdf_path):
    # Step 1: Extract text from DOCX file
    text = docx2txt.process(str(docx_path))
    
    # Step 2: Initialize FPDF to create a PDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Step 3: Add the extracted text to the PDF (page by page)
    # Split the text into lines and add them to the PDF
    lines = text.splitlines()
    for line in lines:
        pdf.cell(200, 10, txt=line, ln=True, align="L")
    
    # Step 4: Output the PDF to the specified path
    pdf.output(str(pdf_path))
    print(f"Converted {docx_path} to {pdf_path}")

def create_zip_from_pdfs(pdf_folder, zip_path):
    # Create a zip file containing all PDFs in the folder
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(pdf_folder):
            for file in files:
                if file.endswith('.pdf'):
                    zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), pdf_folder))

def process_folder_and_create_zip(folder_path):
    folder_path = Path(folder_path)
    pdf_folder = folder_path / "converted_pdfs"

    # Create a directory to store the PDFs
    if not pdf_folder.exists():
        pdf_folder.mkdir()

    # Iterate over all DOCX files in the folder and convert them to PDF
    for docx_file in folder_path.glob("*.docx"):
        pdf_file = pdf_folder / (docx_file.stem + ".pdf")
        convert_docx_to_pdf(docx_file, pdf_file)

    # Create a zip file containing all the converted PDFs
    zip_file_path = folder_path / "converted_pdfs.zip"
    create_zip_from_pdfs(pdf_folder, zip_file_path)

    # Clean up by deleting the folder with individual PDFs
    for pdf_file in pdf_folder.glob("*.pdf"):
        pdf_file.unlink()
    pdf_folder.rmdir()

    print(f"Created zip file at: {zip_file_path}")
    return zip_file_path

# Example usage:
# folder_path = "output"
# zip_file = process_folder_and_create_zip(folder_path)
# print(f"ZIP file path: {zip_file}")
