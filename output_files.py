import docx2txt
from fpdf import FPDF
import os
from pathlib import Path
import zipfile

def convert_docx_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF by extracting text and writing to a PDF."""
    try:
        # Step 1: Extract text from DOCX file
        print(f"Extracting text from {docx_path}...")
        text = docx2txt.process(str(docx_path))
        
        if not text:
            print(f"Warning: No text found in {docx_path}. Skipping conversion.")
            return False  # No content to convert

        # Step 2: Initialize FPDF to create a PDF
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Step 3: Add the extracted text to the PDF
        lines = text.splitlines()
        for line in lines:
            pdf.cell(200, 10, txt=line, ln=True, align="L")

        # Step 4: Output the PDF to the specified path
        pdf.output(str(pdf_path))
        print(f"Converted {docx_path} to {pdf_path}")
        return True  # Successfully converted
    except Exception as e:
        print(f"Error converting {docx_path} to PDF: {e}")
        return False

def create_zip_from_pdfs(pdf_folder, zip_path):
    """Create a ZIP file containing all PDFs in the folder."""
    try:
        # Ensure the zip file is created and open it in write mode
        print(f"Creating ZIP file at {zip_path}...")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(pdf_folder):
                for file in files:
                    if file.endswith('.pdf'):
                        file_path = os.path.join(root, file)
                        # Make sure to add with the correct relative path inside the ZIP
                        arcname = os.path.relpath(file_path, pdf_folder)
                        zipf.write(file_path, arcname)
                        print(f"Added {file} to ZIP archive.")
        print(f"Created ZIP file successfully at: {zip_path}")
        return True
    except Exception as e:
        print(f"Error creating ZIP file: {e}")
        return False

def process_folder_and_create_zip(folder_path):
    """Process DOCX files in a folder, convert them to PDFs, and zip the PDFs."""
    folder_path = Path(folder_path)
    pdf_folder = folder_path / "converted_pdfs"
    
    # Ensure the PDFs folder exists
    if not pdf_folder.exists():
        pdf_folder.mkdir()

    # Step 1: Iterate over all DOCX files in the folder and convert them to PDFs
    pdf_files = []
    for docx_file in folder_path.glob("*.docx"):
        pdf_file = pdf_folder / (docx_file.stem + ".pdf")
        if convert_docx_to_pdf(docx_file, pdf_file):
            pdf_files.append(pdf_file)

    # If no PDFs were generated, exit early
    if not pdf_files:
        print("No PDFs were generated. Exiting...")
        return None

    # Step 2: Create a ZIP file containing all the converted PDFs
    zip_file_path = folder_path / "converted_pdfs.zip"
    if create_zip_from_pdfs(pdf_folder, zip_file_path):
        # Step 3: Clean up by deleting the folder with individual PDFs after zipping
        try:
            for pdf_file in pdf_folder.glob("*.pdf"):
                pdf_file.unlink()  # Remove each PDF file
            pdf_folder.rmdir()  # Remove the empty folder
            print(f"Cleaned up the PDFs folder: {pdf_folder}")
        except Exception as e:
            print(f"Error cleaning up the PDFs folder: {e}")
        return zip_file_path
    else:
        print(f"Failed to create ZIP file at {zip_file_path}")
        return None

# # Example usage:
# folder_path = "/path/to/your/folder"  # Specify the folder containing DOCX files
# zip_file = process_folder_and_create_zip(folder_path)
# if zip_file:
#     print(f"ZIP file created successfully: {zip_file}")
# else:
#     print("Failed to create ZIP file.")
