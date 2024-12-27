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
        # Ensure the zip file is created and open it in write mode (binary mode)
        print(f"Creating ZIP file at {zip_path}...")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(pdf_folder):
                for file in files:
                    if file.endswith('.pdf'):
                        file_path = os.path.join(root, file)
                        # Add with the correct relative path inside the ZIP
                        arcname = os.path.relpath(file_path, pdf_folder)
                        zipf.write(file_path, arcname)
                        print(f"Added {file} to ZIP archive.")
        
        # Validate the ZIP file before returning
        if validate_zip(zip_path):
            print(f"ZIP file created successfully at: {zip_path}")
            return True
        else:
            print(f"Error: The ZIP file {zip_path} seems to be corrupted.")
            return False
    except Exception as e:
        print(f"Error creating ZIP file: {e}")
        return False


def validate_zip(zip_path):
    """Validate the ZIP file integrity by checking its contents."""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            # Test the ZIP file integrity
            corrupted_file = zipf.testzip()  # Returns None if all files are intact
            if corrupted_file is None:
                print("ZIP file is valid.")
                return True
            else:
                print(f"Corrupted file in ZIP: {corrupted_file}")
                return False
    except zipfile.BadZipFile:
        print(f"Error: The file {zip_path} is not a valid ZIP file.")
        return False
    except Exception as e:
        print(f"Unexpected error while validating ZIP file: {e}")
        return False


def process_folder_and_create_zip(folder_path):
    """Process DOCX files in a folder, convert them to PDFs, and create a ZIP file containing the PDFs."""
    folder_path = Path(folder_path)

    # Step 1: Collect all DOCX files in the folder
    docx_files = list(folder_path.glob("*.docx"))
    if not docx_files:
        print("No DOCX files found in the folder. Exiting...")
        return None

    # Step 2: Convert DOCX files to PDFs
    pdf_folder = folder_path / "pdfs"
    pdf_folder.mkdir(exist_ok=True)  # Create a folder to store the PDFs

    for docx_file in docx_files:
        pdf_file = pdf_folder / (docx_file.stem + ".pdf")
        if not convert_docx_to_pdf(docx_file, pdf_file):
            print(f"Failed to convert {docx_file.name} to PDF.")
            return None  # Stop the process if any conversion fails

    # Step 3: Create a ZIP file containing all the PDFs
    zip_file_path = folder_path / "pdf_files.zip"
    try:
        with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for folder_path in pdf_folder.glob("*.docx"):
                zipf.write(pdf_file, arcname=pdf_file.name)
                print(f"Added {pdf_file.name} to ZIP archive.")
        
        # Validate the ZIP file before returning
        if validate_zip(zip_file_path):
            print(f"ZIP file created successfully at: {zip_file_path}")
            return zip_file_path
        else:
            print(f"Error: The ZIP file {zip_file_path} seems to be corrupted.")
            return None
    except Exception as e:
        print(f"Failed to create ZIP file: {e}")
        return None

# Example usage:
# folder_path = "/path/to/your/folder"  # Specify the folder containing DOCX files
# zip_file = process_folder_and_create_zip(folder_path)
# if zip_file:
#     print(f"ZIP file created successfully: {zip_file}")
# else:
#     print("Failed to create ZIP file.")
