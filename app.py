from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import pandas as pd
from main import AutomationTool
from fastapi.middleware.cors import CORSMiddleware
import os
from pathlib import Path
from output_files import process_folder_and_create_zip

app = FastAPI()

origins = [
    "http://localhost",
    "http://localhost:3000",
    "*"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global storage for file paths
uploaded_files = {"csv": None, "template": None}

@app.post("/upload-csv/")
async def upload_csv(csv_file: UploadFile = File(...)):
    if csv_file.content_type != 'text/csv':
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a CSV file.")
    
    try:
        # Read CSV file to ensure it's valid
        df = pd.read_csv(csv_file.file)
        csv_file.file.seek(0)  # Reset file pointer after reading
        
        # Save CSV file
        csv_path = f"temp_{csv_file.filename}"
        with open(csv_path, "wb") as f:
            f.write(csv_file.file.read())
        
        # Store path in global dictionary
        uploaded_files["csv"] = csv_path
        
        return {"message": "CSV file uploaded successfully", "csv_path": csv_path}
    except pd.errors.EmptyDataError:
        raise HTTPException(status_code=400, detail="The uploaded CSV file is empty. Please upload a valid CSV file.")
    except pd.errors.ParserError:
        raise HTTPException(status_code=400, detail="The uploaded file is not a valid CSV. Please upload a valid CSV file.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing the CSV file: {e}")

@app.post("/upload-template/")
async def upload_template(template_file: UploadFile = File(...)):
    if template_file.content_type != 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a DOCX file.")
    
    # Save DOCX template
    template_path = f"temp_{template_file.filename}"
    try:
        with open(template_path, "wb") as f:
            f.write(template_file.file.read())
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error saving the template file: {e}")
    
    # Store path in global dictionary
    uploaded_files["template"] = template_path
    
    return {"message": "Template file uploaded successfully", "template_path": template_path}

@app.get("/generate-docs/")
async def generate_docs(folder_name: str):
    if not folder_name:
        raise HTTPException(status_code=400, detail="Folder name must be provided.")
    
    csv_path = uploaded_files.get("csv")
    template_path = uploaded_files.get("template")
    
    if not csv_path or not template_path:
        raise HTTPException(status_code=400, detail="Both CSV and DOCX files must be uploaded before generating documents.")
    
    try:
        output_folder = folder_name
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        tool = AutomationTool(csv_path, template_path, output_folder)
        tool.process_csv_and_generate_docs()
        
        return {"message": "Files generated successfully!", "output_folder": output_folder}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred while generating files: {e}")

@app.get("/convert-to-pdf-zip/")
async def convert_to_pdf_zip(folder_name: str):
    folder_path = Path(folder_name)
    
    # Check if the provided folder exists
    if not folder_path.exists() or not folder_path.is_dir():
        raise HTTPException(status_code=404, detail="Folder not found")

    # Process the folder and create the zip file
    zip_file_path = process_folder_and_create_zip(folder_path)

    # Return the zip file as a downloadable response
    return FileResponse(zip_file_path, media_type="application/zip", filename="pdfs.zip")

@app.get("/")
def read_root():
    return {"message": "Welcome to DocGenie API"}
