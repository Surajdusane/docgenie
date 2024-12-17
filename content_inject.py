from docxtpl import DocxTemplate
import os

class DocumentGenerator:
    def __init__(self, template_path):
        self.template_path = template_path
        self.doc = DocxTemplate(template_path)
        
    def render_document(self, context, output_folder, output_filename):
       # Render the document with the given context
        self.doc.render(context)
        
        # Ensure the output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # Create the full file path
        file_path = os.path.join(output_folder, output_filename)
        
        # Save the rendered document
        self.doc.save(file_path)
        return file_path