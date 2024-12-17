from csv_to_json import CSVConverter
from content_inject import DocumentGenerator
import os

class AutomationTool:
    def __init__(self, csv_file, template_file, output_folder):
        self.csv_file = csv_file
        self.template_file = template_file
        self.output_folder = output_folder
        self.csv_converter = CSVConverter(self.csv_file)
        self.doc_generator = DocumentGenerator(self.template_file)

    def process_csv_and_generate_docs(self):
        # Read and convert CSV to dictionary
        data_list = self.csv_converter.to_dict()

        # Generate documents for each entry in the data list
        for i, data in enumerate(data_list):
            output_filename = f"output_{i+1}.docx"
            self.doc_generator.render_document(data, self.output_folder, output_filename)
            print(f"Document saved to: {os.path.join(self.output_folder, output_filename)}")
