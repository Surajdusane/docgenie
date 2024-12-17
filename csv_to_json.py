import pandas as pd

class CSVConverter:
    def __init__(self, csv_file_path):
        self.csv_file_path = csv_file_path
        self.data_frame = pd.read_csv(self.csv_file_path)

    def to_dict(self):
        return self.data_frame.to_dict(orient='records')