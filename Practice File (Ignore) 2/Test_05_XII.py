import re
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

class DataProcessorXII:
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file
        self.read_data()
        self.process_data()
        self.calculate_qpi()
        self.calculate_percentage_counts()
        self.calculate_subject_percentage_counts()
        self.calculate_highest_marks()
        self.save_data_to_excel()

    def read_data(self):
        with open(self.input_file) as file:
            self.reader = file.read()
    
    def process_data(self):
        
    
    
    



if __name__ == "__main__":
    input_file = '65027 XII.txt'
    output_file = 'Output_Excel_2.xlsx'
    data_processor = DataProcessorXII(input_file, output_file)
