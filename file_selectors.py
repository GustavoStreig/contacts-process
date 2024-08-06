import tkinter as tk
from tkinter import filedialog

class FileSelectors:
    def __init__(self, excel_file_var, csv_file_var, output_dir_var):
        self.excel_file_var = excel_file_var
        self.csv_file_var = csv_file_var
        self.output_dir_var = output_dir_var

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls;*.xlsx')])
        self.excel_file_var.set(file_path)

    def select_csv_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])
        self.csv_file_var.set(file_path)

    def select_output_directory(self):
        directory = filedialog.askdirectory()
        self.output_dir_var.set(directory)
