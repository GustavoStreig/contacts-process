import tkinter as tk
from tkinter import ttk
from file_selectors import FileSelectors

class App:
    def __init__(self, root, processor):
        self.processor = processor
        self.root = root  # Adiciona root como um atributo da classe

        self.root.title('Excel Processor')
        self.root.geometry('500x200')

        self.excel_file_var = tk.StringVar()
        self.csv_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()

        self.file_selectors = FileSelectors(self.excel_file_var, self.csv_file_var, self.output_dir_var)

        self.create_widgets()
            
    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="10")  # Use self.root em vez de root
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W)
        self.excel_file_entry = ttk.Entry(frame, textvariable=self.excel_file_var, width=50)
        self.excel_file_entry.grid(row=0, column=1)
        ttk.Button(frame, text="Browse", command=self.file_selectors.select_excel_file).grid(row=0, column=2)

        ttk.Label(frame, text="CSV File:").grid(row=1, column=0, sticky=tk.W)
        self.csv_file_entry = ttk.Entry(frame, textvariable=self.csv_file_var, width=50)
        self.csv_file_entry.grid(row=1, column=1)
        ttk.Button(frame, text="Browse", command=self.file_selectors.select_csv_file).grid(row=1, column=2)

        ttk.Label(frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W)
        self.output_dir_entry = ttk.Entry(frame, textvariable=self.output_dir_var, width=50)
        self.output_dir_entry.grid(row=2, column=1)
        ttk.Button(frame, text="Browse", command=self.file_selectors.select_output_directory).grid(row=2, column=2)

        ttk.Button(frame, text="Process", command=self.process).grid(row=3, column=1, pady=10)

    def process(self):
        self.processor.excel_file_path = self.excel_file_var.get()
        self.processor.csv_file_path = self.csv_file_var.get()
        self.processor.output_folder = self.output_dir_var.get()
        self.processor.process_file()
