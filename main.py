import tkinter as tk
from processors import DataProcessor
from gui import App

if __name__ == "__main__":
    root = tk.Tk()
    processor = DataProcessor('', '', '')
    app = App(root, processor)
    root.mainloop()
