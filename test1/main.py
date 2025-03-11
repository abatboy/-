# FILEPATH: E:/hello world/Excel-Quotation-Program/main.py

import tkinter as tk
from excel_uploader_app import ExcelUploaderApp

def main():
    root = tk.Tk()
    app = ExcelUploaderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
