# FILEPATH: E:/hello world/Excel-Quotation-Program/utils.py

import os
import tkinter as tk
from datetime import datetime

def create_database_folder():
    db_folder = "数据库"
    if not os.path.exists(db_folder):
        os.makedirs(db_folder)
    return db_folder

def log_message(log_text, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    log_text.insert(tk.END, log_entry)
    log_text.see(tk.END)
