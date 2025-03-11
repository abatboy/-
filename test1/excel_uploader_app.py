# FILEPATH: E:/hello world/Excel-Quotation-Program/excel_uploader_app.py

import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import os
import shutil
from datetime import datetime
from utils import create_database_folder, log_message

class ExcelUploaderApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel文件上传器")
        master.geometry("400x400")

        self.db_folder = create_database_folder()

        # 创建并设置控件
        self.create_widgets()

        self.selected_file = None

    def create_widgets(self):
        self.label = tk.Label(self.master, text="选择一个Excel文件 (.xlsx)")
        self.label.pack(pady=10)

        self.select_button = tk.Button(self.master, text="选择文件", command=self.select_file)
        self.select_button.pack(pady=5)

        self.file_label = tk.Label(self.master, text="未选择文件")
        self.file_label.pack(pady=5)

        self.upload_button = tk.Button(self.master, text="上传", command=self.upload_file, state=tk.DISABLED)
        self.upload_button.pack(pady=10)

        self.scan_button = tk.Button(self.master, text="扫描数据库", command=self.scan_database)
        self.scan_button.pack(pady=5)

        self.log_label = tk.Label(self.master, text="日志输出")
        self.log_label.pack(pady=(10, 5))

        self.log_text = tk.Text(self.master, height=10, width=50)
        self.log_text.pack(pady=5)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.upload_button.config(state=tk.NORMAL)
            self.log_message(f"已选择文件: {os.path.basename(file_path)}")
        else:
            self.selected_file = None
            self.file_label.config(text="未选择文件")
            self.upload_button.config(state=tk.DISABLED)

    def upload_file(self):
        if self.selected_file:
            try:
                load_workbook(self.selected_file)
                destination = os.path.join(self.db_folder, os.path.basename(self.selected_file))
                shutil.copy2(self.selected_file, destination)
                self.log_message(f"文件 '{os.path.basename(self.selected_file)}' 已上传并复制到数据库文件夹")
            except Exception as e:
                error_message = f"无法处理文件: {str(e)}"
                messagebox.showerror("错误", error_message)
                self.log_message(f"错误: {error_message}")
        else:
            messagebox.showwarning("警告", "请先选择一个文件")
            self.log_message("警告: 未选择文件就尝试上传")

    def scan_database(self):
        self.log_message("开始扫描数据库文件夹...")
        files = os.listdir(self.db_folder)
        if files:
            for file in files:
                file_path = os.path.join(self.db_folder, file)
                file_size = os.path.getsize(file_path)
                file_mtime = os.path.getmtime(file_path)
                mtime_str = datetime.fromtimestamp(file_mtime).strftime("%Y-%m-%d %H:%M:%S")
                self.log_message(f"文件: {file}, 大小: {file_size} 字节, 修改时间: {mtime_str}")
        else:
            self.log_message("数据库文件夹为空")
        self.log_message("扫描完成")

    def log_message(self, message):
        log_message(self.log_text, message)
