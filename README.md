#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, scrolledtext
from openpyxl import load_workbook
import unicodedata
from datetime import datetime

class DiagnosticApp:
    def __init__(self, root):
        self.root = root
        root.title("5100 文件深度诊断工具")
        root.geometry("700x500")
        self.btn = tk.Button(root, text="选择 5100 文件并诊断结构", command=self.diagnose, bg="blue", fg="white")
        self.btn.pack(pady=10)
        self.log_text = scrolledtext.ScrolledText(root, width=80, height=25)
        self.log_text.pack(padx=10, pady=10)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def diagnose(self):
        fpath = filedialog.askopenfilename(title="选择 5100 文件")
        if not fpath: return
        
        self.log(f"正在读取: {fpath}")
        try:
            wb = load_workbook(fpath, data_only=True)
            self.log(f"✅ 文件打开成功！")
            self.log(f"检测到的所有 Sheet 页: {wb.sheetnames}")
            
            # 随机测试第一个 Sheet 的前 20 行，看看能不能抓到“行番号”
            ws = wb.worksheets[0]
            self.log(f"\n开始诊断 Sheet: {ws.title} 的前 20 行内容:")
            for r in range(1, 21):
                row_vals = []
                for c in range(1, 10):
                    val = ws.cell(row=r, column=c).value
                    if val:
                        row_vals.append(str(val))
                if row_vals:
                    self.log(f"行 {r}: {' | '.join(row_vals)}")
            
            self.log("\n--- 诊断完成 ---")
            self.log("请截图发给我，我立刻就能看到程序为什么‘看不见’你的数据！")
        except Exception as e:
            self.log(f"❌ 错误: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    DiagnosticApp(root).mainloop()
