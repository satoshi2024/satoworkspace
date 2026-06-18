import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

def process_strict_mapping(old_csv_path, old_excel_path, new_excel_path, target_sheet, added_col, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载 Excel 文件...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        # ==========================================
        # 1. 结构比对 (V4引擎：追踪平移)
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在比对物理结构与侦测最新表头...\n")
        log_widget.update()

        def get_col_sig(ws, col_idx):
            vals = []
            for r in range(1, 50):
                val = ws.cell(row=r, column=col_idx).value
                v_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                v_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', v_str)
                vals.append(v_str)
            return "|".join(vals)

        old_col_sigs = [get_col_sig(ws_old, c) for c in range(1, max_c)]
        new_col_sigs = [get_col_sig(ws_new, c) for c in range(1, max_c)]

        sm_col = difflib.SequenceMatcher(None, old_col_sigs, new_col_sigs)
        col_map = {}
        for tag, i1, i2, j1, j2 in sm_col.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    col_map[old_i + 1] = new_j + 1

        def get_row_sig(ws, row_idx):
            vals = []
            for c in range(1, 50):
                val = ws.cell(row=row_idx, column=c).value
                v_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                v_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', v_str)
                vals.append(v_str)
            return "|".join(vals)

        old_row_sigs = [get_row_sig(ws_old, r) for r in range(1, max_r)]
        new_row_sigs = [get_row_sig(ws_new, r) for r in range(1, max_r)]
        sm_row = difflib.SequenceMatcher(None, old_row_sigs, new_row_sigs)
        row_map = {}
        for tag, i1, i2, j1, j2 in sm_row.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    row_map[old_i + 1] = new_j + 1

        # ==========================================
        # 2. 精准表头侦测 (兼容全角数字)
        # ==========================================
        new_col_to_logic_id = {}
        target_phys_col = None # 如果输入了30，记录30在Excel的物理列
        
        for c in range(1, max_c):
            for r in range(1, 50):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").
