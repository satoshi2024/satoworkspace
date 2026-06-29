#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from openpyxl import load_workbook
import os
import re
import unicodedata
from datetime import datetime

def read_any_file(path):
    if not path or not os.path.exists(path): return None
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        for enc in ['utf-8-sig', 'cp932', 'shift_jis', 'gbk', 'utf-8']:
            try:
                df = pd.read_csv(path, header=None, encoding=enc, dtype=str, low_memory=False)
                if not df.empty: return df
            except Exception: continue
        return pd.read_csv(path, header=None, dtype=str, low_memory=False)
    else:
        return pd.read_excel(path, header=None, dtype=str)

def build_5100_index(file_paths, log_callback):
    """
    全域地毯式扫描，通过关键词锚点自适应定位行号和列坐标
    """
    index = {}
    for fpath in file_paths:
        if not fpath or not os.path.exists(fpath): continue
        log_callback(f"正在全域自适应扫描: {os.path.basename(fpath)} ...")
        
        try:
            wb = load_workbook(fpath, data_only=True)
            for ws in wb.worksheets:
                sheet_name = ws.title
                index[sheet_name] = {'rows': {}, 'cols': {}}
                
                row_anchor = None
                col_anchors = [] # 存储 (列号数字, cell对象)

                # 1. 遍历全表，寻找锚点关键词
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value:
                            val = unicodedata.normalize('NFKC', str(cell.value)).replace(" ", "")
                            # 定位行锚点
                            if "行番号" in val: row_anchor = cell
                            # 定位列锚点
                            m = re.search(r'[\(（](\d+)[\)）]', val)
                            if m: col_anchors.append((m.group(1), cell))

                # 2. 基于行锚点偏移，向右/向下提取碎片数据
                if row_anchor:
                    for r_idx in range(row_anchor.row + 1, ws.max_row + 1):
                        digits = []
                        # 左右搜索碎片行号
                        for c_idx in range(max(1, row_anchor.column - 3), min(ws.max_column, row_anchor.column + 15)):
                            v = str(ws.cell(row=r_idx, column=c_idx).value or "")
                            for char in v:
                                if char.isdigit(): digits.append(char)
                        if len(digits) >= 2:
                            row_code = "".join(digits[:2]) # 提取行号前两位
                            index[sheet_name]['rows'][row_code] = r_idx

                # 3. 提取列锚点
                for num, cell in col_anchors:
                    index[sheet_name]['cols'][num] = cell.column_letter
            wb.close()
        except Exception as e:
            log_callback(f"❌ 扫描错误: {e}")
    return index

class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei 智能坐标同步工具 v3.4 (自适应全域版)")
        root.geometry("860x650")
        self.wrt_path, self.send_path = tk.StringVar(), tk.StringVar()
        self.file5100_1_path, self.file5100_2_path = tk.StringVar(), tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 坐标动态定位与同步工具", font=("Microsoft YaHei", 15, "bold"), fg="#2c3e50").pack(pady=10)
        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=5, fill="x")
        self._add_row(frame, "1. WRT 映射文件:", self.wrt_path, self.sel_wrt, 0)
        self._add_row(frame, "2. SEND 旧文件:", self.send_path, self.sel_send, 1)
        self._add_row(frame, "3. 最新 5100 文件①:", self.file5100_1_path, self.sel_5100_1, 2)
        self._add_row(frame, "4. 最新 5100 文件②(可选):", self.file5100_2_path, self.sel_5100_2, 3)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始同步并计算新坐标", command=self.transform, bg="#27ae60", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25)
        self.run_btn.pack(side="left", padx=10)

        self.log_text = scrolledtext.ScrolledText(self.root, height=20, width=100, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=5, fill="both", expand=True)

    def _add_row(self, p, txt, var, cmd, r):
        tk.Label(p, text=txt, width=20, anchor="e").grid(row=r, column=0, pady=5)
        tk.Entry(p, textvariable=var, width=55).grid(row=r, column=1, padx=5)
        tk.Button(p, text="浏览...", command=cmd).grid(row=r, column=2)

    def sel_wrt(self): self.wrt_path.set(filedialog.askopenfilename())
    def sel_send(self): self.send_path.set(filedialog.askopenfilename())
    def sel_5100_1(self): self.file5100_1_path.set(filedialog.askopenfilename())
    def sel_5100_2(self): self.file5100_2_path.set(filedialog.askopenfilename())
    
    def log(self, msg):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def transform(self):
        wrt_file, send_file = self.wrt_path.get().strip(), self.send_path.get().strip()
        f5100_1, f5100_2 = self.file5100_1_path.get().strip(), self.file5100_2_path.get().strip()
        if not wrt_file or not send_file: return messagebox.showerror("错误", "请选择文件")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            df_wrt = read_any_file(wrt_file)
            df_send = read_any_file(send_file)
            max_cols = max(len(df_send.columns), 5)
            
            send_dict = {str(row.iloc[1]).strip(): ["" if pd.isna(x) else str(x) for x in row.tolist()] + [""]*(max_cols - len(row.tolist())) for _, row in df_send.iterrows() if pd.notna(row.iloc[1])}

            self.log("=== 启动全域自适应坐标扫描 ===")
            index_5100 = build_5100_index([f5100_1, f5100_2], self.log)
            
            new_data = []
            for _, wrt_row in df_wrt.iterrows():
                a, b, c = str(wrt_row.iloc[0]), str(wrt_row.iloc[1]), str(wrt_row.iloc[2])
                cur_row = send_dict.get(b, [""] * max_cols).copy()
                cur_row[0], cur_row[1], cur_row[2] = a, b, c
                
                if len(a) >= 5 and a.isdigit():
                    col_code, row_code, target_num = a[-2:], a[-4:-2], re.sub(r'\D', '', a[:-4])
                    # 模糊匹配 Sheet
                    target_sheet = next((s for s in index_5100 if re.sub(r'\D', '', s) == target_num), None)
                    
                    if target_sheet:
                        cur_row[3] = target_sheet
                        r, c_letter = index_5100[target_sheet]['rows'].get(row_code), index_5100[target_sheet]['cols'].get(col_code)
                        if r and c_letter: cur_row[4] = f"{c_letter}{r}"
                new_data.append(cur_row)

            pd.DataFrame(new_data).to_csv(os.path.splitext(send_file)[0] + "_updated.csv", index=False, header=False, encoding='utf-8-sig')
            self.log("✅ 同步完成！已更新所有坐标。")
        except Exception as e: self.log(f"❌ 发生错误: {e}")
        finally: self.run_btn.config(state="normal", text="开始同步并计算新坐标")

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()
