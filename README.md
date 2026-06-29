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
            except Exception:
                continue
        return pd.read_csv(path, header=None, dtype=str, low_memory=False)
    else:
        return pd.read_excel(path, header=None, dtype=str)

def build_5100_index(file_paths, log_callback):
    index = {}
    for fpath in file_paths:
        if not fpath or not os.path.exists(fpath): continue
        log_callback(f"正在启动碎片拼接扫描: {os.path.basename(fpath)} ... (请稍候)")
        
        try:
            wb = load_workbook(fpath, data_only=True)
        except Exception as e:
            log_callback(f"❌ 无法读取文件: {e}")
            continue

        for sheet_name in wb.sheetnames:
            if sheet_name not in index:
                index[sheet_name] = {'rows': {}, 'cols': {}}
            ws = wb[sheet_name]

            # ==========================================
            # 1. 寻找“行番号”所在的锚点位置
            # ==========================================
            xing_col_idx = -1
            xing_row_idx = -1
            for r_idx in range(1, 100):
                for cell in ws[r_idx]:
                    if cell.value:
                        val_str = unicodedata.normalize('NFKC', str(cell.value)).replace(" ", "")
                        if "行番号" in val_str:
                            xing_col_idx = cell.column
                            xing_row_idx = cell.row
                            break
                if xing_col_idx != -1: break

            # ==========================================
            # 2. 智能提取表头列 (Columns)
            # ==========================================
            cols_map = {}
            for r_idx in range(1, 100):
                row_pure_numbers = {}
                row_paren_numbers = {}
                for cell in ws[r_idx]:
                    if cell.value is not None:
                        val_str = unicodedata.normalize('NFKC', str(cell.value)).strip()
                        
                        # 匹配 (12), [12], （１２）
                        m_paren = re.match(r'^[\(\[<（]?\s*(\d+)\s*[\)\]>）]?$', val_str)
                        if m_paren:
                            num = str(int(m_paren.group(1))) # 01 变 1，12 变 12
                            if "(" in val_str or "（" in val_str or "<" in val_str:
                                row_paren_numbers[num] = cell.column_letter
                            elif int(num) <= 150: # 纯数字不能太大
                                row_pure_numbers[num] = cell.column_letter
                
                # 保存括号数字
                for k, v in row_paren_numbers.items():
                    if k not in cols_map: cols_map[k] = v
                
                # 如果这行有多个数字，说明确实是表头行，存入纯数字
                if len(row_pure_numbers) >= 2 or len(row_paren_numbers) >= 1:
                    for k, v in row_pure_numbers.items():
                        if k not in cols_map: cols_map[k] = v
            
            index[sheet_name]['cols'] = cols_map

            # ==========================================
            # 3. 智能提取行番号 (Rows) - 碎片拼接法
            # ==========================================
            rows_map = {}
            if xing_col_idx != -1:
                # 从“行番号”下一行开始往下扫
                for r_idx in range(xing_row_idx + 1, ws.max_row + 1):
                    digits = []
                    # 针对“方眼纸”排版：向右扫描 15 个极其狭窄的单元格，把打碎的数字吸起来
                    for c_idx in range(xing_col_idx, xing_col_idx + 15):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        if cell.value is not None:
                            val_str = unicodedata.normalize('NFKC', str(cell.value))
                            # 提取所有数字字符
                            for char in val_str:
                                if char.isdigit():
                                    digits.append(char)
                    
                    if len(digits) >= 2:
                        # 你的规则：3位行番号，无视第三位 -> 取前两位 (如 0,4,1 变成 "04")
                        row_code = "".join(digits[:2])
                        if row_code not in rows_map:
                            rows_map[row_code] = r_idx
            
            index[sheet_name]['rows'] = rows_map
        wb.close()
    return index

class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei 坐标动态计算工具 v3.2 (方眼纸粉碎版)")
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
        self.run_btn = tk.Button(btn_frame, text="开始同步并计算新坐标", command=self.transform, bg="#e74c3c", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25)
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
        if not wrt_file or not send_file: return messagebox.showerror("错误", "请选择 WRT 和 SEND")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            df_wrt = read_any_file(wrt_file)
            df_send = read_any_file(send_file)
            max_cols = max(len(df_send.columns), 5)
            
            send_dict = {}
            for _, row in df_send.iterrows():
                b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if b_val and b_val not in send_dict:
                    lst = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                    lst.extend([""] * (max_cols - len(lst)))
                    send_dict[b_val] = lst

            self.log("=== 开始读取并破解 5100 坐标结构 ===")
            index_5100 = build_5100_index([f5100_1, f5100_2], self.log)
            self.log("✅ 5100 坐标字典建立完成！\n")

            new_data = []
            calc_success = 0
            calc_fail = 0
            intercept_count = 0

            self.log("=== 开始计算最新坐标 ===")
            for _, wrt_row in df_wrt.iterrows():
                wrt_a = str(wrt_row.iloc[0]).strip() if pd.notna(wrt_row.iloc[0]) else ""
                wrt_b = str(wrt_row.iloc[1]).strip() if pd.notna(wrt_row.iloc[1]) else ""
                wrt_c = str(wrt_row.iloc[2]).strip() if pd.notna(wrt_row.iloc[2]) else ""
                if not wrt_b: continue
                
                cur_row = send_dict.get(wrt_b, [""] * max_cols).copy()
                cur_row[0], cur_row[1] = wrt_a, wrt_b
                if wrt_c: cur_row[2] = wrt_c

                if len(wrt_a) >= 5 and wrt_a.isdigit():
                    # 解析 A 列：前缀=表名，中间两位=行号，最后两位=列表头
                    col_code = str(int(wrt_a[-2:]))
                    row_code = wrt_a[-4:-2]
                    sheet_num = wrt_a[:-4]
                    target_sheet = f"表{sheet_num}"
                    
                    if target_sheet in index_5100:
                        cur_row[3] = target_sheet 
                        
                        r = index_5100[target_sheet]['rows'].get(row_code)
                        c_letter = index_5100[target_sheet]['cols'].get(col_code)
                        
                        if r and c_letter:
                            cur_row[4] = f"{c_letter}{r}"
                            calc_success += 1
                        else:
                            calc_fail += 1
                            if calc_fail <= 10: # 最多打印10条防止刷屏
                                self.log(f"⚠️ [警告] A={wrt_a}: 找到了 {target_sheet}，但在表中没找到 行[{row_code}] 或 列[({col_code})]")
                    else:
                        intercept_count += 1
                        if intercept_count <= 10:
                            self.log(f"🛑 [拦截] A={wrt_a}: 目标表 [{target_sheet}] 不存在，保持原样。")
                
                new_data.append(cur_row)

            output_path = os.path.splitext(send_file)[0] + "_updated.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log(f"\n✅ 全部处理完成！共生成 {len(new_data)} 行。")
            self.log(f"🎯 成功动态计算并更新了 {calc_success} 个新坐标！")
            self.log(f"🛑 有 {intercept_count} 行因为表不存在被拦截。")
            self.log(f"保留原有 SEND 结构，文件已保存至：{os.path.basename(output_path)}")

        except Exception as e:
            self.log(f"❌ 发生错误: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始同步并计算新坐标")

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()
