#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from openpyxl import load_workbook
import os
from datetime import datetime

def read_any_file(path):
    if not path or not os.path.exists(path):
        return None
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        for enc in ['utf-8-sig', 'cp932', 'shift_jis', 'gbk', 'utf-8']:
            try:
                df = pd.read_csv(path, header=None, encoding=enc, dtype=str, low_memory=False)
                return df
            except Exception:
                continue
        return pd.read_csv(path, header=None, dtype=str, low_memory=False)
    else:
        return pd.read_excel(path, header=None, dtype=str)

def get_all_sheet_names(file_paths):
    all_sheets = set()
    for fpath in file_paths:
        if not fpath or not os.path.exists(fpath):
            continue
        try:
            wb = load_workbook(fpath, read_only=True, data_only=False)
            all_sheets.update(wb.sheetnames)
            wb.close()
        except Exception as e:
            print(f"警告: 无法读取 {fpath} 的 sheets: {e}")
    return all_sheets

def parse_a_to_sheet(a_val):
    """仅仅解析出目标 Sheet 名称，不做多余操作"""
    if pd.isna(a_val):
        return ""
    a_str = str(a_val).strip()
    if '.' in a_str:
        a_str = a_str.split('.')[0]
    if 'e' in a_str.lower():
        try:
            a_str = str(int(float(a_str)))
        except:
            pass
    if not a_str or not a_str.isdigit():
        return ""

    if len(a_str) >= 6:
        sheet_num = a_str[:2]
    else:
        sheet_num = a_str[0]
    return f"表{sheet_num}"


class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei Send 文件更新工具 v2.0 (精准修复版)")
        root.geometry("820x620")
        root.resizable(True, True)

        self.wrt_path = tk.StringVar()
        self.send_path = tk.StringVar()
        self.file5100_1_path = tk.StringVar()
        self.file5100_2_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(self.root, text="Send 文件自动同步更新工具", font=("Microsoft YaHei", 16, "bold"), fg="#2c3e50")
        title_label.pack(pady=10)

        desc = tk.Label(self.root, text="修复多余列丢失问题，严格执行5100校验拦截", font=("Microsoft YaHei", 10), fg="#7f8c8d")
        desc.pack(pady=5)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=10, fill="x")

        self._add_file_row(frame, "1. WRT 映射文件（母体）:", self.wrt_path, self.select_wrt, 0)
        self._add_file_row(frame, "2. SEND 文件（待更新）:", self.send_path, self.select_send, 1)
        self._add_file_row(frame, "3. 5100 文件①:", self.file5100_1_path, self.select_5100_1, 2)
        self._add_file_row(frame, "4. 5100 文件②（可选）:", self.file5100_2_path, self.select_5100_2, 3)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)

        self.transform_btn = tk.Button(btn_frame, text="开始变换 / 更新 SEND", command=self.transform, bg="#27ae60", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25, height=2, relief="raised")
        self.transform_btn.pack(side="left", padx=10)

        self.clear_btn = tk.Button(btn_frame, text="清空日志", command=self.clear_log, bg="#95a5a6", fg="white", font=("Microsoft YaHei", 11), width=12, height=2)
        self.clear_btn.pack(side="left", padx=10)

        log_label = tk.Label(self.root, text="处理日志：", font=("Microsoft YaHei", 10))
        log_label.pack(anchor="w", padx=20)

        self.log_text = scrolledtext.ScrolledText(self.root, height=18, width=95, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=5, fill="both", expand=True)

        self.status_var = tk.StringVar(value="就绪 | 请先选择文件，然后点击「开始变换」")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief="sunken", anchor="w", fg="#34495e")
        status_bar.pack(side="bottom", fill="x")

    def _add_file_row(self, parent, label_text, var, cmd, row):
        tk.Label(parent, text=label_text, font=("Microsoft YaHei", 10), width=22, anchor="e").grid(row=row, column=0, padx=5, pady=6, sticky="e")
        entry = tk.Entry(parent, textvariable=var, width=55, font=("Consolas", 9))
        entry.grid(row=row, column=1, padx=5, pady=6)
        btn = tk.Button(parent, text="浏览...", command=cmd, width=8, font=("Microsoft YaHei", 9))
        btn.grid(row=row, column=2, padx=5, pady=6)

    def select_wrt(self): path = filedialog.askopenfilename(title="选择 WRT", filetypes=[("CSV/Excel", "*.csv *.xlsx *.xlsm")]); self.wrt_path.set(path) if path else None
    def select_send(self): path = filedialog.askopenfilename(title="选择 SEND", filetypes=[("CSV/Excel", "*.csv *.xlsx *.xlsm")]); self.send_path.set(path) if path else None
    def select_5100_1(self): path = filedialog.askopenfilename(title="选择 5100①", filetypes=[("Excel", "*.xlsx *.xlsm")]); self.file5100_1_path.set(path) if path else None
    def select_5100_2(self): path = filedialog.askopenfilename(title="选择 5100②", filetypes=[("Excel", "*.xlsx *.xlsm")]); self.file5100_2_path.set(path) if path else None

    def log(self, msg):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def transform(self):
        wrt_file, send_file = self.wrt_path.get().strip(), self.send_path.get().strip()
        f5100_1, f5100_2 = self.file5100_1_path.get().strip(), self.file5100_2_path.get().strip()

        if not wrt_file or not send_file:
            messagebox.showerror("错误", "必须选择 WRT 文件 和 SEND 文件！")
            return

        self.transform_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)
        
        try:
            # 1. 读取数据
            df_wrt = read_any_file(wrt_file)
            df_send = read_any_file(send_file)
            all_sheets = get_all_sheet_names([f5100_1, f5100_2])
            
            # 【修复1】：获取原 SEND 的最大列数，确保没有任何一列丢失
            max_cols = max(len(df_send.columns), 5) 
            
            new_data = []
            c_preserved_count = 0
            invalid_sheet_count = 0

            # 为了快速在旧 SEND 查找 B 列，建立字典索引
            send_dict = {}
            for _, row in df_send.iterrows():
                b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if b_val and b_val not in send_dict:
                    # 将整行转为 list 存起来，补齐长度
                    row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                    if len(row_list) < max_cols:
                        row_list.extend([""] * (max_cols - len(row_list)))
                    send_dict[b_val] = row_list

            self.log("开始以 WRT 为母体重建数据并进行 5100 验证...")
            
            # 2. 遍历 WRT，严格重组数据
            for _, wrt_row in df_wrt.iterrows():
                wrt_a = str(wrt_row.iloc[0]).strip() if pd.notna(wrt_row.iloc[0]) else ""
                wrt_b = str(wrt_row.iloc[1]).strip() if pd.notna(wrt_row.iloc[1]) else ""
                wrt_c = str(wrt_row.iloc[2]).strip() if pd.notna(wrt_row.iloc[2]) else ""
                
                if not wrt_b: continue # 跳过没有 B 列的空行
                
                # 初始化新行（默认全是空字符串，长度保证覆盖原SEND所有列）
                current_row = [""] * max_cols
                
                # 如果旧 SEND 里有匹配项，完整继承旧数据（包括F/G/H等所有后续列）
                if wrt_b in send_dict:
                    current_row = send_dict[wrt_b].copy()
                
                # 更新 A 列和 B 列
                current_row[0] = wrt_a
                current_row[1] = wrt_b
                
                # 更新 C 列 (如果WRT有值用WRT，否则保留旧的)
                if wrt_c:
                    current_row[2] = wrt_c
                elif current_row[2]:
                    c_preserved_count += 1
                
                # 【修复2】：严格的 Sheet 拦截逻辑
                target_sheet = parse_a_to_sheet(wrt_a)
                
                if target_sheet:
                    # 如果用户选了 5100 文件，需要验证
                    if all_sheets:
                        if target_sheet in all_sheets:
                            current_row[3] = target_sheet
                        else:
                            # 拦截！目标 Sheet 不存在，不修改 D列，也不修改 E列
                            invalid_sheet_count += 1
                            self.log(f"⚠️ 拦截: B列[{wrt_b[:15]}] 对应的目标表[{target_sheet}]不存在! 已保留原样。")
                    else:
                        # 没选 5100 文件，直接更新
                        current_row[3] = target_sheet
                
                # 将处理好的一行加入结果
                new_data.append(current_row)

            # 3. 生成输出结果
            df_new = pd.DataFrame(new_data)
            
            output_path = os.path.splitext(send_file)[0] + "_updated.csv"
            df_new.to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log(f"\n✅ 处理完成！")
            self.log(f"👉 成功保留了 SEND 原有的 {max_cols} 列结构。")
            self.log(f"👉 总输出行数：{len(df_new)}")
            if invalid_sheet_count > 0:
                self.log(f"🛑 成功拦截了 {invalid_sheet_count} 次无效的 Sheet 更改，相关行的 D/E 坐标未受影响。")
            self.status_var.set(f"完成 | 已生成 {os.path.basename(output_path)}")

        except Exception as e:
            self.log(f"❌ 错误: {str(e)}")
            self.status_var.set("错误 | 请查看日志")
        finally:
            self.transform_btn.config(state="normal", text="开始变换 / 更新 SEND")

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()
