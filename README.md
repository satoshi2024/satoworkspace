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
        log_callback(f"正在强力扫描建立索引: {os.path.basename(fpath)} ... (请稍候)")
        
        try:
            wb = load_workbook(fpath, data_only=True)
        except Exception as e:
            log_callback(f"❌ 无法读取文件: {e}")
            continue

        for sheet_name in wb.sheetnames:
            if sheet_name not in index:
                index[sheet_name] = {'rows': {}, 'cols': {}}
            ws = wb[sheet_name]

            # 1. 扩大“行番号”列的搜索范围（应对合并单元格）
            xing_cols = set()
            for row in ws.iter_rows(min_row=1, max_row=50):
                for cell in row:
                    if cell.value:
                        # 强制转半角并去掉空格
                        val_str = unicodedata.normalize('NFKC', str(cell.value)).replace(" ", "")
                        if "行番号" in val_str:
                            # 包含自身及向右偏移的3列，防止 Excel 合并单元格导致数据藏在右边
                            xing_cols.update([cell.column, cell.column+1, cell.column+2, cell.column+3])
            
            # 2. 全表地毯式搜索坐标
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None: continue
                    
                    # 统一转换为半角，这是破解日文 Excel 格式的关键
                    raw_str = str(cell.value)
                    val_str = unicodedata.normalize('NFKC', raw_str)
                    
                    # --- 寻找列坐标: (1), (12), <12>, [12] (无视前后的多余文字或换行) ---
                    col_match = re.search(r'[\(\[<]\s*(\d+)\s*[\)\]>]', val_str)
                    if col_match:
                        col_num = str(int(col_match.group(1))) # 把 01 变成 1，12 变成 12
                        if col_num not in index[sheet_name]['cols']:
                            index[sheet_name]['cols'][col_num] = cell.column_letter

                    # --- 寻找行坐标: 0 1 1, 0:1:1, 011 ---
                    if cell.column in xing_cols:
                        # 暴力剔除一切空格、冒号、小数点等干扰符
                        clean_val = re.sub(r'[\s:：.-]', '', val_str)
                        if clean_val.isdigit():
                            # 将数字补齐为3位，例如 11 补齐为 011
                            padded = f"{int(clean_val):03d}"
                            if len(padded) == 3: # 确保是标准3位数
                                row_code = padded[:2] # 取前两位
                                # 只保留第一次找到的行号（最上方的）
                                if row_code not in index[sheet_name]['rows']:
                                    index[sheet_name]['rows'][row_code] = cell.row
        wb.close()
    return index

class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei 坐标动态计算工具 v3.1 (全角破解版)")
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
            
            # 建立旧 SEND 索引以保留多余列
            send_dict = {}
            for _, row in df_send.iterrows():
                b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if b_val and b_val not in send_dict:
                    lst = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                    lst.extend([""] * (max_cols - len(lst)))
                    send_dict[b_val] = lst

            # 构建 5100 坐标字典
            self.log("=== 开始破解并解析 5100 坐标结构 ===")
            index_5100 = build_5100_index([f5100_1, f5100_2], self.log)
            self.log("✅ 5100 坐标字典建立完成！\n")

            new_data = []
            calc_success = 0
            calc_fail = 0

            self.log("=== 开始根据 WRT 的 A 列计算最新坐标 ===")
            for _, wrt_row in df_wrt.iterrows():
                wrt_a = str(wrt_row.iloc[0]).strip() if pd.notna(wrt_row.iloc[0]) else ""
                wrt_b = str(wrt_row.iloc[1]).strip() if pd.notna(wrt_row.iloc[1]) else ""
                wrt_c = str(wrt_row.iloc[2]).strip() if pd.notna(wrt_row.iloc[2]) else ""
                if not wrt_b: continue
                
                # 继承结构
                cur_row = send_dict.get(wrt_b, [""] * max_cols).copy()
                cur_row[0], cur_row[1] = wrt_a, wrt_b
                if wrt_c: cur_row[2] = wrt_c

                # === 核心：根据 A 列计算 D 和 E ===
                if len(wrt_a) >= 5 and wrt_a.isdigit():
                    col_code = str(int(wrt_a[-2:]))  # 最后2位（列代码，如 12）
                    row_code = wrt_a[-4:-2]          # 中间2位（行代码，如 04）
                    sheet_num = wrt_a[:-4]           # 剩下的前缀（如 11 或 60）
                    target_sheet = f"表{sheet_num}"
                    
                    if target_sheet in index_5100:
                        cur_row[3] = target_sheet # 写入 D 列
                        
                        r = index_5100[target_sheet]['rows'].get(row_code)
                        c_letter = index_5100[target_sheet]['cols'].get(col_code)
                        
                        if r and c_letter:
                            # 成功计算出最新坐标！
                            new_coord = f"{c_letter}{r}"
                            cur_row[4] = new_coord
                            calc_success += 1
                        else:
                            calc_fail += 1
                            self.log(f"⚠️ [警告] A={wrt_a} ({target_sheet}): 找不到 行番号[{row_code}] 或 表头列[({col_code})]，保留了原坐标。")
                    else:
                        calc_fail += 1
                        self.log(f"🛑 [拦截] A={wrt_a}: 目标表 [{target_sheet}] 在5100文件中不存在！")
                
                new_data.append(cur_row)

            # 输出文件
            output_path = os.path.splitext(send_file)[0] + "_updated.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log(f"\n✅ 全部处理完成！共生成 {len(new_data)} 行。")
            self.log(f"🎯 成功动态计算并更新坐标 (E列)：{calc_success} 条")
            self.log(f"保留原有 SEND 结构，文件已保存至：{os.path.basename(output_path)}")

        except Exception as e:
            self.log(f"❌ 发生错误: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始同步并计算新坐标")

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()
