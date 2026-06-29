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
    if not path or not os.path.exists(path):
        return None
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        for enc in ['utf-8-sig', 'cp932', 'shift_jis', 'gbk', 'utf-8']:
            try:
                df = pd.read_csv(path, header=None, encoding=enc, dtype=str, low_memory=False)
                if not df.empty:
                    return df
            except Exception:
                continue
        return pd.read_csv(path, header=None, dtype=str, low_memory=False)
    else:
        return pd.read_excel(path, header=None, dtype=str)


def build_5100_index(file_paths, log_callback):
    index = {}
    for fpath in file_paths:
        if not fpath or not os.path.exists(fpath):
            continue
        log_callback(f"\n▶ 正在启动纯净扫描: {os.path.basename(fpath)}")
        try:
            wb = load_workbook(fpath, data_only=True)
            for ws in wb.worksheets:
                sheet_name = ws.title
                index[sheet_name] = {'rows': {}, 'cols': {}}

                row_anchors = []
                col_anchors = []

                # 1. 寻找所有的“行番号”和表头“(1)”锚点
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            val = unicodedata.normalize('NFKC', str(cell.value)).replace(" ", "")
                            if "行番号" in val:
                                row_anchors.append(cell)
                            
                            # 完美匹配你说的 (1), (25) 这种带括号的列格式
                            m = re.search(r'[\(（](\d+)[\)）]', val)
                            if m:
                                col_anchors.append((m.group(1), cell))

                # 2. 提取行号映射（纯净数字过滤器法）
                for anchor in row_anchors:
                    for r_idx in range(anchor.row + 1, min(ws.max_row + 1, anchor.row + 200)):
                        digits = []
                        # 只在行番号正下方及右侧3列内寻找
                        start_col = max(1, anchor.column - 1)
                        end_col = min(ws.max_column + 1, anchor.column + 3)
                        
                        for c_idx in range(start_col, end_col):
                            v = ws.cell(row=r_idx, column=c_idx).value
                            if v is not None:
                                v_str = unicodedata.normalize('NFKC', str(v)).strip()
                                
                                # 【核心修复：最严过滤器】
                                # 彻底拒绝 "10万円"、"市" 等任何带文字的单元格
                                # 只有完全由数字、冒号、空格组成的单元格（如 "0:1:0" 或 "0"）才允许被提取！
                                if re.match(r'^[\d\s:：\.]+$', v_str):
                                    for char in v_str:
                                        if char.isdigit():
                                            digits.append(char)
                        
                        # 提取行号（前两位），如 0,1,0 提取出 01
                        if len(digits) >= 2 and len(digits) <= 4:
                            row_code = "".join(digits[:2])
                            if row_code not in index[sheet_name]['rows']:
                                index[sheet_name]['rows'][row_code] = r_idx

                # 3. 提取列号映射
                for num, cell in col_anchors:
                    if num not in index[sheet_name]['cols']:
                        index[sheet_name]['cols'][num] = cell.column_letter

                # 打印日志核对
                if index[sheet_name]['rows'] or index[sheet_name]['cols']:
                    log_callback(f"  👉 {sheet_name}: 提取到 {len(index[sheet_name]['rows'])} 个行, {len(index[sheet_name]['cols'])} 个列")
                    if index[sheet_name]['rows']:
                        sample_r = sorted(list(index[sheet_name]['rows'].keys()))[:6]
                        log_callback(f"      纯净行号前6个: {sample_r}")

            wb.close()
        except Exception as e:
            log_callback(f"❌ 扫描文件 {os.path.basename(fpath)} 时发生致命错误: {e}")
    return index


def get_smart_sheet_candidates(a_str):
    candidates = []
    if len(a_str) >= 6:
        prefix = a_str[:2]
    else:
        prefix = a_str[0]
    candidates.extend([f"表{prefix}", f"表0{prefix}", f"第{prefix}表", f"表{prefix}(2)", f"第{prefix}表(2)"])
    return list(dict.fromkeys(candidates))


def get_row_code(a_str):
    if len(a_str) < 4: return ""
    return a_str[-4:-2]


def get_col_code(a_str):
    """完美切合你的思路：将 01 转换为 1，去匹配表头里的 (1)"""
    if len(a_str) < 2: return ""
    last_two = a_str[-2:]
    if last_two.startswith('0') and last_two[1].isdigit():
        return last_two[1]
    return last_two


class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei 智能坐标同步工具 v4.0 (纯净提纯版)")
        root.geometry("900x700")

        self.wrt_path = tk.StringVar()
        self.send_path = tk.StringVar()
        self.file5100_1_path = tk.StringVar()
        self.file5100_2_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 坐标动态定位与同步工具", font=("Microsoft YaHei", 16, "bold")).pack(pady=10)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=5, fill="x")

        self._add_row(frame, "1. WRT 映射文件:", self.wrt_path, self.sel_wrt, 0)
        self._add_row(frame, "2. SEND 旧文件:", self.send_path, self.sel_send, 1)
        self._add_row(frame, "3. 最新 5100 文件①:", self.file5100_1_path, self.sel_5100_1, 2)
        self._add_row(frame, "4. 最新 5100 文件②(可选):", self.file5100_2_path, self.sel_5100_2, 3)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始同步并计算新坐标", command=self.transform,
                                 bg="#e74c3c", fg="white", font=("Microsoft YaHei", 12, "bold"), width=26)
        self.run_btn.pack(side="left", padx=10)

        self.log_text = scrolledtext.ScrolledText(self.root, height=26, width=110, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=5, fill="both", expand=True)

    def _add_row(self, p, txt, var, cmd, r):
        tk.Label(p, text=txt, width=20, anchor="e").grid(row=r, column=0, pady=5)
        tk.Entry(p, textvariable=var, width=60).grid(row=r, column=1, padx=5)
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
        wrt_file = self.wrt_path.get().strip()
        send_file = self.send_path.get().strip()
        f5100_1 = self.file5100_1_path.get().strip()
        f5100_2 = self.file5100_2_path.get().strip()

        if not wrt_file or not send_file:
            return messagebox.showerror("错误", "必须选择 WRT 和 SEND 文件")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            df_wrt = read_any_file(wrt_file)
            df_send = read_any_file(send_file)
            
            # 【修复列结构】严格按 A(ID), B(Key), C(Sheet), D(Coord) 输出
            max_cols = max(len(df_send.columns), 4)

            send_dict = {}
            for _, row in df_send.iterrows():
                b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if b_val:
                    row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                    if len(row_list) < max_cols:
                        row_list += [""] * (max_cols - len(row_list))
                    send_dict[b_val] = row_list

            self.log("=== 开始执行高精度数字纯净扫描 ===")
            index_5100 = build_5100_index([f5100_1, f5100_2], self.log)
            self.log("\n=== 扫描完成，开始计算坐标 ===")

            new_data = []
            success_count = 0
            fail_count = 0

            for _, wrt_row in df_wrt.iterrows():
                a = str(wrt_row.iloc[0]).strip() if pd.notna(wrt_row.iloc[0]) else ""
                b = str(wrt_row.iloc[1]).strip() if pd.notna(wrt_row.iloc[1]) else ""

                if not b:
                    continue

                # 继承旧 SEND，严格按照四列格式写入
                cur_row = send_dict.get(b, [""] * max_cols).copy()
                cur_row[0] = a
                cur_row[1] = b

                if len(a) >= 5 and a.isdigit():
                    candidates = get_smart_sheet_candidates(a)
                    target_sheet = next((s for s in candidates if s in index_5100), None)

                    if target_sheet:
                        cur_row[2] = target_sheet  # C列严格写入 Sheet名
                        row_code = get_row_code(a)
                        col_code = get_col_code(a)

                        actual_row = index_5100[target_sheet]['rows'].get(row_code)
                        actual_col = index_5100[target_sheet]['cols'].get(col_code)

                        if actual_row and actual_col:
                            cur_row[3] = f"{actual_col}{actual_row}" # D列严格写入坐标
                            success_count += 1
                        else:
                            fail_count += 1
                            if fail_count <= 20: 
                                miss_parts = []
                                if not actual_row: miss_parts.append(f"行号[{row_code}]")
                                if not actual_col: miss_parts.append(f"列号[({col_code})]")
                                miss_str = " 和 ".join(miss_parts)
                                self.log(f"⚠️ A={a}: 表[{target_sheet}] 内缺失 {miss_str}")
                    else:
                        fail_count += 1

                new_data.append(cur_row)

            output_path = os.path.splitext(send_file)[0] + "_updated.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log(f"\n✅ 处理完成！成功精准生成坐标: {success_count} 条，失败: {fail_count} 条。")
            self.log(f"文件已保存至: {output_path}")

        except Exception as e:
            self.log(f"❌ 发生异常: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始同步并计算新坐标")


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()
