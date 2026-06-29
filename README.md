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
    """全域自适应扫描 5100 文件"""
    index = {}
    for fpath in file_paths:
        if not fpath or not os.path.exists(fpath):
            continue
        log_callback(f"正在全域自适应扫描: {os.path.basename(fpath)} ...")
        try:
            wb = load_workbook(fpath, data_only=True)
            for ws in wb.worksheets:
                sheet_name = ws.title
                index[sheet_name] = {'rows': {}, 'cols': {}}

                row_anchor = None
                col_anchors = []

                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value:
                            val = unicodedata.normalize('NFKC', str(cell.value)).replace(" ", "")
                            if "行番号" in val:
                                row_anchor = cell
                            m = re.search(r'[\(（](\d+)[\)）]', val)
                            if m:
                                col_anchors.append((m.group(1), cell))

                if row_anchor:
                    for r_idx in range(row_anchor.row + 1, ws.max_row + 1):
                        digits = []
                        for c_idx in range(max(1, row_anchor.column - 3), min(ws.max_column, row_anchor.column + 15)):
                            v = str(ws.cell(row=r_idx, column=c_idx).value or "")
                            for char in v:
                                if char.isdigit():
                                    digits.append(char)
                        if len(digits) >= 2:
                            row_code = "".join(digits[:2])
                            index[sheet_name]['rows'][row_code] = r_idx

                for num, cell in col_anchors:
                    index[sheet_name]['cols'][num] = cell.column_letter
            wb.close()
        except Exception as e:
            log_callback(f"❌ 扫描错误: {e}")
    return index


def get_smart_sheet_candidates(a_str):
    """智能生成 Sheet 候选名称"""
    candidates = []
    if len(a_str) >= 6:
        prefix2 = a_str[:2]
        candidates.append(f"表{prefix2}")
        candidates.append(f"表0{prefix2}" if len(prefix2) == 1 else f"表{prefix2}")
    else:
        prefix1 = a_str[0]
        candidates.append(f"表{prefix1}")
        candidates.append(f"表0{prefix1}")
        candidates.append(f"第{prefix1}表")
    return list(dict.fromkeys(candidates))  # 去重


class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei 智能坐标同步工具 v3.5 (改进版)")
        root.geometry("860x650")
        self.wrt_path = tk.StringVar()
        self.send_path = tk.StringVar()
        self.file5100_1_path = tk.StringVar()
        self.file5100_2_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 坐标动态定位与同步工具 v3.5", font=("Microsoft YaHei", 15, "bold")).pack(pady=10)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=5, fill="x")

        self._add_row(frame, "1. WRT 映射文件:", self.wrt_path, self.sel_wrt, 0)
        self._add_row(frame, "2. SEND 旧文件:", self.send_path, self.sel_send, 1)
        self._add_row(frame, "3. 最新 5100 文件①:", self.file5100_1_path, self.sel_5100_1, 2)
        self._add_row(frame, "4. 最新 5100 文件②(可选):", self.file5100_2_path, self.sel_5100_2, 3)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始同步并计算新坐标", command=self.transform,
                                 bg="#27ae60", fg="white", font=("Microsoft YaHei", 12, "bold"), width=26)
        self.run_btn.pack(side="left", padx=10)

        self.log_text = scrolledtext.ScrolledText(self.root, height=22, width=100, font=("Consolas", 9), bg="#f8f9fa")
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
        wrt_file = self.wrt_path.get().strip()
        send_file = self.send_path.get().strip()
        f5100_1 = self.file5100_1_path.get().strip()
        f5100_2 = self.file5100_2_path.get().strip()

        if not wrt_file or not send_file:
            return messagebox.showerror("错误", "请选择 WRT 和 SEND 文件")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            df_wrt = read_any_file(wrt_file)
            df_send = read_any_file(send_file)
            max_cols = max(len(df_send.columns), 5)

            send_dict = {}
            for _, row in df_send.iterrows():
                b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if b_val:
                    row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                    if len(row_list) < max_cols:
                        row_list += [""] * (max_cols - len(row_list))
                    send_dict[b_val] = row_list

            self.log("=== 启动全域自适应坐标扫描 ===")
            index_5100 = build_5100_index([f5100_1, f5100_2], self.log)

            new_data = []
            for _, wrt_row in df_wrt.iterrows():
                a = str(wrt_row.iloc[0]).strip() if pd.notna(wrt_row.iloc[0]) else ""
                b = str(wrt_row.iloc[1]).strip() if pd.notna(wrt_row.iloc[1]) else ""
                c = str(wrt_row.iloc[2]).strip() if pd.notna(wrt_row.iloc[2]) else ""

                if not b:
                    continue

                cur_row = send_dict.get(b, [""] * max_cols).copy()
                cur_row[0] = a
                cur_row[1] = b
                cur_row[2] = c

                # === 智能 Sheet 匹配 ===
                if len(a) >= 5 and a.isdigit():
                    candidates = get_smart_sheet_candidates(a)
                    target_sheet = None
                    for cand in candidates:
                        if cand in index_5100:
                            target_sheet = cand
                            break

                    if target_sheet:
                        cur_row[3] = target_sheet
                        self.log(f"✓ Sheet 匹配成功: {a} -> {target_sheet}")

                        # 尝试计算 E 列坐标
                        try:
                            col_code = a[-2:]
                            row_code = a[-4:-2]
                            rows_dict = index_5100[target_sheet]['rows']
                            cols_dict = index_5100[target_sheet]['cols']

                            actual_row = rows_dict.get(row_code)
                            actual_col = cols_dict.get(col_code)

                            if actual_row and actual_col:
                                cur_row[4] = f"{actual_col}{actual_row}"
                                self.log(f"  → 新坐标: {cur_row[4]}")
                            else:
                                self.log(f"  ⚠️ 行/列未找到 -> row_code={row_code}, col_code={col_code}")
                        except Exception as e:
                            self.log(f"  ❌ 计算坐标出错: {e}")
                    else:
                        self.log(f"⚠️ 未找到匹配的 Sheet: {a}，尝试过: {candidates}")

                new_data.append(cur_row)

            output_path = os.path.splitext(send_file)[0] + "_updated.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log(f"\n✅ 处理完成！文件已保存: {output_path}")
            self.log(f"总行数: {len(new_data)}")

        except Exception as e:
            self.log(f"❌ 发生错误: {e}")
        finally:
            self.run_btn.config(state="normal", text="开始同步并计算新坐标")


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()