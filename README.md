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

# ================= 核心配置区域 =================
# 仅针对以下数字的表单进行坐标更新，其他表单保持原样
ALLOWED_SHEET_NUMS = ['5', '51', '6', '7', '52', '53', '9', '55', '11', '56', '12', '58', '19', '21', '33']
# ===============================================

def read_any_file(path):
    if not path or not os.path.exists(path):
        return None
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
    """锚点半径扫描：只在“行番号”周边寻找坐标，杜绝抓取远处的错误数据"""
    index = {}
    for fpath in file_paths:
        if not fpath or not os.path.exists(fpath):
            continue
        log_callback(f"\n▶ 正在扫描指定表单: {os.path.basename(fpath)}")
        try:
            wb = load_workbook(fpath, data_only=True)
            for ws in wb.worksheets:
                # 【修复个位数表过滤 Bug】: 将 05 转为 5，再与白名单比对
                sheet_num_str = re.sub(r'\D', '', ws.title)
                if not sheet_num_str: continue
                clean_num = str(int(sheet_num_str)) 
                if clean_num not in ALLOWED_SHEET_NUMS:
                    continue

                sheet_name = ws.title
                index[sheet_name] = {'rows': {}, 'cols': {}}

                # 1. 寻找核心锚点：“行番号”
                anchors = []
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            val = unicodedata.normalize('NFKC', str(cell.value)).replace(" ", "")
                            if "行番号" in val:
                                anchors.append(cell)

                # 2. 在核心锚点周边划定半径，提取真实坐标
                for anchor in anchors:
                    # A. 找列（表头）：严格锁定在行番号的上方 5 行到下方 2 行之间，向右查找
                    for r in range(max(1, anchor.row - 5), min(ws.max_row + 1, anchor.row + 3)):
                        for c in range(anchor.column, min(ws.max_column + 1, anchor.column + 150)):
                            cell = ws.cell(row=r, column=c)
                            if cell.value is not None:
                                v_str = unicodedata.normalize('NFKC', str(cell.value)).strip()
                                # 兼容 (1), 1, <1>, （1）
                                m = re.match(r'^[\(（<\[]?\s*(\d+)\s*[\)）>\]]?$', v_str)
                                if m:
                                    col_num = m.group(1)
                                    if col_num not in index[sheet_name]['cols']:
                                        index[sheet_name]['cols'][col_num] = cell.column_letter

                    # B. 找行（行号）：严格锁定在行番号的正下方，左右极窄走廊
                    for r in range(anchor.row + 1, min(ws.max_row + 1, anchor.row + 200)):
                        digits = []
                        for c in range(max(1, anchor.column - 2), min(ws.max_column + 1, anchor.column + 4)):
                            v = str(ws.cell(row=r, column=c).value or "")
                            for char in v:
                                if char.isdigit(): digits.append(char)
                        
                        if 2 <= len(digits) <= 4:
                            row_code = "".join(digits[:2])
                            if row_code not in index[sheet_name]['rows']:
                                index[sheet_name]['rows'][row_code] = r

                if index[sheet_name]['rows'] or index[sheet_name]['cols']:
                    log_callback(f"  ✅ 已锁定并解析: {sheet_name} (提取到 {len(index[sheet_name]['rows'])} 行, {len(index[sheet_name]['cols'])} 列)")

            wb.close()
        except Exception as e:
            log_callback(f"❌ 扫描文件 {os.path.basename(fpath)} 失败: {e}")
    return index


def get_smart_sheet_candidates(a_str):
    if len(a_str) < 5: return []
    # 动态前缀提取：无论 50101 还是 110412，砍掉最后四位坐标码，剩下的绝对是表号
    prefix = a_str[:-4] 
    return [f"表{prefix}", f"表0{prefix}", f"第{prefix}表", f"表{prefix}(2)", f"第{prefix}表(2)"]


def get_row_code(a_str):
    if len(a_str) < 4: return ""
    return a_str[-4:-2]


def get_col_code(a_str):
    """将 01 转换为 1，去精准匹配表头里的 1 或 (1)"""
    if len(a_str) < 2: return ""
    last_two = a_str[-2:]
    if last_two.startswith('0') and last_two[1].isdigit():
        return last_two[1]
    return last_two


class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei 智能坐标同步工具 v5.0 (雷达锁定纯净版)")
        root.geometry("920x720")

        self.wrt_path = tk.StringVar()
        self.send_path = tk.StringVar()
        self.file5100_1_path = tk.StringVar()
        self.file5100_2_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 坐标动态定位与同步工具", font=("Microsoft YaHei", 16, "bold")).pack(pady=10)
        tk.Label(self.root, text=f"当前已锁定处理的表单: {', '.join(ALLOWED_SHEET_NUMS)}", fg="#7f8c8d").pack(pady=2)

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

        self.log_text = scrolledtext.ScrolledText(self.root, height=24, width=110, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=5, fill="both", expand=True)

    def _add_row(self, p, txt, var, cmd, r):
        tk.Label(p, text=txt, width=20, anchor="e").grid(row=r, column=0, pady=5)
        tk.Entry(p, textvariable=var, width=62).grid(row=r, column=1, padx=5)
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
            
            max_cols = max(len(df_send.columns), 4)

            send_dict = {}
            for _, row in df_send.iterrows():
                b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if b_val:
                    row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                    if len(row_list) < max_cols:
                        row_list += [""] * (max_cols - len(row_list))
                    send_dict[b_val] = row_list

            self.log("=== 开始执行指定范围的锚点扫描 ===")
            index_5100 = build_5100_index([f5100_1, f5100_2], self.log)
            self.log("\n=== 扫描完成，开始精确匹配坐标 ===")

            new_data = []
            success_count = 0
            skip_count = 0
            fail_count = 0

            for _, wrt_row in df_wrt.iterrows():
                a = str(wrt_row.iloc[0]).strip() if pd.notna(wrt_row.iloc[0]) else ""
                b = str(wrt_row.iloc[1]).strip() if pd.notna(wrt_row.iloc[1]) else ""

                if not b: continue

                # 无条件继承旧 SEND 整行
                cur_row = send_dict.get(b, [""] * max_cols).copy()
                cur_row[0] = a
                cur_row[1] = b

                if len(a) >= 5 and a.isdigit():
                    # 动态截取表号（砍掉最后四位）
                    raw_sheet_num = a[:-4]
                    
                    if raw_sheet_num in ALLOWED_SHEET_NUMS:
                        candidates = get_smart_sheet_candidates(a)
                        target_sheet = next((s for s in candidates if s in index_5100), None)

                        if target_sheet:
                            cur_row[2] = target_sheet
                            row_code = get_row_code(a)
                            col_code = get_col_code(a)

                            actual_row = index_5100[target_sheet]['rows'].get(row_code)
                            actual_col = index_5100[target_sheet]['cols'].get(col_code)

                            if actual_row and actual_col:
                                cur_row[3] = f"{actual_col}{actual_row}"
                                success_count += 1
                            else:
                                fail_count += 1
                                if fail_count <= 15: 
                                    miss_parts = []
                                    if not actual_row: miss_parts.append(f"行[{row_code}]")
                                    if not actual_col: miss_parts.append(f"列[{col_code}]")
                                    self.log(f"⚠️ A={a}: 表[{target_sheet}] 内缺失 {' 和 '.join(miss_parts)}")
                        else:
                            fail_count += 1
                    else:
                        skip_count += 1

                new_data.append(cur_row)

            output_path = os.path.splitext(send_file)[0] + "_updated.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log(f"\n✅ 处理完成！")
            self.log(f"   ▶ 成功更新准确坐标: {success_count} 条")
            self.log(f"   ▶ 因不在名单跳过保持原样: {skip_count} 条")
            self.log(f"   ▶ 因缺失坐标保留原样: {fail_count} 条")
            self.log(f"文件已保存至: {output_path}")

        except Exception as e:
            self.log(f"❌ 发生异常: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始同步并计算新坐标")


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()
