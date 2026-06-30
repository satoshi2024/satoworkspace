#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
from datetime import datetime

# ================= 核心配置区域 =================
# 你指定需要替换和更新的表单名单
TARGET_SHEET_NUMS = ['5', '51', '6', '7', '52', '53', '9', '55', '11', '56', '12', '58', '19', '21', '33']
# ===============================================

def read_any_file(path):
    """智能读取任意编码的 CSV 或 Excel 文件"""
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

def is_target_row(row):
    """智能判断当前行是否属于目标表单"""
    a_val = str(row[0]).strip() if pd.notna(row[0]) else ""
    
    # 方式1：通过 A 列前缀判断 (最准)
    if len(a_val) >= 5 and a_val.isdigit():
        prefix = a_val[:-4]
        if str(int(prefix)) in TARGET_SHEET_NUMS:
            return True
            
    # 方式2：通过 D 列的 Sheet 名判断 (双重保险)
    if len(row) > 3 and pd.notna(row[3]):
        d_val = str(row[3]).strip()
        for num in TARGET_SHEET_NUMS:
            if d_val == f"表{num}" or d_val == f"表0{num}" or d_val == f"第{num}表":
                return True
                
    return False

class MergeSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei Send 原位替换与合并工具 v2.0")
        root.geometry("850x650")

        self.send1_path = tk.StringVar()
        self.send2_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 数据原位替换与合并工具", font=("Microsoft YaHei", 16, "bold"), fg="#8e44ad").pack(pady=10)
        desc = "逻辑：在 Send1 中原位删除指定表，并在原位置无缝插入 Send2 的新数据。"
        tk.Label(self.root, text=desc, fg="#7f8c8d").pack(pady=2)
        tk.Label(self.root, text=f"当前目标表单: {', '.join(TARGET_SHEET_NUMS)}", fg="#c0392b", font=("Microsoft YaHei", 9, "bold")).pack(pady=2)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=10, fill="x")

        self._add_row(frame, "1. Send 1 (寄存 / 原文件):", self.send1_path, self.sel_send1, 0)
        self._add_row(frame, "2. Send 2 (新规 / 数据源):", self.send2_path, self.sel_send2, 1)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始原位替换", command=self.process_files,
                                 bg="#27ae60", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25)
        self.run_btn.pack()

        self.log_text = scrolledtext.ScrolledText(self.root, height=18, width=100, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=10, fill="both", expand=True)

    def _add_row(self, p, txt, var, cmd, r):
        tk.Label(p, text=txt, width=22, anchor="e").grid(row=r, column=0, pady=8)
        tk.Entry(p, textvariable=var, width=55).grid(row=r, column=1, padx=5)
        tk.Button(p, text="浏览...", command=cmd).grid(row=r, column=2)

    def sel_send1(self): self.send1_path.set(filedialog.askopenfilename(title="选择 Send 1 文件"))
    def sel_send2(self): self.send2_path.set(filedialog.askopenfilename(title="选择 Send 2 文件"))

    def log(self, msg):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def process_files(self):
        file1 = self.send1_path.get().strip()
        file2 = self.send2_path.get().strip()

        if not file1 or not file2:
            return messagebox.showerror("错误", "请同时选择 Send 1 和 Send 2 文件！")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            self.log("正在加载 Send 1 和 Send 2 文件...")
            df_send1 = read_any_file(file1)
            df_send2 = read_any_file(file2)

            if df_send1 is None or df_send2 is None:
                raise ValueError("文件读取失败或内容为空。")
            
            max_cols = max(len(df_send1.columns), len(df_send2.columns), 5)

            # ================= 步骤 1：从 Send 2 提取新数据 =================
            self.log("▶ 正在从 Send 2 提取新规数据...")
            send2_extracted = []
            for _, row in df_send2.iterrows():
                row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                row_list += [""] * (max_cols - len(row_list))
                if is_target_row(row_list):
                    send2_extracted.append(row_list)
                    
            self.log(f"  📥 成功提取了 {len(send2_extracted)} 条新目标表数据。")

            # ================= 步骤 2：在 Send 1 中进行原位替换 =================
            self.log("\n▶ 开始在 Send 1 中进行原位替换...")
            final_data = []
            inserted_new_data = False
            send1_deleted_count = 0
            
            for i, row in df_send1.iterrows():
                row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                row_list += [""] * (max_cols - len(row_list))
                
                if is_target_row(row_list):
                    send1_deleted_count += 1
                    # 【核心逻辑】：遇到需要删除的第一条旧数据时，立刻把所有新数据安插在这里！
                    if not inserted_new_data:
                        final_data.extend(send2_extracted)
                        inserted_new_data = True
                        self.log(f"  🎯 精准定位！已在原文件的第 {i+1} 行位置处无缝插入了所有新数据。")
                    # 旧数据不加入 final_data，即被物理删除
                else:
                    final_data.append(row_list)
                    
            # 兜底保障：如果 Send 1 里压根没有目标老数据，则补充在末尾
            if not inserted_new_data:
                final_data.extend(send2_extracted)
                self.log("  ⚠️ 注意：Send 1 中未发现目标旧数据，新数据已追加至文件末尾。")

            self.log(f"  ✂️ 总计从 Send 1 中清理了 {send1_deleted_count} 条旧目标表数据。")

            # ================= 步骤 3：保存文件 =================
            output_path = os.path.splitext(file1)[0] + "_Replaced.csv"
            pd.DataFrame(final_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log("\n================ 替换完美完成 ================")
            self.log(f"总计生成数据: {len(final_data)} 条。")
            self.log(f"💾 原位替换后的最终文件已保存至:\n   {output_path}")

        except Exception as e:
            self.log(f"❌ 发生异常: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始原位替换")

if __name__ == "__main__":
    root = tk.Tk()
    app = MergeSendApp(root)
    root.mainloop()
