#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
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
                if not df.empty:
                    return df
            except Exception:
                continue
        return pd.read_csv(path, header=None, dtype=str, low_memory=False)
    else:
        return pd.read_excel(path, header=None, dtype=str)

class DeleteSendRowApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei Send 单文件精准删除工具 v1.0")
        root.geometry("800x550")

        self.csv_path = tk.StringVar()
        self.target_sheet = tk.StringVar(value="表11")   
        self.delete_cols = tk.StringVar(value="30")        

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 单文件精准删除工具", font=("Microsoft YaHei", 16, "bold"), fg="#c0392b").pack(pady=10)
        tk.Label(self.root, text="逻辑：在选定文件内，查找指定表名，若A列末两位命中指定数字，则直接删除该行。", fg="#7f8c8d").pack(pady=2)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=10, fill="x")

        # 1. 选择文件
        tk.Label(frame, text="1. 选择待处理文件 (CSV):", width=22, anchor="e").grid(row=0, column=0, pady=10)
        tk.Entry(frame, textvariable=self.csv_path, width=45).grid(row=0, column=1, padx=5, sticky="w")
        tk.Button(frame, text="浏览...", command=self.sel_file).grid(row=0, column=2, sticky="w")

        # 2. 指定表名
        tk.Label(frame, text="2. 目标工作表 (如: 表11):", width=22, anchor="e", fg="#2980b9", font=("Microsoft YaHei", 10, "bold")).grid(row=1, column=0, pady=10)
        tk.Entry(frame, textvariable=self.target_sheet, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=1, column=1, sticky="w", padx=5)

        # 3. 指定删除尾数
        tk.Label(frame, text="3. 删除指定尾数 (A列后2位):", width=24, anchor="e", fg="#8e44ad", font=("Microsoft YaHei", 10, "bold")).grid(row=2, column=0, pady=10)
        tk.Entry(frame, textvariable=self.delete_cols, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=2, column=1, sticky="w", padx=5)
        tk.Label(frame, text="← 可填多个，如 30 或 30,31", fg="#7f8c8d").grid(row=2, column=2, sticky="w")

        # 按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="执行精准删除", command=self.process_file,
                                 bg="#e74c3c", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25)
        self.run_btn.pack()

        # 日志区域
        self.log_text = scrolledtext.ScrolledText(self.root, height=14, width=95, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=10, fill="both", expand=True)

    def sel_file(self): self.csv_path.set(filedialog.askopenfilename(title="选择待处理文件"))

    def log(self, msg):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def process_file(self):
        file_path = self.csv_path.get().strip()
        target = self.target_sheet.get().strip()
        del_str = self.delete_cols.get().strip()

        # 解析强制删除名单 (支持逗号分隔多个)
        del_list = []
        if del_str:
            del_list = [x.strip().zfill(2) for x in del_str.replace('，', ',').split(',') if x.strip().isdigit()]

        if not file_path:
            return messagebox.showerror("错误", "请选择待处理的 CSV 文件！")
        if not target:
            return messagebox.showerror("错误", "请填写目标表名（例如：表11）！")
        if not del_list:
            return messagebox.showerror("错误", "请填写需要删除的A列末两位数字！")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            self.log(f"正在加载文件: {os.path.basename(file_path)} ...")
            df_send = read_any_file(file_path)

            if df_send is None:
                raise ValueError("文件读取失败或内容为空。")

            self.log(f"\n开始检查数据，锁定表单: 【{target}】")
            self.log(f"💣 开启狙击模式: 遇到尾数为 {del_list} 的行将直接删除！")
            
            new_data = []
            deleted_count = 0
            kept_count = 0
            skipped_count = 0

            for idx, row in df_send.iterrows():
                row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                a_val = row_list[0].strip().split('.')[0] if len(row_list) > 0 else ""

                # 1. 判定是否为目标工作表
                is_target_sheet = False
                if len(row_list) > 3 and target in str(row_list[3]).strip():
                    is_target_sheet = True
                
                # 双重保险：通过 A 列前缀判定表名
                if not is_target_sheet and len(a_val) >= 5 and a_val.isdigit():
                    prefix = a_val[:-4]
                    if target == f"表{prefix}" or target == f"表0{prefix}":
                        is_target_sheet = True

                # 2. 核心判定业务逻辑
                if is_target_sheet:
                    col_suffix = a_val[-2:] if (len(a_val) >= 2 and a_val[-2:].isdigit()) else ""
                    
                    if col_suffix in del_list:
                        deleted_count += 1
                        if deleted_count <= 20:
                            self.log(f"✂️ 成功击杀: A={a_val} (命中删除名单)")
                        elif deleted_count == 21:
                            self.log(f"✂️ ... (后续删除日志已折叠，避免刷屏)")
                    else:
                        new_data.append(row_list)
                        kept_count += 1
                else:
                    # 非目标表，无条件放行
                    new_data.append(row_list)
                    skipped_count += 1

            # 3. 输出过滤后的新文件
            output_path = os.path.splitext(file_path)[0] + f"_Deleted_{target}.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log("\n================ 狙击完成 ================")
            self.log(f"✅ 目标范围【{target}】内的操作结果：")
            self.log(f"   💣 命中规则并被删除: {deleted_count} 条")
            self.log(f"   ▶ 目标表内安全保留: {kept_count} 条")
            self.log(f"   ▶ 非目标表，跳过并保留: {skipped_count} 条")
            self.log(f"\n💾 处理后的新文件已安全保存至:\n   {output_path}")

        except Exception as e:
            self.log(f"❌ 发生异常: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="执行精准删除")

if __name__ == "__main__":
    root = tk.Tk()
    app = DeleteSendRowApp(root)
    root.mainloop()
