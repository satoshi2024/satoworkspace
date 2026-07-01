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

class CompareSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei Send 差异对比剔除工具 v6.0 (终极生杀版)")
        root.geometry("880x680")

        self.send1_path = tk.StringVar()
        self.send2_path = tk.StringVar()
        self.target_sheet = tk.StringVar(value="表11")   
        self.protect_start = tk.StringVar(value="30")   
        self.protect_end = tk.StringVar(value="36")     
        self.delete_col = tk.StringVar(value="")        # 新增：强制删除列号

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 差异对比剔除工具", font=("Microsoft YaHei", 16, "bold"), fg="#2980b9").pack(pady=10)
        tk.Label(self.root, text="逻辑：比对A列。支持【区间保护】和【强制删除】双重干预。", fg="#7f8c8d").pack(pady=2)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=10, fill="x")

        # 1. Send 1
        tk.Label(frame, text="1. Send 1 (寄存/现有):", width=24, anchor="e").grid(row=0, column=0, pady=8)
        tk.Entry(frame, textvariable=self.send1_path, width=48).grid(row=0, column=1, padx=5, sticky="w")
        tk.Button(frame, text="浏览...", command=self.sel_send1).grid(row=0, column=2, sticky="w")

        # 2. Send 2
        tk.Label(frame, text="2. Send 2 (新规/待过滤):", width=24, anchor="e").grid(row=1, column=0, pady=8)
        tk.Entry(frame, textvariable=self.send2_path, width=48).grid(row=1, column=1, padx=5, sticky="w")
        tk.Button(frame, text="浏览...", command=self.sel_send2).grid(row=1, column=2, sticky="w")

        # 3. 指定表名
        tk.Label(frame, text="3. 目标工作表 (如: 表11):", width=24, anchor="e", fg="#c0392b", font=("Microsoft YaHei", 10, "bold")).grid(row=2, column=0, pady=8)
        tk.Entry(frame, textvariable=self.target_sheet, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=2, column=1, sticky="w", padx=5)

        # 4. 保护起始列号
        tk.Label(frame, text="4. 保护起始列号 (A列末两位):", width=24, anchor="e", fg="#27ae60", font=("Microsoft YaHei", 10, "bold")).grid(row=3, column=0, pady=8)
        tk.Entry(frame, textvariable=self.protect_start, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=3, column=1, sticky="w", padx=5)
        tk.Label(frame, text="← 例: 30", fg="#7f8c8d").grid(row=3, column=2, sticky="w")

        # 5. 保护结束列号
        tk.Label(frame, text="5. 保护结束列号 (A列末两位):", width=24, anchor="e", fg="#27ae60", font=("Microsoft YaHei", 10, "bold")).grid(row=4, column=0, pady=8)
        tk.Entry(frame, textvariable=self.protect_end, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=4, column=1, sticky="w", padx=5)
        tk.Label(frame, text="← 例: 36 (即保留 30~36 区间)", fg="#7f8c8d").grid(row=4, column=2, sticky="w")

        # 6. 新增：强制删除列号
        tk.Label(frame, text="6. 强制删除列号 (可选):", width=24, anchor="e", fg="#8e44ad", font=("Microsoft YaHei", 10, "bold")).grid(row=5, column=0, pady=8)
        tk.Entry(frame, textvariable=self.delete_col, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=5, column=1, sticky="w", padx=5)
        tk.Label(frame, text="← 例: 30 或 30,31 (优先级最高)", fg="#7f8c8d").grid(row=5, column=2, sticky="w")

        # 按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始对比并剔除", command=self.process_files,
                                 bg="#e74c3c", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25)
        self.run_btn.pack()

        # 日志区域
        self.log_text = scrolledtext.ScrolledText(self.root, height=13, width=100, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=10, fill="both", expand=True)

    def sel_send1(self): self.send1_path.set(filedialog.askopenfilename(title="选择 Send 1 文件"))
    def sel_send2(self): self.send2_path.set(filedialog.askopenfilename(title="选择 Send 2 文件"))

    def log(self, msg):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def process_files(self):
        file1 = self.send1_path.get().strip()
        file2 = self.send2_path.get().strip()
        target = self.target_sheet.get().strip()
        start_str = self.protect_start.get().strip()
        end_str = self.protect_end.get().strip()
        del_str = self.delete_col.get().strip()

        # 解析区间范围
        p_start = int(start_str) if start_str.isdigit() else None
        p_end = int(end_str) if end_str.isdigit() else None

        # 解析强制删除名单 (支持逗号分隔多个，如 "30, 31")
        del_list = []
        if del_str:
            del_list = [x.strip().zfill(2) for x in del_str.replace('，', ',').split(',') if x.strip().isdigit()]

        if not file1 or not file2:
            return messagebox.showerror("错误", "请同时选择 Send 1 和 Send 2 文件！")
        if not target:
            return messagebox.showerror("错误", "请填写目标表名（例如：表11）！")
        if p_start is None or p_end is None:
            return messagebox.showerror("错误", "起始列号和结束列号必须填写有效的数字！")
        if p_start > p_end:
            return messagebox.showerror("错误", "起始列号不能大于结束列号！")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            self.log("正在加载 Send 1 和 Send 2 文件...")
            df_send1 = read_any_file(file1)
            df_send2 = read_any_file(file2)

            if df_send1 is None or df_send2 is None:
                raise ValueError("文件读取失败或内容为空。")

            self.log("正在提取 Send 1 (寄存) 的 A 列数据字典...")
            send1_a_set = set(str(x).strip().split('.')[0] for x in df_send1[0] if pd.notna(x))
            self.log(f"Send 1 共提取到 {len(send1_a_set)} 个唯一 A 列编号。")

            self.log(f"\n开始检查 Send 2，锁定表单: 【{target}】")
            self.log(f"🛡️ 区间保护: {p_start} ~ {p_end}")
            if del_list:
                self.log(f"💣 强制击杀: 尾数为 {del_list} 的行将无条件删除！")
            
            new_data = []
            kept_count = 0
            protected_count = 0
            explicit_deleted_count = 0
            normal_deleted_count = 0
            skipped_count = 0

            for idx, row in df_send2.iterrows():
                row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                a_val = row_list[0].strip().split('.')[0] if len(row_list) > 0 else ""

                # 判断是否为目标表
                is_target_sheet = False
                if len(row_list) > 3 and target in str(row_list[3]).strip():
                    is_target_sheet = True
                if not is_target_sheet and len(a_val) >= 5 and a_val.isdigit():
                    prefix = a_val[:-4]
                    if target == f"表{prefix}" or target == f"表0{prefix}":
                        is_target_sheet = True

                # ================= 核心判定业务逻辑 =================
                if is_target_sheet:
                    col_suffix = a_val[-2:] if (len(a_val) >= 2 and a_val[-2:].isdigit()) else ""
                    
                    # 场景 1: 【最高优先级】命中了用户指定的强制删除列表
                    if col_suffix in del_list:
                        explicit_deleted_count += 1
                        if explicit_deleted_count <= 15:
                            self.log(f"💣 强制击杀: A={a_val} (命中强制删除名单)")
                    else:
                        # 判断是否在保护区间
                        in_protected_range = False
                        if col_suffix:
                            col_num = int(col_suffix)
                            if p_start <= col_num <= p_end:
                                in_protected_range = True

                        # 场景 2: 处于保护区间 -> 强制留存
                        if in_protected_range:
                            new_data.append(row_list)
                            protected_count += 1
                            if protected_count <= 10:
                                self.log(f"🛡️ 保护留存: A={a_val} (在 {p_start}~{p_end} 区间)")
                                
                        # 场景 3: Send 1 里存在 -> 留存
                        elif a_val in send1_a_set:
                            new_data.append(row_list)
                            kept_count += 1
                            
                        # 场景 4: 既不在保护区间，Send 1 里又没有 -> 剔除
                        else:
                            normal_deleted_count += 1
                            if normal_deleted_count <= 10:
                                self.log(f"✂️ 剔除数据: A={a_val} (非区间内且寄存无此数据)")
                else:
                    # 非目标表，无条件完整放行
                    new_data.append(row_list)
                    skipped_count += 1

            output_path = os.path.splitext(file2)[0] + f"_Filtered_{target}.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log("\n================ 处理完成 ================")
            self.log(f"✅ 目标范围【{target}】内的对比结果：")
            self.log(f"   ▶ 基础比对匹配成功并保留: {kept_count} 条")
            self.log(f"   ▶ 处于区间 [{p_start}~{p_end}] 强制保护保留: {protected_count} 条")
            self.log(f"   💣 命中强制删除名单被击杀: {explicit_deleted_count} 条")
            self.log(f"   ▶ 常规比对未匹配并被剔除: {normal_deleted_count} 条")
            self.log(f"   ▶ 非目标表，无条件跳过并保留: {skipped_count} 条")
            self.log(f"\n💾 处理后的新规文件已安全保存至:\n   {output_path}")

        except Exception as e:
            self.log(f"❌ 发生异常: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始对比并剔除")

if __name__ == "__main__":
    root = tk.Tk()
    app = CompareSendApp(root)
    root.mainloop()
