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
        root.title("DafKazei Send 差异对比剔除工具 v3.0 (保护新增列版)")
        root.geometry("850x600")

        self.send1_path = tk.StringVar()
        self.send2_path = tk.StringVar()
        self.target_sheet = tk.StringVar(value="表11")  # 默认表名
        self.protect_col = tk.StringVar(value="36")     # 默认受保护的列号(A列最后两位)

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Send 差异对比剔除工具", font=("Microsoft YaHei", 16, "bold"), fg="#2980b9").pack(pady=10)
        tk.Label(self.root, text="逻辑：在指定表内，对比 A 列。若输入保护列号(如36)，则 Send2 中结尾为36的编号无条件保留。", fg="#7f8c8d").pack(pady=2)

        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=10, fill="x")

        # 1. Send 1 选择
        tk.Label(frame, text="1. Send 1 (寄存/现有):", width=22, anchor="e").grid(row=0, column=0, pady=8)
        tk.Entry(frame, textvariable=self.send1_path, width=55).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="浏览...", command=self.sel_send1).grid(row=0, column=2)

        # 2. Send 2 选择
        tk.Label(frame, text="2. Send 2 (新规/待过滤):", width=22, anchor="e").grid(row=1, column=0, pady=8)
        tk.Entry(frame, textvariable=self.send2_path, width=55).grid(row=1, column=1, padx=5)
        tk.Button(frame, text="浏览...", command=self.sel_send2).grid(row=1, column=2)

        # 3. 指定表名
        tk.Label(frame, text="3. 目标工作表 (如: 表11):", width=22, anchor="e", fg="#c0392b", font=("Microsoft YaHei", 10, "bold")).grid(row=2, column=0, pady=8)
        tk.Entry(frame, textvariable=self.target_sheet, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=2, column=1, sticky="w", padx=5)

        # 4. 新增列保护 (A列最后两位)
        tk.Label(frame, text="4. 保护新增列号 (A列末两位):", width=24, anchor="e", fg="#27ae60", font=("Microsoft YaHei", 10, "bold")).grid(row=3, column=0, pady=8)
        tk.Entry(frame, textvariable=self.protect_col, width=20, font=("Microsoft YaHei", 10, "bold")).grid(row=3, column=1, sticky="w", padx=5)
        tk.Label(frame, text="例如输入 36，则 112136 会被保留", fg="#7f8c8d").grid(row=3, column=1, padx=180, sticky="w")

        # 按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="开始对比并剔除", command=self.process_files,
                                 bg="#e74c3c", fg="white", font=("Microsoft YaHei", 12, "bold"), width=25)
        self.run_btn.pack()

        # 日志区域
        self.log_text = scrolledtext.ScrolledText(self.root, height=15, width=100, font=("Consolas", 9), bg="#f8f9fa")
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
        protect_val = self.protect_col.get().strip()

        # 将保护列号格式化为两位数 (比如输入 6 变成 "06", 输入 36 还是 "36")
        if protect_val.isdigit():
            protect_val = protect_val.zfill(2)

        if not file1 or not file2:
            return messagebox.showerror("错误", "请同时选择 Send 1 和 Send 2 文件！")
        if not target:
            return messagebox.showerror("错误", "请填写目标表名（例如：表11）！")

        self.run_btn.config(state="disabled", text="处理中...")
        self.log_text.delete("1.0", tk.END)

        try:
            self.log("正在加载 Send 1 和 Send 2 文件...")
            df_send1 = read_any_file(file1)
            df_send2 = read_any_file(file2)

            if df_send1 is None or df_send2 is None:
                raise ValueError("文件读取失败或内容为空。")

            # 1. 提取 Send 1 的 A 列集合 (清理潜在的小数点)
            self.log("正在提取 Send 1 (寄存) 的 A 列数据字典...")
            send1_a_set = set(str(x).strip().split('.')[0] for x in df_send1[0] if pd.notna(x))
            self.log(f"Send 1 共提取到 {len(send1_a_set)} 个唯一 A 列编号。")

            # 2. 遍历 Send 2，进行精准过滤
            self.log(f"\n开始检查 Send 2，锁定表单: 【{target}】")
            if protect_val:
                self.log(f"🛡️ 开启新增列保护: A列最后两位为 【{protect_val}】 的数据将无条件保留！")
            
            new_data = []
            kept_count = 0
            protected_count = 0
            deleted_count = 0
            skipped_count = 0

            for idx, row in df_send2.iterrows():
                # 【修复核心：完全不去加空字符串补齐列，保持原汁原味】
                row_list = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                a_val = row_list[0].strip().split('.')[0] if len(row_list) > 0 else ""

                # 判断当前行是否属于用户指定的目标表
                is_target_sheet = False
                if len(row_list) > 3 and target in str(row_list[3]).strip():
                    is_target_sheet = True

                # 双重保险：通过 A 列前缀解析
                if not is_target_sheet and len(a_val) >= 5 and a_val.isdigit():
                    prefix = a_val[:-4]
                    if target == f"表{prefix}" or target == f"表0{prefix}":
                        is_target_sheet = True

                # ================= 业务逻辑：判定生死 =================
                if is_target_sheet:
                    # 场景 1: Send 1 里有，正常保留
                    if a_val in send1_a_set:
                        new_data.append(row_list)
                        kept_count += 1
                        
                    # 场景 2: Send 1 里没有，但它属于受保护的“新增列”！
                    elif protect_val and len(a_val) >= 2 and a_val[-2:] == protect_val:
                        new_data.append(row_list)
                        protected_count += 1
                        if protected_count <= 10:
                            self.log(f"🛡️ 成功保护新增列数据: A={a_val}")
                            
                    # 场景 3: Send 1 里没有，也不是新增列 -> 干掉它！
                    else:
                        deleted_count += 1
                        if deleted_count <= 10:
                            self.log(f"✂️ 剔除数据: A={a_val} (Send 1中无此数据)")
                else:
                    # 不是目标表，无条件保留，绝对不动原数据
                    new_data.append(row_list)
                    skipped_count += 1

            # 4. 输出过滤后的文件
            output_path = os.path.splitext(file2)[0] + f"_Filtered_{target}.csv"
            pd.DataFrame(new_data).to_csv(output_path, index=False, header=False, encoding='utf-8-sig')

            self.log("\n================ 处理完成 ================")
            self.log(f"✅ 目标范围【{target}】内的对比结果：")
            self.log(f"   ▶ 匹配成功并保留: {kept_count} 条")
            self.log(f"   ▶ 触发保护机制强制保留: {protected_count} 条")
            self.log(f"   ▶ 未匹配并被剔除: {deleted_count} 条")
            self.log(f"   ▶ 非目标表，跳过并保留: {skipped_count} 条")
            self.log(f"\n💾 剔除后的 Send 2 新文件已保存至:\n   {output_path}")

        except Exception as e:
            self.log(f"❌ 发生异常: {str(e)}")
        finally:
            self.run_btn.config(state="normal", text="开始对比并剔除")

if __name__ == "__main__":
    root = tk.Tk()
    app = CompareSendApp(root)
    root.mainloop()
