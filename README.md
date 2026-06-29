#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Send文件更新工具 GUI
功能：
1. 选择 WRT 文件（映射母体）
2. 选择两个 5100 文件（用于验证 sheet 名，防止乱改）
3. 选择 SEND 文件
4. 点击“开始变换”，脚本会分两阶段处理：
   第一阶段（以 WRT 为准对齐）：
     - 严格以 WRT 为母体重建 SEND（该增则增、该删则删）
     - 更新 A/B/C 列，保留能保留的 D/E
   第二阶段（用 A 列更新坐标）：
     - 根据 A 列规则自动生成 D (sheet，如 表6、表15) 和 E (逻辑坐标)
     - 严格验证 D 列 sheet 是否存在于所选 5100 文件中
   输出：updated_send.csv（或 .xlsx）
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from openpyxl import load_workbook
import os
from datetime import datetime

def read_any_file(path):
    """智能读取 csv 或 xlsx/xlsm，自动尝试常见日文/中文编码"""
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
        # 最后尝试不指定编码
        return pd.read_csv(path, header=None, dtype=str, low_memory=False)
    else:
        # xlsx / xlsm
        return pd.read_excel(path, header=None, dtype=str)

def get_all_sheet_names(file_paths):
    """从一个或多个 5100 文件中收集所有 sheet 名"""
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


def parse_a_to_sheet_and_coord(a_val):
    """
    根据用户规则解析 A 列，生成 D (sheet)
    通用规则 + 健壮清洗（修复 float/.0 导致的6位数识别失败）：
    - 先把 A 清洗成纯数字字符串（去掉 .0、科学计数法等）
    - 6 位数字 → 取前两位（表12、表15 等）
    - 5 位数字 → 取第一位（表6、表7 等）
    - E 列暂时不覆盖
    """
    if pd.isna(a_val):
        return "", ""

    # 健壮清洗 A 值
    a_str = str(a_val).strip()
    if '.' in a_str:
        a_str = a_str.split('.')[0]
    if 'e' in a_str.lower():
        try:
            a_str = str(int(float(a_str)))
        except:
            pass

    if not a_str or not a_str.isdigit():
        return "", ""

    # === Sheet 名解析（通用规则） ===
    if len(a_str) == 6:
        sheet_num = a_str[:2]      # 6位 → 表12, 表15, 表11 等
    else:
        sheet_num = a_str[0]       # 5位或其他 → 表6, 表7 等
    sheet_name = f"表{sheet_num}"

    return sheet_name, ""

class UpdateSendApp:
    def __init__(self, root):
        self.root = root
        root.title("DafKazei Send 文件更新工具 v1.0")
        root.geometry("820x620")
        root.resizable(True, True)

        # 变量
        self.wrt_path = tk.StringVar()
        self.send_path = tk.StringVar()
        self.file5100_1_path = tk.StringVar()
        self.file5100_2_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="Send 文件自动同步更新工具", 
                               font=("Microsoft YaHei", 16, "bold"), fg="#2c3e50")
        title_label.pack(pady=10)

        # 说明
        desc = tk.Label(self.root, text="以 WRT 为母体同步 A/B/C 列，保留 D/E 列，严格按 sheet 名验证 5100 文件", 
                        font=("Microsoft YaHei", 10), fg="#7f8c8d")
        desc.pack(pady=5)

        # 文件选择区域
        frame = tk.Frame(self.root)
        frame.pack(padx=20, pady=10, fill="x")

        # WRT
        self._add_file_row(frame, "1. WRT 映射文件（母体）:", self.wrt_path, self.select_wrt, 0)

        # SEND
        self._add_file_row(frame, "2. SEND 文件（待更新）:", self.send_path, self.select_send, 1)

        # 5100-1
        self._add_file_row(frame, "3. 5100 文件①:", self.file5100_1_path, self.select_5100_1, 2)

        # 5100-2
        self._add_file_row(frame, "4. 5100 文件②（可选）:", self.file5100_2_path, self.select_5100_2, 3)

        # 按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)

        self.transform_btn = tk.Button(btn_frame, text="开始变换 / 更新 SEND", 
                                       command=self.transform,
                                       bg="#27ae60", fg="white", font=("Microsoft YaHei", 12, "bold"),
                                       width=25, height=2, relief="raised")
        self.transform_btn.pack(side="left", padx=10)

        self.clear_btn = tk.Button(btn_frame, text="清空日志", 
                                   command=self.clear_log,
                                   bg="#95a5a6", fg="white", font=("Microsoft YaHei", 11),
                                   width=12, height=2)
        self.clear_btn.pack(side="left", padx=10)

        # 日志区域
        log_label = tk.Label(self.root, text="处理日志：", font=("Microsoft YaHei", 10))
        log_label.pack(anchor="w", padx=20)

        self.log_text = scrolledtext.ScrolledText(self.root, height=18, width=95, 
                                                  font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(padx=20, pady=5, fill="both", expand=True)

        # 底部状态
        self.status_var = tk.StringVar(value="就绪 | 请先选择文件，然后点击「开始变换」")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                              bd=1, relief="sunken", anchor="w", fg="#34495e")
        status_bar.pack(side="bottom", fill="x")

    def _add_file_row(self, parent, label_text, var, cmd, row):
        tk.Label(parent, text=label_text, font=("Microsoft YaHei", 10), width=22, anchor="e").grid(row=row, column=0, padx=5, pady=6, sticky="e")
        entry = tk.Entry(parent, textvariable=var, width=55, font=("Consolas", 9))
        entry.grid(row=row, column=1, padx=5, pady=6)
        btn = tk.Button(parent, text="浏览...", command=cmd, width=8, font=("Microsoft YaHei", 9))
        btn.grid(row=row, column=2, padx=5, pady=6)

    def select_wrt(self):
        path = filedialog.askopenfilename(
            title="选择 WRT 文件",
            filetypes=[("CSV/Excel 文件", "*.csv *.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if path:
            self.wrt_path.set(path)
            self.log(f"已选择 WRT: {os.path.basename(path)}")

    def select_send(self):
        path = filedialog.askopenfilename(
            title="选择 SEND 文件",
            filetypes=[("CSV/Excel 文件", "*.csv *.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if path:
            self.send_path.set(path)
            self.log(f"已选择 SEND: {os.path.basename(path)}")

    def select_5100_1(self):
        path = filedialog.askopenfilename(
            title="选择 5100 文件①",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if path:
            self.file5100_1_path.set(path)
            self.log(f"已选择 5100①: {os.path.basename(path)}")

    def select_5100_2(self):
        path = filedialog.askopenfilename(
            title="选择 5100 文件②（可选）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if path:
            self.file5100_2_path.set(path)
            self.log(f"已选择 5100②: {os.path.basename(path)}")

    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def transform(self):
        wrt_file = self.wrt_path.get().strip()
        send_file = self.send_path.get().strip()
        f5100_1 = self.file5100_1_path.get().strip()
        f5100_2 = self.file5100_2_path.get().strip()

        if not wrt_file or not send_file:
            messagebox.showerror("错误", "必须选择 WRT 文件 和 SEND 文件！")
            return

        self.transform_btn.config(state="disabled", text="处理中...")
        self.status_var.set("正在处理，请稍候...")
        self.log_text.delete("1.0", tk.END)
        self.log("=== 开始处理 ===")

        try:
            # 1. 读取文件
            self.log("正在读取 WRT 文件...")
            df_wrt = read_any_file(wrt_file)
            if df_wrt is None or df_wrt.empty:
                raise ValueError("WRT 文件读取失败或为空")
            self.log(f"WRT 加载成功，共 {len(df_wrt)} 行，{len(df_wrt.columns)} 列")

            self.log("正在读取 SEND 文件...")
            df_send = read_any_file(send_file)
            if df_send is None or df_send.empty:
                raise ValueError("SEND 文件读取失败或为空")
            self.log(f"SEND 加载成功，共 {len(df_send)} 行，{len(df_send.columns)} 列")

            # 2. 收集 5100 的所有 sheet 名
            self.log("正在扫描 5100 文件的 Sheet 名称（严格验证模式）...")
            all_sheets = get_all_sheet_names([f5100_1, f5100_2])
            self.log(f"共收集到 {len(all_sheets)} 个唯一 Sheet 名（来自两个 5100 文件）")
            if all_sheets:
                sample = list(all_sheets)[:8]
                self.log(f"示例 Sheet: {sample} ...")

            # 3. 核心逻辑：以 WRT 为准重建 SEND
            self.log("【第一阶段】开始以 WRT 为准对齐 SEND 结构（增删 + 更新 A/B/C）...")

            # 构建 WRT key 字典（使用 B 列 index=1 作为唯一标识）
            wrt_dict = {}
            for idx, row in df_wrt.iterrows():
                b_val = str(row[1]).strip() if pd.notna(row[1]) else ""
                if b_val:
                    wrt_dict[b_val] = {
                        'A': row[0] if pd.notna(row[0]) else "",
                        'B': row[1] if pd.notna(row[1]) else "",
                        'C': row[2] if pd.notna(row[2]) else ""
                    }

            # 构建旧 SEND 的 key -> (D, E) 映射
            send_de_map = {}
            for idx, row in df_send.iterrows():
                b_val = str(row[1]).strip() if pd.notna(row[1]) else ""
                if b_val:
                    d_val = str(row[3]).strip() if pd.notna(row[3]) else ""
                    e_val = str(row[4]).strip() if pd.notna(row[4]) else ""
                    if b_val not in send_de_map:  # 取第一个匹配
                        send_de_map[b_val] = (d_val, e_val)

            # 重建新数据
            new_data = []
            carried_count = 0
            new_count = 0
            c_preserved_count = 0   # 新增：记录从旧SEND保留C列的情况

            for b_key in wrt_dict.keys():
                wrt_row = wrt_dict[b_key]
                old_c = ""
                if b_key in send_de_map:
                    d_val, e_val = send_de_map[b_key]
                    # 额外获取旧SEND的C列（如果WRT的C为空则保留它，兼容“表05”等情况）
                    # 这里我们需要重新从df_send查找完整旧行
                    old_row = df_send[df_send[1].astype(str).str.strip() == b_key]
                    if not old_row.empty:
                        old_c = str(old_row.iloc[0, 2]).strip() if pd.notna(old_row.iloc[0, 2]) else ""
                    carried_count += 1
                else:
                    d_val = ""
                    e_val = ""
                    new_count += 1

                # C列处理：优先用WRT的，如果WRT的C为空则保留旧SEND的C（保留“表05”等信息）
                final_c = wrt_row['C']
                if (not final_c or str(final_c).strip() == "") and old_c:
                    final_c = old_c
                    c_preserved_count += 1

                new_data.append([
                    wrt_row['A'],
                    wrt_row['B'],
                    final_c,
                    d_val,
                    e_val
                ])

            df_new = pd.DataFrame(new_data, columns=None)

            self.log(f"重建完成：从 WRT 同步 {len(df_new)} 行")
            self.log(f"  - 保留旧 SEND 的 D/E 列：{carried_count} 行")
            self.log(f"  - 新增项（旧 SEND 无匹配）：{new_count} 行")
            if c_preserved_count > 0:
                self.log(f"  - C列从旧SEND保留（WRT为空时）：{c_preserved_count} 行（例如保留「表05」）")

            # ==================== 第二阶段：用 A 列解析并更新 D/E 列 ====================
            self.log("开始第二阶段：根据 A 列规则更新 D (sheet) 和 E (坐标)...")
            updated_d_count = 0
            updated_e_count = 0

            for idx in df_new.index:
                a_val = df_new.iloc[idx, 0]
                new_d, new_e = parse_a_to_sheet_and_coord(a_val)

                if new_d:
                    old_d = str(df_new.iloc[idx, 3]).strip() if pd.notna(df_new.iloc[idx, 3]) else ""
                    df_new.iloc[idx, 3] = new_d
                    if new_d != old_d:
                        updated_d_count += 1

                if new_e:
                    old_e = str(df_new.iloc[idx, 4]).strip() if pd.notna(df_new.iloc[idx, 4]) else ""
                    df_new.iloc[idx, 4] = new_e
                    if new_e != old_e:
                        updated_e_count += 1

            self.log(f"第二阶段完成：更新了 {updated_d_count} 行 D 列 (sheet)，{updated_e_count} 行 E 列 (坐标)")
            self.log("注意：E 列目前使用逻辑坐标 (RxxCyy)，如需转为真实 Excel 单元格地址（如 AK91），请提供映射规则或例子。")

            # 4. 检查 obsolete（旧 SEND 中但新 WRT 已无）
            old_b_set = set(str(x).strip() for x in df_send[1] if pd.notna(x))
            new_b_set = set(wrt_dict.keys())
            obsolete = old_b_set - new_b_set
            if obsolete:
                self.log(f"警告：以下 {len(obsolete)} 个旧 SEND 项在新 WRT 中已不存在，已自动移除：")
                for ob in list(obsolete)[:10]:
                    self.log(f"    - {ob}")
                if len(obsolete) > 10:
                    self.log(f"    ... 还有 {len(obsolete)-10} 个")

            # 5. 严格 sheet 名验证（防止乱改）
            if all_sheets:
                invalid_rows = []
                for i, row in df_new.iterrows():
                    sheet_name = str(row[3]).strip() if pd.notna(row[3]) else ""
                    if sheet_name and sheet_name not in all_sheets:
                        invalid_rows.append((i+2, sheet_name, str(row[1])[:20] if pd.notna(row[1]) else ""))  # Excel 行号近似

                if invalid_rows:
                    self.log(f"⚠️  检测到 {len(invalid_rows)} 行 的 D 列 sheet 名在所选 5100 文件中不存在：")
                    for r in invalid_rows[:15]:
                        self.log(f"    第 {r[0]} 行 | Sheet: {r[1]} | B列: {r[2]}")
                    if len(invalid_rows) > 15:
                        self.log(f"    ... 还有 {len(invalid_rows)-15} 行")
                    self.log("这些行的坐标 (E 列) 将保持原样，不进行任何自动修改（符合「不允许改动」要求）")
                else:
                    self.log("✅ 所有 SEND 的 D 列 sheet 名均在所选 5100 文件中存在，验证通过")

            # 6. 保存结果
            base_name = os.path.splitext(send_file)[0]
            ext = os.path.splitext(send_file)[1].lower()
            if ext == '.csv':
                output_path = f"{base_name}_updated.csv"
                df_new.to_csv(output_path, index=False, header=False, encoding='utf-8-sig')
            else:
                output_path = f"{base_name}_updated.xlsx"
                df_new.to_excel(output_path, index=False, header=False)

            self.log(f"✅ 更新完成！新文件已保存至：\n    {output_path}")
            self.log("=== 处理结束 ===")

            messagebox.showinfo("处理完成", 
                f"Send 文件已成功更新！\n\n"
                f"输出文件：{os.path.basename(output_path)}\n"
                f"总行数：{len(df_new)}\n"
                f"保留 D/E 列：{carried_count} 行\n"
                f"新增项：{new_count} 行\n\n"
                f"详细日志请查看下方窗口。")

            self.status_var.set(f"完成 | 已生成 {os.path.basename(output_path)}")

        except Exception as e:
            self.log(f"❌ 错误: {str(e)}")
            messagebox.showerror("处理失败", f"发生错误：\n{str(e)}\n\n请检查文件格式或联系开发者。")
            self.status_var.set("错误 | 请查看日志")
        finally:
            self.transform_btn.config(state="normal", text="开始变换 / 更新 SEND")

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateSendApp(root)
    root.mainloop()