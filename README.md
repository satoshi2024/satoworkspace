import os
import re
import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import get_column_letter

def process_reconstruction_mapping(old_csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】启动 V5 状态重构引擎，读取文件...\n")
        log_widget.update()
        
        # V5 引擎不需要对比旧 Excel，直接读取新 Excel 作为绝对真理！
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        if str(target_sheet) not in wb_new.sheetnames:
            messagebox.showerror("错误", f"最新的 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_new = wb_new[str(target_sheet)]

        # ==========================================
        # 1. 从 CSV 中提取当前 Sheet 的逻辑 ID 前缀 (例如 55)
        # ==========================================
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        prefixes = {}
        for row in csv_data:
            if len(row) >= 3 and row[1].strip() == str(target_sheet):
                id_val = row[0].strip()
                if id_val.isdigit() and len(id_val) >= 4:
                    # 剥离最后4位(行号2位+列号2位)，剩下的是前缀
                    p = id_val[:-4] 
                    prefixes[p] = prefixes.get(p, 0) + 1
                    
        if not prefixes:
            messagebox.showerror("错误", f"在 CSV 中找不到 Sheet [{target_sheet}] 的历史记录，无法推断 ID 前缀。")
            return
            
        # 获取最常见的前缀 (例如 '55')
        target_prefix = max(prefixes, key=prefixes.get)
        log_widget.insert(tk.END, f"➔ 成功提取当前 Sheet 的逻辑前缀: [{target_prefix}]\n")

        # ==========================================
        # 2. 扫描新 Excel，提取绝对真理：行坐标与列坐标
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在解析最新 Excel 的物理布局...\n")
        
        row_map = {} # 逻辑行号 '01' -> 物理行号 14
        id_col_idx = None
        
        # 找 "行番号" 列
        for r in range(1, 40):
            for c in range(1, 20):
                val = ws_new.cell(row=r, column=c).value
                if val and "行番号" in str(val).replace(" ", ""):
                    id_col_idx = c
                    break
            if id_col_idx: break
            
        if id_col_idx:
            for r in range(1, ws_new.max_row + 1):
                val = ws_new.cell(row=r, column=id_col_idx).value
                if val is not None:
                    v_str = str(val).strip()
                    if v_str.isdigit() and 1 <= int(v_str) <= 99:
                        row_map[v_str.zfill(2)] = r

        col_map = {} # 逻辑列号 30 -> 物理列字母 'DM'
        for r in range(1, 40):
            for c in range(1, ws_new.max_column + 1):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    v_str = str(val).strip()
                    m = re.search(r'^[\(（]\s*(\d+)\s*[\)）]$', v_str)
                    if m:
                        col_map[int(m.group(1))] = get_column_letter(c)

        log_widget.insert(tk.END, f"➔ 侦测到 {len(row_map)} 个数据行，{len(col_map)} 个括号列表头。\n\n")

        # ==========================================
        # 3. 交叉相乘，生成完美的新 CSV 映射块
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在根据 Excel 画面重构理想坐标字典...\n")
        
        new_mappings = {}
        for r_logic, r_idx in row_map.items():
            for c_logic, c_str in col_map.items():
                new_id = f"{target_prefix}{r_logic}{c_logic:02d}"
                new_cell = f"{c_str}{r_idx}"
                new_mappings[new_id] = [new_id, str(target_sheet), new_cell]

        # ==========================================
        # 4. 融合写入 CSV (解决跨页漂移和新增列)
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在将重构数据覆盖至 CSV (自动处理跨 Sheet)...\n")
        
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        final_rows = []
        inserted = False
        
        updated_count = 0
        deleted_count = 0
        added_count = 0

        for row in csv_data:
            if len(row) < 3:
                final_rows.append(row)
                continue
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()

            # 情况A：这个 ID 在我们刚刚生成的完美字典里！
            if id_val in new_mappings:
                if not inserted:
                    # 在遇到第一个属于新字典的 ID 时，把整个完美字典一波性按顺序插入！
                    sorted_new = sorted(new_mappings.values(), key=lambda x: x[0])
                    final_rows.extend(sorted_new)
                    inserted = True
                
                # 统计差异 (纯为了打印日志)
                new_row = new_mappings[id_val]
                if sheet_val != str(target_sheet):
                    log_widget.insert(tk.END, f" 🚀 [跨页迁移] {id_val} 从 Sheet {sheet_val} 搬家到了当前 {target_sheet}！\n")
                    updated_count += 1
                elif old_cell != new_row[2]:
                    updated_count += 1
                
                # 扔掉旧的那行，因为完美的已经插入了
                continue

            # 情况B：这个 ID 属于当前 Sheet，但在新字典里没有（物理删除了）
            if sheet_val == str(target_sheet) and id_val.startswith(target_prefix):
                log_widget.insert(tk.END, f" ❌ [彻底删除] 画面上找不到 {id_val} ({old_cell})，已清理。\n")
                deleted_count += 1
                continue

            # 情况C：其他毫无关系的数据（保持原样）
            final_rows.append(row)

        # 极端保护机制：如果原表是空的，强制加在最后
        if not inserted and new_mappings:
            sorted_new = sorted(new_mappings.values(), key=lambda x: x[0])
            final_rows.extend(sorted_new)
            added_count = len(new_mappings)
            log_widget.insert(tk.END, f" ➕ [全新生成] CSV 原本无此页记录，已全新生成 {added_count} 条数据。\n")

        # 保存覆盖
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(final_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n共生成绝对真理映射 {len(new_mappings)} 条。\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 彻底重构完成！\n请检查 CSV！")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")


# ==========================================
# GUI 保持不变，让你顺手
# ==========================================
class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射重构工具 (绝对真理重绘版) v5.0")
        self.root.geometry("750x650")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        self.csv_entry = create_file_picker(root, "1. 请选择【原本】的映射 CSV 文件 (例: DafKazeiWrt.csv):", [("CSV Files", "*.csv")])
        self.old_xl_entry = create_file_picker(root, "2. [已无需此项, 保持原样即可] 请选择未修改的 Excel:", [("Excel Files", "*.xlsm *.xlsx")])
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        
        tk.Label(root, text="4. 请输入本次要更新的 Sheet 名称 (例如: 47):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        self.btn_start = tk.Button(root, text="🚀 基于新 Excel 画面彻底重构本页映射", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        tk.Label(root, text="执行日志:", font=("MS Gothic", 9)).pack(anchor="w", padx=15)
        self.log_text = scrolledtext.ScrolledText(root, height=12, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
    def browse_file(self, entry_widget, file_types):
        path = filedialog.askopenfilename(filetypes=file_types)
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            
    def start_process(self):
        csv_p = self.csv_entry.get().strip()
        old_xl = self.old_xl_entry.get().strip()
        new_xl = self.new_xl_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        
        if not all([csv_p, old_xl, new_xl, sheet_n]):
            messagebox.showwarning("提示", "请完整选择 3 个文件并填写 Sheet 名称！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_reconstruction_mapping(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
