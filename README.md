import os
import re
import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import get_column_letter

def process_upsert_mapping(old_csv_path, new_excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】启动 V6 追加更新引擎，读取文件...\n")
        log_widget.update()
        
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        if str(target_sheet) not in wb_new.sheetnames:
            messagebox.showerror("错误", f"最新的 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_new = wb_new[str(target_sheet)]

        # ==========================================
        # 1. 提取当前 Sheet 的逻辑 ID 前缀
        # ==========================================
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        prefixes = {}
        for row in csv_data:
            if len(row) >= 3 and row[1].strip() == str(target_sheet):
                id_val = row[0].strip()
                if id_val.isdigit() and len(id_val) >= 4:
                    p = id_val[:-4] 
                    prefixes[p] = prefixes.get(p, 0) + 1
                    
        if not prefixes:
            messagebox.showerror("错误", f"在 CSV 中找不到 Sheet [{target_sheet}] 的历史记录，无法自动推断逻辑前缀！")
            return
            
        target_prefix = max(prefixes, key=prefixes.get)
        log_widget.insert(tk.END, f"➔ 成功提取当前 Sheet 逻辑前缀: [{target_prefix}]\n")

        # ==========================================
        # 2. 扫描新 Excel，提取绝对物理坐标
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在解析最新 Excel 物理布局...\n")
        
        row_map = {} 
        id_col_idx = None
        
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

        col_map = {} 
        for r in range(1, 40):
            for c in range(1, ws_new.max_column + 1):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    v_str = str(val).strip()
                    m = re.search(r'^[\(（]\s*(\d+)\s*[\)）]$', v_str)
                    if m:
                        col_map[int(m.group(1))] = get_column_letter(c)

        log_widget.insert(tk.END, f"➔ 侦测到 {len(row_map)} 个数据行，{len(col_map)} 个逻辑列。\n\n")

        # ==========================================
        # 3. 生成当前画面的理想映射字典
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在重构当前画面的坐标期望值...\n")
        
        expected_mappings = {}
        for r_logic, r_idx in row_map.items():
            for c_logic, c_str in col_map.items():
                new_id = f"{target_prefix}{r_logic}{c_logic:02d}"
                new_cell = f"{c_str}{r_idx}"
                expected_mappings[new_id] = [new_id, str(target_sheet), new_cell]

        # ==========================================
        # 4. 只更新和追加，绝不删除！
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在将新坐标融合至 CSV (安全保留原有数据)...\n")
        
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        final_rows = []
        
        updated_count = 0
        cross_sheet_count = 0
        processed_ids = set()

        # 第一步：遍历原 CSV，做无损更新
        for row in csv_data:
            if len(row) < 3:
                final_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()

            if id_val in expected_mappings:
                # 这个 ID 存在于最新画面中，强制更新为新坐标和新 Sheet！
                new_mapping = expected_mappings[id_val]
                final_rows.append(new_mapping)
                processed_ids.add(id_val)
                
                if sheet_val != str(target_sheet):
                    cross_sheet_count += 1
                    log_widget.insert(tk.END, f" 🚀 [跨页迁移] {id_val} 已从 Sheet {sheet_val} 迁移至 {target_sheet}！\n")
                elif old_cell != new_mapping[2]:
                    updated_count += 1
            else:
                # 【核心修复】：画面上没有它，或者它属于其他表，原封不动保留！绝不删除！
                final_rows.append(row)

        # 第二步：追加纯新增的 ID (原本 CSV 压根没有的列)
        added_count = 0
        for id_val, new_mapping in expected_mappings.items():
            if id_val not in processed_ids:
                final_rows.append(new_mapping)
                added_count += 1
                log_widget.insert(tk.END, f" ➕ [全新追加] CSV 原本无此列记录，已追加: {id_val} ➔ {new_mapping[2]}\n")

        # ==========================================
        # 5. 导出结果并备份
        # ==========================================
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(final_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n同页坐标更新: {updated_count} 条\n跨页接管迁移: {cross_sheet_count} 条\n全新追加生成: {added_count} 条\n(所有原数据已绝对安全保留)\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 追加更新完成！\n\n更新: {updated_count}\n迁移: {cross_sheet_count}\n追加: {added_count}\n\n已安全保留所有非相关数据。")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")


class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射智能追加工具 (无损 Upsert 版) v6.0")
        self.root.geometry("750x550")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        # 界面精简，只留 2 个输入框！
        self.csv_entry = create_file_picker(root, "1. 请选择映射 CSV 文件 (例: DafKazeiWrt.csv):", [("CSV Files", "*.csv")])
        self.new_xl_entry = create_file_picker(root, "2. 请选择【改修后】最新的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        
        tk.Label(root, text="3. 请输入本次要处理的 Sheet 名称 (例如: 47):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        self.btn_start = tk.Button(root, text="🚀 开始无损更新与追加映射", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        tk.Label(root, text="执行日志 (只追加、更新，绝不删除):", font=("MS Gothic", 9)).pack(anchor="w", padx=15)
        self.log_text = scrolledtext.ScrolledText(root, height=12, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
    def browse_file(self, entry_widget, file_types):
        path = filedialog.askopenfilename(filetypes=file_types)
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            
    def start_process(self):
        csv_p = self.csv_entry.get().strip()
        new_xl = self.new_xl_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        
        if not all([csv_p, new_xl, sheet_n]):
            messagebox.showwarning("提示", "请完整选择 2 个文件并填写 Sheet 名称！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_upsert_mapping(csv_p, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
