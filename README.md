import os
import re
import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import get_column_letter

def process_v13_strict_mapping(csv_path, new_excel_path, target_sheet, added_col, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载最新 Excel 文件...\n")
        log_widget.update()
        
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        if target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_new = wb_new[target_sheet]
        max_c = ws_new.max_column + 10

        # ==========================================
        # 1. 终极解析 Excel 坐标 (解决行番号拆分问题)
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在解析最新表头与行号锚点...\n")
        log_widget.update()

        # 解析列号 (18), (19), (30)...
        logic_col_to_phys_col = {}
        for c in range(1, max_c):
            for r in range(1, 50):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                    m = re.search(r'[\(（](\d+)[\)）]', v_str)
                    if m:
                        logic_col_to_phys_col[int(m.group(1))] = get_column_letter(c)
                        break

        # 解析行号 (解决 0 和 1 分在两列的问题)
        logic_row_to_phys_row = {}
        id_col_idx = None
        id_row_idx = None
        
        # 找 "行番号" 表头在哪
        for r in range(1, 40):
            for c in range(1, 20):
                val = ws_new.cell(row=r, column=c).value
                if val and "行番号" in str(val).replace(" ", ""):
                    id_col_idx = c
                    id_row_idx = r
                    break
            if id_col_idx: break
            
        if id_col_idx:
            # 往下扫，把相邻两列的数字拼起来 (0 + 1 = 01)
            for r in range(id_row_idx + 1, ws_new.max_row + 1):
                v1 = ws_new.cell(row=r, column=id_col_idx).value
                v2 = ws_new.cell(row=r, column=id_col_idx + 1).value
                
                s1 = str(v1).strip() if v1 is not None else ""
                s2 = str(v2).strip() if v2 is not None else ""
                
                # 如果是分在两列 (如 s1='0', s2='1')
                if s1.isdigit() and s2.isdigit():
                    logic_row = s1 + s2
                    logic_row_to_phys_row[logic_row] = r
                # 如果是合并在一起的 (如 s1='01')
                elif s1.isdigit() and len(s1) >= 2:
                    logic_row_to_phys_row[s1.zfill(2)] = r

        log_widget.insert(tk.END, f"➔ 成功解析 {len(logic_col_to_phys_col)} 个列坐标，{len(logic_row_to_phys_row)} 个行坐标。\n\n")

        # ==========================================
        # 2. 第一次通读 CSV：寻找是否有跨页数据
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在严格比对 CSV 的 A 列与 B 列...\n")
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        target_prefix = None
        # 寻找目标 Sheet 的逻辑前缀 (例如 55)
        for row in csv_data:
            if len(row) >= 3 and row[1].strip() == str(target_sheet):
                id_val = row[0].strip()
                if id_val.isdigit() and len(id_val) >= 4:
                    target_prefix = id_val[:-4]
                    break
                    
        if not target_prefix:
            messagebox.showerror("错误", f"在 CSV 中找不到 B 列为 {target_sheet} 的数据！")
            return

        added_col_padded = added_col.zfill(2) if added_col else ""
        existing_added_ids = set() # 记录目标列 (如 30) 在全局 CSV 中是否已经存在
        
        if added_col:
            for row in csv_data:
                if len(row) >= 1:
                    id_val = row[0].strip()
                    if id_val.startswith(target_prefix) and id_val.endswith(added_col_padded):
                        existing_added_ids.add(id_val)

        # ==========================================
        # 3. 完美原序重建与精确追加 (绝对不使用 Sort)
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在无损更新坐标并精准插入...\n")
        
        new_rows = []
        updated_count = 0
        migrated_count = 0
        injected_count = 0
        injected_tracker = set()

        for row in csv_data:
            if len(row) < 3:
                new_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()

            # --- 场景 A：该行就是我们要找的跨页列 (存在于库中，但可能在 48 页) ---
            if added_col and id_val in existing_added_ids:
                logic_row_str = id_val[-4:-2]
                phys_row = logic_row_to_phys_row.get(logic_row_str)
                phys_col = logic_col_to_phys_col.get(int(added_col))
                
                if phys_row and phys_col:
                    new_cell = f"{phys_col}{phys_row}"
                    new_rows.append([id_val, str(target_sheet), new_cell]) # 强制改回 47 页
                    
                    if sheet_val != str(target_sheet):
                        migrated_count += 1
                        log_widget.insert(tk.END, f" 🚀 [跨页拉回] ID:{id_val} | Sheet {sheet_val} ➔ {target_sheet}\n")
                    elif old_cell != new_cell:
                        updated_count += 1
                else:
                    new_rows.append(row)
                continue

            # --- 场景 B：正常处理当前 Sheet 的数据 ---
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                logic_row_str = id_val[-4:-2]
                logic_col_num = int(id_val[-2:])
                
                phys_row = logic_row_to_phys_row.get(logic_row_str)
                phys_col = logic_col_to_phys_col.get(logic_col_num)
                
                # 正常更新坐标
                if phys_row and phys_col:
                    new_cell = f"{phys_col}{phys_row}"
                    new_rows.append([id_val, sheet_val, new_cell])
                    if old_cell != new_cell:
                        updated_count += 1
                else:
                    new_rows.append(row)
                
                # --- 【绝杀：跟随追加】---
                # 如果当前读到的是 29 列，而用户输入了 30 列，且 30 列在库中不存在
                if added_col and logic_col_num == int(added_col) - 1:
                    needed_id = f"{target_prefix}{logic_row_str}{added_col_padded}"
                    
                    if needed_id not in existing_added_ids and needed_id not in injected_tracker:
                        add_phys_col = logic_col_to_phys_col.get(int(added_col))
                        if phys_row and add_phys_col:
                            add_cell = f"{add_phys_col}{phys_row}"
                            # 直接紧跟在 29 下方追加 30 的数据！
                            new_rows.append([needed_id, str(target_sheet), add_cell])
                            injected_tracker.add(needed_id)
                            injected_count += 1
                            log_widget.insert(tk.END, f" ➕ [原位追加] 成功在 {id_val} 下方追加: {needed_id} ➔ {add_cell}\n")
            
            # --- 场景 C：其他毫无关系的数据 ---
            else:
                new_rows.append(row)

        # ==========================================
        # 4. 导出结果 (覆盖原文件)
        # ==========================================
        output_csv_path = csv_path.replace(".csv", "_Updated.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        shutil.copyfile(csv_path, csv_path + ".bak")
        shutil.copyfile(output_csv_path, csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n同页坐标更新: {updated_count} 个\n跨页接管拉回: {migrated_count} 个\n原位紧跟追加: {injected_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 严格追加完成！\n\n坐标更新: {updated_count}\n跨页拉回: {migrated_count}\n原位追加: {injected_count}\n\n(已保留原表100%排序，请查看 WinMerge)")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")


class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射引擎 (V13 无损破壁原序版)")
        self.root.geometry("750x620")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        # 界面精简：只留 2 个文件选择！
        self.csv_entry = create_file_picker(root, "1. 请选择映射 CSV 文件 (例: DafKazeiWrt.csv):", [("CSV Files", "*.csv")])
        self.new_xl_entry = create_file_picker(root, "2. 请选择【改修后】最新的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        
        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=15, pady=8)
        
        tk.Label(input_frame, text="3. 目标 Sheet (如 47):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        tk.Label(input_frame, text="4. 需校验/追加的列号 (选填,如 30):", font=("MS Gothic", 9, "bold")).pack(side="left", padx=(20, 0))
        self.col_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="red")
        self.col_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 开始严格校验与原位追加", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        tk.Label(root, text="执行日志 (绝对保证原表排序):", font=("MS Gothic", 9)).pack(anchor="w", padx=15)
        self.log_text = scrolledtext.ScrolledText(root, height=14, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
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
        add_col = self.col_entry.get().strip()
        
        if not all([csv_p, new_xl, sheet_n]):
            messagebox.showwarning("提示", "请填写前 3 项必填内容！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_v13_strict_mapping(csv_p, new_xl, sheet_n, add_col, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
