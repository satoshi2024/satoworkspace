import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

def process_v12_mapping(old_csv_path, old_excel_path, new_excel_path, target_sheet, added_col, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载 Excel 文件...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        # ==========================================
        # 1. 结构比对 (V4引擎：计算物理偏移量)
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在比对物理结构与侦测最新表头...\n")
        log_widget.update()

        def get_col_sig(ws, col_idx):
            vals = []
            for r in range(1, 50):
                val = ws.cell(row=r, column=col_idx).value
                v_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                v_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', v_str)
                vals.append(v_str)
            return "|".join(vals)

        old_col_sigs = [get_col_sig(ws_old, c) for c in range(1, max_c)]
        new_col_sigs = [get_col_sig(ws_new, c) for c in range(1, max_c)]

        sm_col = difflib.SequenceMatcher(None, old_col_sigs, new_col_sigs)
        col_map = {}
        for tag, i1, i2, j1, j2 in sm_col.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    col_map[old_i + 1] = new_j + 1

        def get_row_sig(ws, row_idx):
            vals = []
            for c in range(1, 50):
                val = ws.cell(row=row_idx, column=c).value
                v_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                v_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', v_str)
                vals.append(v_str)
            return "|".join(vals)

        old_row_sigs = [get_row_sig(ws_old, r) for r in range(1, max_r)]
        new_row_sigs = [get_row_sig(ws_new, r) for r in range(1, max_r)]
        sm_row = difflib.SequenceMatcher(None, old_row_sigs, new_row_sigs)
        row_map = {}
        for tag, i1, i2, j1, j2 in sm_row.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    row_map[old_i + 1] = new_j + 1

        # ==========================================
        # 2. 精准表头侦测 (兼容全角数字)
        # ==========================================
        new_col_to_logic_id = {}
        target_phys_col = None 
        
        for c in range(1, max_c):
            for r in range(1, 50):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                    m = re.search(r'[\(（](\d+)[\)）]', v_str)
                    if m:
                        logic_num = int(m.group(1))
                        new_col_to_logic_id[c] = logic_num
                        if added_col and logic_num == int(added_col):
                            target_phys_col = c
                        break

        # ==========================================
        # 3. 提取 CSV 历史数据的物理锚点 (彻底抛弃 Excel 读取行号)
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在严格校验 CSV 并提取物理锚点...\n")
        
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        target_sheet_row_prefixes = set() 
        logic_base_to_phys_row = {} # 核心：记忆 5501 -> 第14行
        
        for row in csv_data:
            if len(row) >= 3 and row[1].strip() == str(target_sheet):
                id_val, old_cell = row[0].strip(), row[2].strip()
                if id_val.isdigit() and len(id_val) >= 4:
                    prefix = id_val[:-2]
                    target_sheet_row_prefixes.add(prefix)
                    
                    # 利用老坐标结合平移量，算出它现在在第几行
                    match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                    if match:
                        old_row_idx = int(match.group(2))
                        new_row_idx = row_map.get(old_row_idx)
                        if new_row_idx:
                            logic_base_to_phys_row[prefix] = new_row_idx
                            
        if not target_sheet_row_prefixes:
            messagebox.showerror("错误", f"在 CSV 中完全找不到 B 列为 {target_sheet} 的数据！")
            return
            
        log_widget.insert(tk.END, f"➔ 成功锁定 {len(logic_base_to_phys_row)} 个行锚点。\n\n")

        # ==========================================
        # 4. 执行更新、跨页拉取 与 精确追加
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】开始安全更新与排版追加...\n")
        
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        target_sheet_rows = []
        other_sheet_rows = []
        
        updated_count = 0
        deleted_count = 0
        migrated_count = 0
        
        has_target_col = False
        added_col_padded = added_col.zfill(2) if added_col else ""

        for row in csv_data:
            if len(row) < 3:
                other_sheet_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            # --- 场景 A：跨页拉取 ---
            if added_col and id_val.endswith(added_col_padded) and id_val[:-2] in target_sheet_row_prefixes:
                if sheet_val != str(target_sheet):
                    phys_row = logic_base_to_phys_row.get(id_val[:-2])
                    
                    if target_phys_col and phys_row:
                        new_cell = f"{get_column_letter(target_phys_col)}{phys_row}"
                        target_sheet_rows.append([id_val, str(target_sheet), new_cell])
                        has_target_col = True 
                        migrated_count += 1
                        log_widget.insert(tk.END, f" 🚀 [跨页拉取] ID:{id_val} | Sheet {sheet_val} ➔ {target_sheet} | 坐标 ➔ {new_cell}\n")
                        continue

            # --- 场景 B：正常当前 Sheet 更新 ---
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                if added_col and id_val.endswith(added_col_padded):
                    has_target_col = True 
                    
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                if match:
                    col_str, row_str = match.groups()
                    old_col_idx = column_index_from_string(col_str)
                    old_row_idx = int(row_str)
                    
                    new_col_idx = col_map.get(old_col_idx)
                    new_row_idx = row_map.get(old_row_idx)
                    
                    if new_col_idx and new_row_idx:
                        new_cell = f"{get_column_letter(new_col_idx)}{new_row_idx}"
                        detected_logic_num = new_col_to_logic_id.get(new_col_idx)
                        
                        if detected_logic_num is not None:
                            new_id_val = f"{id_val[:-2]}{detected_logic_num:02d}"
                        else:
                            new_id_val = id_val
                            
                        target_sheet_rows.append([new_id_val, sheet_val, new_cell])
                        if new_id_val != id_val or old_cell != new_cell:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [追踪更新] {id_val} ➔ {new_id_val} | {old_cell} ➔ {new_cell}\n")
                    else:
                        deleted_count += 1
                        log_widget.insert(tk.END, f" ❌ [废弃删除] {id_val} 原坐标已物理删除。\n")
                else:
                    target_sheet_rows.append(row)
            else:
                other_sheet_rows.append(row)

        # --- 场景 C：严格校验后的精确追加 ---
        added_count = 0
        if added_col and not has_target_col and target_phys_col:
            log_widget.insert(tk.END, f"\n 💡 校验完毕：CSV 中未发现 {added_col} 列数据，开始精确追加...\n")
            
            for row_prefix in sorted(target_sheet_row_prefixes):
                phys_row = logic_base_to_phys_row.get(row_prefix)
                
                if phys_row:
                    new_id = f"{row_prefix}{added_col_padded}"
                    new_cell = f"{get_column_letter(target_phys_col)}{phys_row}"
                    target_sheet_rows.append([new_id, str(target_sheet), new_cell])
                    added_count += 1
                    log_widget.insert(tk.END, f" ➕ [全新列追加] 生成: {new_id} ➔ {new_cell}\n")

        # ==========================================
        # 5. 全局排版合并
        # ==========================================
        final_rows = target_sheet_rows + other_sheet_rows
        final_rows = [row for row in final_rows if row and any(row)]
        final_rows.sort(key=lambda x: str(x[0]).strip())

        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(final_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n常规更新: {updated_count} 个\n跨页拉取: {migrated_count} 个\n校验后追加: {added_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 严格校验处理完成！\n\n更新: {updated_count}\n拉取: {migrated_count}\n追加: {added_count}\n\n已自动全局排序。")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射引擎 (V12 锚点修复版)")
        self.root.geometry("750x680")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        self.csv_entry = create_file_picker(root, "1. 请选择【原本】的映射 CSV 文件:", [("CSV Files", "*.csv")])
        self.old_xl_entry = create_file_picker(root, "2. 请选择【原本】未修改的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        
        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=15, pady=8)
        
        tk.Label(input_frame, text="4. 目标 Sheet (如 47):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        tk.Label(input_frame, text="5. 需校验/追加的列号 (选填,如 30):", font=("MS Gothic", 9, "bold")).pack(side="left", padx=(20, 0))
        self.col_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="red")
        self.col_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 校验 CSV AB列并精确处理", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
        add_col = self.col_entry.get().strip()
        
        if not all([csv_p, old_xl, new_xl, sheet_n]):
            messagebox.showwarning("提示", "请填写前 4 项必填内容！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_v12_mapping(csv_p, old_xl, new_xl, sheet_n, add_col, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
