import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

def process_v10_mapping(old_csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载新旧 Excel 文件...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        # ==========================================
        # 1. 结构比对 (追踪平移)
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
        # 2. V4 表头侦测：精准扫描新列号 (强化全角兼容)
        # ==========================================
        new_col_to_logic_id = {}
        for c in range(1, max_c):
            for r in range(1, 50):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    # 将可能的全角数字转换为半角，防止匹配失败
                    v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                    m = re.search(r'[\(（](\d+)[\)）]', v_str)
                    if m:
                        new_col_to_logic_id[c] = int(m.group(1))
                        break
                        
        log_widget.insert(tk.END, f"➔ 侦测到 {len(new_col_to_logic_id)} 个显式列头。\n\n")

        # ==========================================
        # 3. 数据更新与“绝对锚点”提取
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在追踪老坐标并提取物理行锚点...\n")
        
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        target_sheet_rows = []
        other_sheet_rows = []
        logic_base_to_phys_row = {} 
        processed_active_ids = set()
        
        updated_count = 0
        deleted_count = 0

        for row in csv_data:
            if len(row) < 3:
                other_sheet_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
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
                        processed_active_ids.add(new_id_val)
                        
                        # 记住这个前缀对应的物理行锚点
                        base_id = id_val[:-2] 
                        logic_base_to_phys_row[base_id] = new_row_idx
                        
                        if new_id_val != id_val or old_cell != new_cell:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [追踪更新] {id_val} ➔ {new_id_val} | {old_cell} ➔ {new_cell}\n")
                    else:
                        deleted_count += 1
                        log_widget.insert(tk.END, f" ❌ [物理废弃] {id_val} 的原坐标已被删除。\n")
            else:
                other_sheet_rows.append(row)

        # ==========================================
        # 4. 完美追加：基于锚点和新表头繁衍数据
        # ==========================================
        log_widget.insert(tk.END, f"\n【4/4】正在侦测新增列并执行全局排版...\n")
        added_count = 0
        
        for base_id, phys_row in logic_base_to_phys_row.items():
            for phys_col, logic_col_num in new_col_to_logic_id.items():
                expected_id = f"{base_id}{logic_col_num:02d}"
                
                if expected_id not in processed_active_ids:
                    expected_cell = f"{get_column_letter(phys_col)}{phys_row}"
                    target_sheet_rows.append([expected_id, str(target_sheet), expected_cell])
                    processed_active_ids.add(expected_id)
                    added_count += 1
                    log_widget.insert(tk.END, f" ➕ [全新列追加] 生成: {expected_id} ➔ {expected_cell}\n")

        # 跨页残留清理
        final_other_rows = []
        for row in other_sheet_rows:
            if len(row) >= 3 and row[0].strip() in processed_active_ids:
                pass # 丢弃老数据
            else:
                final_other_rows.append(row)

        # ==========================================
        # 5. 全局排序 (绝杀：让 30 完美插在 29 下方)
        # ==========================================
        final_rows = target_sheet_rows + final_other_rows
        # 过滤空行并严格按照 A 列（ID）从小到大全局排序
        final_rows = [row for row in final_rows if row and any(row)]
        final_rows.sort(key=lambda x: str(x[0]).strip())

        # ==========================================
        # 导出结果
        # ==========================================
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(final_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n更新: {updated_count} 个\n新列生成并排版: {added_count} 个\n删除: {deleted_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 排序与追加完成！\n\n新列生成: {added_count}个 (已自动插入对应位置)")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标引擎 (V10 全局排序追加版)")
        self.root.geometry("750x600")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        self.csv_entry = create_file_picker(root, "1. 请选择映射 CSV 文件 (例: DafKazeiWrt.csv):", [("CSV Files", "*.csv")])
        self.old_xl_entry = create_file_picker(root, "2. 请选择【原本】未修改的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        
        tk.Label(root, text="4. 目标 Sheet 名称 (例如: 47):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        self.btn_start = tk.Button(root, text="🚀 启动追踪与精准无缝追加", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
            messagebox.showwarning("提示", "请完整填写！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_v10_mapping(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
