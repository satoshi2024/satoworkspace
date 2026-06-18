import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

def process_ultimate_mapping(old_csv_path, old_excel_path, new_excel_path, target_sheet, added_col, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载 Excel 文件，请稍候...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在某个 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        # ==========================================
        # 1. 常规 Diff：对比新旧 Excel 的物理结构差异 (追踪平移)
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在对比物理结构并启动 V4 表头侦测...\n")
        log_widget.update()

        def get_col_sig(ws, col_idx):
            vals = []
            for r in range(1, 51):
                val = ws.cell(row=r, column=col_idx).value
                val_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                # 抹除数字干扰，保证只比对汉字结构
                val_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', val_str)
                vals.append(val_str)
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
            for c in range(1, 51):
                val = ws.cell(row=row_idx, column=c).value
                val_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                val_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', val_str)
                vals.append(val_str)
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
        # 2. V4 核心：扫描新表头逻辑列号及行番号字典
        # ==========================================
        new_col_to_logic_id = {}
        logic_id_to_new_col = {} # 用于跨页抓取
        for c in range(1, max_c):
            for r in range(1, 30):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    m = re.search(r'^[\(（]\s*(\d+)\s*[\)）]$', str(val).strip())
                    if m:
                        logic_num = int(m.group(1))
                        new_col_to_logic_id[c] = logic_num
                        logic_id_to_new_col[logic_num] = c
                        break

        logic_row_to_phys_row = {}
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
                        logic_row_to_phys_row[v_str.zfill(2)] = r

        log_widget.insert(tk.END, f"➔ 侦测到 {len(new_col_to_logic_id)} 个显式表头！\n\n")

        # ==========================================
        # 3. 读取 CSV，融合 V4 改名与定向跨页抓取
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在更新映射 (结合表头纠正与跨页迁移)...\n")
        
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        # 提取业务前缀 (例如 '55')
        prefixes = {}
        for row in csv_data:
            if len(row) >= 3 and row[1].strip() == str(target_sheet):
                id_val = row[0].strip()
                if id_val.isdigit() and len(id_val) >= 4:
                    prefixes[id_val[:-4]] = prefixes.get(id_val[:-4], 0) + 1
        target_prefix = max(prefixes, key=prefixes.get) if prefixes else None

        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        new_rows = []
        updated_count = 0
        deleted_count = 0
        migrated_count = 0
        
        for row in csv_data:
            if len(row) < 3:
                new_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            # --- 场景 A: 定向跨页抓取 (你输入的追加列) ---
            if added_col and target_prefix and id_val.startswith(target_prefix) and id_val.endswith(added_col.zfill(2)):
                target_c_idx = logic_id_to_new_col.get(int(added_col))
                target_r_phys = logic_row_to_phys_row.get(id_val[-4:-2])
                
                if target_c_idx and target_r_phys:
                    new_cell = f"{get_column_letter(target_c_idx)}{target_r_phys}"
                    new_rows.append([id_val, str(target_sheet), new_cell])
                    
                    if sheet_val != str(target_sheet):
                        migrated_count += 1
                        log_widget.insert(tk.END, f" 🚀 [跨页迁移] ID:{id_val} | Sheet {sheet_val} ➔ {target_sheet} | 坐标 ➔ {new_cell}\n")
                    else:
                        updated_count += 1
                        log_widget.insert(tk.END, f" 🔄 [同页更新] ID:{id_val} | 坐标 ➔ {new_cell}\n")
                    continue

            # --- 场景 B: 正常页内追踪与 V4 改名 ---
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                if match:
                    col_str, row_str = match.groups()
                    old_col_idx = column_index_from_string(col_str)
                    old_row_idx = int(row_str)
                    
                    # 1. 使用 Diff 追踪它物理漂移到了哪里
                    new_col_idx = col_map.get(old_col_idx)
                    new_row_idx = row_map.get(old_row_idx)
                    
                    if new_col_idx and new_row_idx:
                        new_cell = f"{get_column_letter(new_col_idx)}{new_row_idx}"
                        
                        # 2. V4 绝杀：抬头看这列现在叫什么名字？
                        detected_logic_num = new_col_to_logic_id.get(new_col_idx)
                        
                        if detected_logic_num is not None:
                            new_id_val = f"{id_val[:-2]}{detected_logic_num:02d}"
                        else:
                            new_id_val = id_val
                            
                        new_rows.append([new_id_val, sheet_val, new_cell])
                        
                        if new_id_val != id_val or old_cell != new_cell:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🎯 [V4精准追踪] ID: {id_val} ➔ {new_id_val} | 坐标: {old_cell} ➔ {new_cell}\n")
                    else:
                        deleted_count += 1
                        log_widget.insert(tk.END, f" ❌ [结构删除] ID:{id_val} | 原坐标 {old_cell} 已被彻底删除。\n")
                else:
                    new_rows.append(row)
            else:
                # 不受影响的数据
                new_rows.append(row)
                
        # ==========================================
        # 4. 导出结果
        # ==========================================
        log_widget.insert(tk.END, f"\n【4/4】正在保存结果...\n")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\nV4 表头改名与追踪: {updated_count} 个\n跨页定向迁移: {migrated_count} 个\n失效删除: {deleted_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 融合修正完成！\n\n表头追踪修正: {updated_count}\n跨页强制迁移: {migrated_count}\n废弃删除: {deleted_count}")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射智能引擎 (V4表头侦测 + 跨页抓取版)")
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

        self.csv_entry = create_file_picker(root, "1. 请选择【原本】的映射 CSV 文件 (例: DafKazeiWrt.csv):", [("CSV Files", "*.csv")])
        self.old_xl_entry = create_file_picker(root, "2. 请选择【原本】未修改的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        
        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=15, pady=8)
        
        tk.Label(input_frame, text="4. 目标 Sheet (如 47):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        tk.Label(input_frame, text="5. 需定向抓取的跨页列号(选填, 如 30):", font=("MS Gothic", 9, "bold")).pack(side="left", padx=(20, 0))
        self.col_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="red")
        self.col_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 启动表头侦测与跨页抓取", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
            messagebox.showwarning("提示", "请完整选择 3 个文件并填写 Sheet 名称！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_ultimate_mapping(csv_p, old_xl, new_xl, sheet_n, add_col, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
