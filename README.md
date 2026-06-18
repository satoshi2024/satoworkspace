import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

def process_dynamic_mapping(old_csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
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

        # ==========================================
        # 1. 免疫表头数字的 Diff 算法，精准追踪物理坐标
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在分析 Excel 结构差异 (追踪插入/删除)...\n")
        log_widget.update()

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        def get_col_sig(ws, col_idx):
            vals = []
            for r in range(1, 40):
                val = ws.cell(row=r, column=col_idx).value
                if val is not None:
                    v_str = str(val).strip()
                    # 抹除 (18), （19） 等数字，保证结构匹配不受列名数字变化影响
                    v_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', v_str)
                    v_str = v_str.replace(" ", "").replace("　", "")
                    if v_str:
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
            for c in range(1, 30):
                val = ws.cell(row=row_idx, column=c).value
                if val is not None:
                    v_str = str(val).strip()
                    v_str = re.sub(r'[\(（]\s*\d+\s*[\)）]', '', v_str)
                    v_str = v_str.replace(" ", "").replace("　", "")
                    if v_str:
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
        # 2. 从新 Excel 显式提取最新列名 (28, 29, 30...)
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在侦测新版 Excel 表头中的逻辑列号...\n")
        new_col_to_logic_id = {}
        for c in range(1, max_c):
            for r in range(1, 30):
                val = ws_new.cell(row=r, column=c).value
                if val is not None:
                    v_str = str(val).strip()
                    # 精准抓取单元格里单独的 (28) 或 （28）
                    m = re.search(r'^[\(（]\s*(\d+)\s*[\)）]$', v_str)
                    if m:
                        new_col_to_logic_id[c] = int(m.group(1))
                        break # 找到了这列的列号就跳出内层循环

        log_widget.insert(tk.END, f"➔ 成功侦测到 {len(new_col_to_logic_id)} 个带有括号的列号。\n\n")

        # ==========================================
        # 3. 读取 CSV，映射物理坐标，并根据新表头强制改名
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在更新 CSV 映射...\n")
        
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            rows = list(csv.reader(f))
            
        final_rows = []
        updated_count = 0
        deleted_count = 0
        
        for row in rows:
            if len(row) < 3:
                final_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            # 只处理指定 Sheet 且 ID 是有效数字（长度至少4位，保证能切分出后两位）
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
                        
                        # 【核心突破】：抬头看一眼这列在最新的 Excel 里，表头数字是多少？
                        detected_logic_num = new_col_to_logic_id.get(new_col_idx)
                        
                        if detected_logic_num is not None:
                            # 保留前面的前缀 (比如 923)，直接替换最后两位列名
                            new_id_val = f"{id_val[:-2]}{detected_logic_num:02d}"
                        else:
                            # 万一这列没写括号数字，安全起见保持原 ID
                            new_id_val = id_val
                            
                        final_rows.append([new_id_val, sheet_val, new_cell])
                        
                        if new_id_val != id_val or new_cell != old_cell:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [修正] ID: {id_val} ➔ {new_id_val} | 坐标: {old_cell} ➔ {new_cell}\n")
                    else:
                        # 物理列/行被彻彻底底删除了
                        deleted_count += 1
                        log_widget.insert(tk.END, f" ❌ [删除] ID: {id_val} | 原坐标 {old_cell} 的物理列已被移除。\n")
                else:
                    final_rows.append(row)
            else:
                final_rows.append(row)
                
        # ==========================================
        # 4. 导出结果并替换
        # ==========================================
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(final_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n由于新增/删除列，成功修正并重命了 {updated_count} 个坐标及 ID。\n移除了 {deleted_count} 个失效项。\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 动态修正完成！\n修正数量: {updated_count} 个\n移除失效: {deleted_count} 个")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射智能修正工具 (AI 表头侦测版) v4.0")
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
        self.old_xl_entry = create_file_picker(root, "2. 请选择【原本】未修改的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        
        tk.Label(root, text="4. 请输入本次要更新的 Sheet 名称 (例如: 47):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        self.btn_start = tk.Button(root, text="🚀 侦测表头并修正 CSV 映射", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
        process_dynamic_mapping(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicUpdaterApp(app_root)
    app_root.mainloop()
