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
        
        # 载入新旧 Excel
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在某个 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        # ==========================================
        # 算法核心：提取列特征并进行差异比对 (Diff)
        # ==========================================
        log_widget.insert(tk.END, f"【2/4】正在对比新旧 Excel 的物理结构差异...\n")
        log_widget.update()

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        # 生成列指纹的函数 (取前50行内容作为该列的唯一特征)
        def get_col_sig(ws, col_idx):
            vals = []
            for r in range(1, 51):
                val = ws.cell(row=r, column=col_idx).value
                # 去除空格以提高容错率
                val_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                vals.append(val_str)
            return "|".join(vals)

        old_col_sigs = [get_col_sig(ws_old, c) for c in range(1, max_c)]
        new_col_sigs = [get_col_sig(ws_new, c) for c in range(1, max_c)]

        # 使用 SequenceMatcher 计算列的移动路径
        sm_col = difflib.SequenceMatcher(None, old_col_sigs, new_col_sigs)
        col_map = {}  # { 旧列索引 : 新列索引 }
        
        for tag, i1, i2, j1, j2 in sm_col.get_opcodes():
            # equal: 完全一致的列; replace: 内容发生轻微修改的列 (例如 19 改成了 18)
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    col_map[old_i + 1] = new_j + 1  # openpyxl 索引从 1 开始

        # 同理，对行也做一次 Diff（防止用户不仅删了列，还删了行）
        def get_row_sig(ws, row_idx):
            vals = []
            for c in range(1, 51):
                val = ws.cell(row=row_idx, column=c).value
                val_str = str(val).replace(" ", "").replace("　", "").strip() if val is not None else ""
                vals.append(val_str)
            return "|".join(vals)

        old_row_sigs = [get_row_sig(ws_old, r) for r in range(1, max_r)]
        new_row_sigs = [get_row_sig(ws_new, r) for r in range(1, max_r)]
        sm_row = difflib.SequenceMatcher(None, old_row_sigs, new_row_sigs)
        row_map = {} # { 旧行索引 : 新行索引 }
        
        for tag, i1, i2, j1, j2 in sm_row.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    row_map[old_i + 1] = new_j + 1

        log_widget.insert(tk.END, "➔ 结构对比完成！已锁定所有行列的偏移坐标。\n\n")

        # ==========================================
        # 读取旧 CSV 并生成新映射
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在读取 原本CSV 并推算新坐标...\n")
        
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        
        rows = []
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
        new_rows = []
        updated_count = 0
        deleted_count = 0
        
        for row in rows:
            if len(row) < 3:
                new_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            # 只处理指定 Sheet 的数据
            if sheet_val == str(target_sheet):
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                if match:
                    col_str, row_str = match.groups()
                    old_col_idx = column_index_from_string(col_str)
                    old_row_idx = int(row_str)
                    
                    # 通过刚才算出的字典，查询新坐标
                    new_col_idx = col_map.get(old_col_idx)
                    new_row_idx = row_map.get(old_row_idx)
                    
                    if new_col_idx and new_row_idx:
                        new_col_str = get_column_letter(new_col_idx)
                        new_cell = f"{new_col_str}{new_row_idx}"
                        
                        new_rows.append([id_val, sheet_val, new_cell])
                        if old_cell != new_cell:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [坐标漂移] ID:{id_val} | 物理坐标: {old_cell} ➔ {new_cell}\n")
                        else:
                            # 坐标没变的也原样保留
                            pass 
                    else:
                        # 如果在 col_map 里找不到，说明这一列被用户物理删除了！
                        deleted_count += 1
                        log_widget.insert(tk.END, f" ❌ [结构删除] ID:{id_val} | 原坐标 {old_cell} 所在的物理区域已被删除。\n")
                else:
                    new_rows.append(row)
            else:
                # 其他 Sheet 的数据原样保留
                new_rows.append(row)
                
        # ==========================================
        # 导出结果
        # ==========================================
        log_widget.insert(tk.END, f"\n【4/4】正在保存结果至: {os.path.basename(output_csv_path)}\n")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        # 安全机制：备份并覆盖原文件
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n由于结构变动偏移了 {updated_count} 个坐标\n因结构删除移除了 {deleted_count} 个旧映射。\n原CSV已更新（自动生成了.bak备份）。\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 动态修正完成！\n偏移坐标: {updated_count} 个\n废弃失效: {deleted_count} 个")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

# ==========================================
# GUI 界面类 (新增4个输入框)
# ==========================================
class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射动态追踪工具 (双模对比版) v2.0")
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

        # 1. 原始 CSV
        self.csv_entry = create_file_picker(root, "1. 请选择【原本】的映射 CSV 文件 (例: DafKazeiWrt.csv):", [("CSV Files", "*.csv")])
        # 2. 原始 Excel
        self.old_xl_entry = create_file_picker(root, "2. 请选择【原本】未修改的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        # 3. 修改后 Excel
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件 (.xlsm/.xlsx):", [("Excel Files", "*.xlsm *.xlsx")])
        
        # 4. Sheet 名称
        tk.Label(root, text="4. 请输入本次要更新的 Sheet 名称 (例如: 43):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        # 5. 执行按钮
        self.btn_start = tk.Button(root, text="🚀 开始智能比对并修正 CSV 坐标", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        # 6. 日志
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
