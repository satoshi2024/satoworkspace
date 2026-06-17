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
        log_widget.insert(tk.END, f"【2/4】正在分析 Excel 结构差异...\n")
        log_widget.update()

        max_c = max(ws_old.max_column, ws_new.max_column) + 10
        max_r = max(ws_old.max_row, ws_new.max_row) + 10

        def get_col_sig(ws, col_idx):
            vals = []
            for r in range(1, 40):  # 扫描前40行特征
                val = ws.cell(row=r, column=col_idx).value
                if val is not None:
                    v_str = str(val).strip()
                    # 核心突破：剔除带括号的数字 (18), （19），防止手工修改表头导致匹配失败！
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

        # 行对比（常规匹配）
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
        # 2. 读取旧 CSV，分组并推算新坐标
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在计算逻辑 ID 动态重排...\n")
        
        output_csv_path = old_csv_path.replace(".csv", "_Updated.csv")
        
        with open(old_csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            rows = list(csv.reader(f))
            
        target_group = {}
        
        # 将指定 Sheet 的数据按照前缀分组 (例如 '92318' 分组为前缀 '923'，逻辑列 18)
        for row in rows:
            if len(row) >= 3:
                id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
                if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                    group_prefix = id_val[:-2]
                    col_num = int(id_val[-2:])
                    
                    match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                    survived = False
                    new_cell = old_cell
                    new_col_idx = 9999
                    
                    if match:
                        col_str, row_str = match.groups()
                        old_col_idx = column_index_from_string(col_str)
                        old_row_idx = int(row_str)
                        
                        new_c_idx = col_map.get(old_col_idx)
                        new_r_idx = row_map.get(old_row_idx)
                        
                        if new_c_idx and new_r_idx:
                            survived = True
                            new_col_idx = new_c_idx
                            new_cell = f"{get_column_letter(new_c_idx)}{new_r_idx}"
                            
                    if group_prefix not in target_group:
                        target_group[group_prefix] = []
                        
                    target_group[group_prefix].append({
                        'original_row': row,
                        'col_num': col_num,
                        'survived': survived,
                        'new_col_idx': new_col_idx,
                        'new_cell': new_cell,
                        'id_val': id_val,
                        'sheet_val': sheet_val
                    })

        # ==========================================
        # 3. 按物理顺序重新分配逻辑 ID，并按原顺序合并 CSV
        # ==========================================
        final_rows = []
        processed_groups = set()
        updated_count = 0
        deleted_count = 0
        
        for row in rows:
            if len(row) < 3:
                final_rows.append(row)
                continue
                
            id_val, sheet_val = row[0].strip(), row[1].strip()
            
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                group_prefix = id_val[:-2]
                
                # 同一个前缀组，只在遇到第一行时进行批量插入，保证 CSV 排序不乱
                if group_prefix not in processed_groups:
                    items = target_group[group_prefix]
                    original_min_col = min(item['col_num'] for item in items)
                    
                    # 过滤出存活的列，并【核心】根据新 Excel 从左到右的物理顺序排序
                    surviving_items = [item for item in items if item['survived']]
                    surviving_items.sort(key=lambda x: x['new_col_idx'])
                    
                    deleted_count += (len(items) - len(surviving_items))
                    
                    # 重新分配逻辑 ID (18, 19, 20...)
                    for idx, item in enumerate(surviving_items):
                        new_col_num = original_min_col + idx
                        new_id = f"{group_prefix}{new_col_num:02d}"
                        final_rows.append([new_id, item['sheet_val'], item['new_cell']])
                        
                        if new_id != item['id_val'] or item['new_cell'] != item['original_row'][2].strip():
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [修正] {item['id_val']} ({item['original_row'][2]}) ➔ {new_id} ({item['new_cell']})\n")
                            
                    processed_groups.add(group_prefix)
            else:
                final_rows.append(row)
                
        # ==========================================
        # 4. 导出结果
        # ==========================================
        log_widget.insert(tk.END, f"\n【4/4】正在保存结果至: {os.path.basename(output_csv_path)}\n")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(final_rows)
            
        shutil.copyfile(old_csv_path, old_csv_path + ".bak")
        shutil.copyfile(output_csv_path, old_csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n逻辑重排并偏移了 {updated_count} 个坐标\n因物理删除移除了 {deleted_count} 个旧映射。\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 动态修正完成！\n偏移并重排 ID: {updated_count} 个\n移除失效项: {deleted_count} 个")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

# ==========================================
# GUI 界面类保持不变
# ==========================================
class DynamicUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射动态追踪工具 (AI 自动重排版) v3.0")
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
        
        tk.Label(root, text="4. 请输入本次要更新的 Sheet 名称 (例如: 43):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        self.btn_start = tk.Button(root, text="🚀 开始智能比对并修正 CSV 坐标", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
