import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def process_row_logic_shift(csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载新旧 Excel 文件...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        # ==========================================
        # 1. 核心提取器：拼装被拆分的“行番号”并提取行指纹
        # ==========================================
        def get_row_data(ws, version_name):
            id_col_start = None
            id_row_start = None
            
            # 寻找 "行番号" 所在的列
            for r in range(1, 40):
                for c in range(1, 20):
                    val = ws.cell(row=r, column=c).value
                    if val and "行番号" in str(val).replace(" ", ""):
                        id_col_start = c
                        id_row_start = r
                        break
                if id_col_start: break
                
            phys_to_logic = {}
            signatures = []
            phys_rows = []
            
            skip_cols = []
            if id_col_start:
                skip_cols = [id_col_start, id_col_start+1, id_col_start+2]
            
            if id_col_start:
                # 往下扫，把相邻三列的数字拼起来
                for r in range(id_row_start + 1, ws.max_row + 1):
                    v1 = ws.cell(row=r, column=id_col_start).value
                    v2 = ws.cell(row=r, column=id_col_start + 1).value
                    v3 = ws.cell(row=r, column=id_col_start + 2).value
                    
                    s1 = str(v1).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip() if v1 is not None else ""
                    s2 = str(v2).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip() if v2 is not None else ""
                    s3 = str(v3).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip() if v3 is not None else ""
                    
                    logic_row = ""
                    if s1.isdigit(): logic_row += s1
                    if s2.isdigit(): logic_row += s2
                    if s3.isdigit(): logic_row += s3
                    
                    # 只有当这一行确实有“行番号”时，我们才提取它的指纹！
                    if logic_row: 
                        phys_to_logic[r] = logic_row
                        
                        vals = []
                        # 提取前面 20 列的文本作为这一行的“骨架”
                        for c in range(1, 20):
                            if c in skip_cols: continue
                            val = ws.cell(row=r, column=c).value
                            v_str = str(val).replace(" ", "").replace("　", "").replace("\n", "").strip() if val is not None else ""
                            # 核心：抹除所有数字，只留下汉字对比（例如“一般四輪乗用営業用”）
                            v_str = re.sub(r'\d+', '', v_str) 
                            vals.append(v_str)
                            
                        sig = "|".join(vals)
                        signatures.append(sig)
                        phys_rows.append(r)
                        
            log_widget.insert(tk.END, f"➔ [{version_name}] 成功解析 {len(phys_to_logic)} 个有效数据行。\n")
            return phys_to_logic, signatures, phys_rows

        log_widget.insert(tk.END, f"【2/4】正在解析行番号字典与文本骨架...\n")
        log_widget.update()
        
        old_phys_to_logic, old_sigs, old_phys_rows = get_row_data(ws_old, "旧版 Excel")
        new_phys_to_logic, new_sigs, new_phys_rows = get_row_data(ws_new, "新版 Excel")

        # ==========================================
        # 2. Diff 行文骨架，找出物理行的漂移映射
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在比对文本，定位物理行的漂移轨迹...\n")
        sm = difflib.SequenceMatcher(None, old_sigs, new_sigs)
        old_r_to_new_r = {}
        
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_idx, new_idx in zip(range(i1, i2), range(j1, j2)):
                    old_r = old_phys_rows[old_idx]
                    new_r = new_phys_rows[new_idx]
                    old_r_to_new_r[old_r] = new_r

        # ==========================================
        # 3. 读取 CSV，执行坐标更新与 ID 翻译
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在更新 CSV 坐标并翻译逻辑 ID...\n")
        
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        new_rows = []
        updated_coord_count = 0
        updated_id_count = 0
        deleted_count = 0

        for row in csv_data:
            if len(row) < 3:
                new_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            # 只处理指定的 Sheet
            if sheet_val == str(target_sheet):
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                if match:
                    col_str = match.group(1) # 列没变，直接继承 (例如 'AK')
                    old_phys_row = int(match.group(2)) # 例如 38
                    
                    # 查 Diff 字典：旧的 38 行变成了新版的第几行？
                    new_phys_row = old_r_to_new_r.get(old_phys_row)
                    
                    if new_phys_row:
                        new_cell = f"{col_str}{new_phys_row}"
                        
                        # 查行号字典：获取新老版本的逻辑行号 (例如 210 -> 200)
                        old_logic = old_phys_to_logic.get(old_phys_row)
                        new_logic = new_phys_to_logic.get(new_phys_row)
                        
                        new_id_val = id_val
                        # 如果行号变了，并且旧行号存在于 ID 中，执行替换！
                        if old_logic and new_logic and old_logic != new_logic:
                            # 逆向替换：替换 ID 中最后一次出现的行号 (防止前缀重复)
                            head, sep, tail = id_val.rpartition(old_logic)
                            if sep:
                                new_id_val = head + new_logic + tail
                                
                        new_rows.append([new_id_val, sheet_val, new_cell])
                        
                        if new_cell != old_cell or new_id_val != id_val:
                            log_msg = f" 🔄 [追踪] "
                            if new_id_val != id_val:
                                log_msg += f"ID: {id_val} ➔ {new_id_val} | "
                                updated_id_count += 1
                            if new_cell != old_cell:
                                log_msg += f"坐标: {old_cell} ➔ {new_cell}"
                                updated_coord_count += 1
                            log_widget.insert(tk.END, log_msg + "\n")
                    else:
                        deleted_count += 1
                        log_widget.insert(tk.END, f" ❌ [行删除] ID:{id_val} | 物理行 {old_phys_row} 在新版中被移除。\n")
                else:
                    new_rows.append(row)
            else:
                new_rows.append(row)

        # ==========================================
        # 4. 导出结果
        # ==========================================
        output_csv_path = csv_path.replace(".csv", "_Updated.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        shutil.copyfile(csv_path, csv_path + ".bak")
        shutil.copyfile(output_csv_path, csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n坐标发生漂移: {updated_coord_count} 个\n逻辑 ID 翻译更名: {updated_id_count} 个\n废弃删除: {deleted_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 行号与 ID 修正完成！\n\n坐标更新: {updated_coord_count}\nID 更名: {updated_id_count}\n\n(列未受影响，保留绝对原序)")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

# ==========================================
# 极简 GUI 界面
# ==========================================
class RowLogicTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射引擎 (V17 行位漂移+ID翻译专属版)")
        self.root.geometry("680x520")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        self.csv_entry = create_file_picker(root, "1. 请选择映射 CSV 文件:", [("CSV Files", "*.csv")])
        self.old_xl_entry = create_file_picker(root, "2. 请选择【原本】未修改的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        self.new_xl_entry = create_file_picker(root, "3. 请选择【改修后】最新的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        
        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=15, pady=15)
        
        tk.Label(input_frame, text="4. 需更新的目标 Sheet (如 95):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=15, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 开始精准修正行坐标与逻辑 ID", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=5, ipady=5)
        
        self.log_text = scrolledtext.ScrolledText(root, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=15)
        
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
            messagebox.showwarning("提示", "4 项输入均为必填！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_row_logic_shift(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = RowLogicTranslatorApp(app_root)
    app_root.mainloop()
