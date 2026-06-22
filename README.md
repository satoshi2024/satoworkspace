import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def process_v18_number_sequence(csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在安全加载新老 Excel 文件...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        # ==========================================
        # 1. 核心算法：提取纯“行番号”数字序列 (免疫合并单元格)
        # ==========================================
        def get_logic_numbers(ws, version_name):
            id_col_start = None
            id_row_start = None
            
            # 找 "行番号" 表头
            for r in range(1, 40):
                for c in range(1, 20):
                    val = ws.cell(row=r, column=c).value
                    if val and "行番号" in str(val).replace(" ", ""):
                        id_col_start = c
                        id_row_start = r
                        break
                if id_col_start: break
                
            logic_list = []
            logic_to_phys = {}
            
            if id_col_start:
                # 往下扫描数字
                for r in range(id_row_start + 1, ws.max_row + 1):
                    s1 = str(ws.cell(row=r, column=id_col_start).value or "").translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                    s2 = str(ws.cell(row=r, column=id_col_start+1).value or "").translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                    s3 = str(ws.cell(row=r, column=id_col_start+2).value or "").translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                    
                    logic_str = ""
                    if s1.isdigit(): logic_str += s1
                    if s2.isdigit(): logic_str += s2
                    if s3.isdigit(): logic_str += s3
                    
                    if logic_str and len(logic_str) >= 2:
                        # 核心规则：提取前两位作为逻辑 ID (例如 140 -> 14, 010 -> 01)
                        logic_code = logic_str[:2]
                        if logic_code not in logic_list:
                            logic_list.append(logic_code)
                        if logic_code not in logic_to_phys:
                            logic_to_phys[logic_code] = r
                            
            log_widget.insert(tk.END, f"➔ [{version_name}] 成功提取到 {len(logic_list)} 个逻辑行号。\n")
            return logic_list, logic_to_phys

        log_widget.insert(tk.END, f"【2/4】正在提取行番号数字序列...\n")
        log_widget.update()
        
        old_logic_list, old_logic_to_phys = get_logic_numbers(ws_old, "旧版 Excel")
        new_logic_list, new_logic_to_phys = get_logic_numbers(ws_new, "新版 Excel")

        # ==========================================
        # 2. 数字序列 Diff：找出 ID 的变更规则
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在比对行号，追踪逻辑 ID 变更...\n")
        sm = difflib.SequenceMatcher(None, old_logic_list, new_logic_list)
        logic_translation = {} # 旧逻辑行 -> 新逻辑行 (例如 '21' -> '20')
        
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            # 无论是一一对应(equal)还是被替换(replace)，都建立映射
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    old_logic = old_logic_list[old_i]
                    new_logic = new_logic_list[new_j]
                    logic_translation[old_logic] = new_logic

        # ==========================================
        # 3. 安全更新 CSV (绝不删除任何未匹配数据)
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在安全更新坐标并修正 ID...\n")
        
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_data = list(csv.reader(f))
            
        new_rows = []
        updated_count = 0
        unmatched_count = 0

        for row in csv_data:
            if len(row) < 3:
                new_rows.append(row)
                continue
                
            id_val, sheet_val, old_cell = row[0].strip(), row[1].strip(), row[2].strip()
            
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                # 从 ID 中提取逻辑行 (例如 331401 提取出 14)
                old_logic_code = id_val[-4:-2]
                new_logic_code = logic_translation.get(old_logic_code)
                
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                
                # 如果能找到映射，且坐标解析正常
                if new_logic_code and match:
                    new_phys_row = new_logic_to_phys.get(new_logic_code)
                    if new_phys_row:
                        # 继承老列字母，拼上新物理行
                        col_str = match.group(1)
                        new_cell = f"{col_str}{new_phys_row}"
                        
                        # 拼接新 ID
                        new_id_val = id_val[:-4] + new_logic_code + id_val[-2:]
                        
                        new_rows.append([new_id_val, sheet_val, new_cell])
                        
                        if new_cell != old_cell or new_id_val != id_val:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [安全修正] ID: {id_val} ➔ {new_id_val} | 坐标: {old_cell} ➔ {new_cell}\n")
                        continue
                
                # 【防误删铁律】：只要找不到映射，原样保留，绝对不删！
                unmatched_count += 1
                new_rows.append(row)
            else:
                new_rows.append(row)

        # ==========================================
        # 4. 导出覆盖
        # ==========================================
        output_csv_path = csv_path.replace(".csv", "_Updated.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        shutil.copyfile(csv_path, csv_path + ".bak")
        shutil.copyfile(output_csv_path, csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n成功追踪并修正: {updated_count} 个\n安全保留(未发生变更/未匹配): {unmatched_count} 个\n误删数量: 0 个（已开启绝对保护）\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 纯数字修正完成！\n\n成功修正: {updated_count}\n保留原样: {unmatched_count}\n误删数: 0\n\n请去 WinMerge 查看，绝无大片红区！")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")


class SafeRowUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标引擎 (V18 数字行号绝对安全版)")
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
        
        self.btn_start = tk.Button(root, text="🚀 启动绝对安全更新 (纯读数字/不删代码)", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
        process_v18_number_sequence(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = SafeRowUpdaterApp(app_root)
    app_root.mainloop()
