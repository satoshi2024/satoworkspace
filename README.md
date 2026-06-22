import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def process_v20_bulletproof_row_sync(csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
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
        # 1. 终极雷达：无视排版，暴力抓取数字序列
        # ==========================================
        def get_logic_numbers(ws, version_name):
            id_col_start = None
            id_row_start = None
            
            # 全景扫描前 100 行 x 100 列
            for r in range(1, 100):
                for c in range(1, 100):
                    val = ws.cell(row=r, column=c).value
                    if val:
                        clean_val = re.sub(r'\s+', '', str(val))
                        # 兼容横向换行或者连续文字
                        if "行番号" in clean_val:
                            id_col_start = c
                            id_row_start = r
                            break
                        # 兼容纵向拆分：如果只看到"行"，检查下面是不是"番"
                        elif clean_val == "行":
                            v_below = ws.cell(row=r+1, column=c).value
                            if v_below and "番" in re.sub(r'\s+', '', str(v_below)):
                                id_col_start = c
                                id_row_start = r + 2 # 跳过表头
                                break
                if id_col_start: break
                
            logic_list = []
            logic_to_phys = {}
            
            if id_col_start:
                # 往下扫描，横向吸尘器提取拆分的数字
                for r in range(id_row_start + 1, ws.max_row + 1):
                    logic_str = ""
                    # 连续检查右侧 6 个小格子，把所有数字吸过来
                    for offset in range(6):
                        v = ws.cell(row=r, column=id_col_start + offset).value
                        if v is not None:
                            s = str(v).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                            if s.isdigit():
                                logic_str += s
                                
                    if len(logic_str) >= 2:
                        # 提取前两位作为核心逻辑 ID (例如 '010' 提取 '01', '140' 提取 '14')
                        logic_code = logic_str[:2]
                        if logic_code not in logic_list:
                            logic_list.append(logic_code)
                        if logic_code not in logic_to_phys:
                            logic_to_phys[logic_code] = r
                            
            log_widget.insert(tk.END, f"➔ [{version_name}] 突破排版，成功抓取到 {len(logic_list)} 个逻辑行锚点！\n")
            return logic_list, logic_to_phys

        log_widget.insert(tk.END, f"【2/4】启动全景雷达，扫描行番号...\n")
        log_widget.update()
        
        old_logic_list, old_logic_to_phys = get_logic_numbers(ws_old, "旧版 Excel")
        new_logic_list, new_logic_to_phys = get_logic_numbers(ws_new, "新版 Excel")
        
        if not old_logic_list or not new_logic_list:
            messagebox.showerror("致命错误", "解析失败！未能找到'行番号'，请确保表单不为空且包含该字段。")
            return

        # ==========================================
        # 2. 数字序列 Diff：精准定位 ID 变更
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在比对序列，追踪逻辑 ID 与行高变更...\n")
        sm = difflib.SequenceMatcher(None, old_logic_list, new_logic_list)
        logic_translation = {} 
        
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            # 获取旧行号到新行号的对应关系
            if tag in ('equal', 'replace'):
                for old_i, new_j in zip(range(i1, i2), range(j1, j2)):
                    old_logic = old_logic_list[old_i]
                    new_logic = new_logic_list[new_j]
                    logic_translation[old_logic] = new_logic

        # ==========================================
        # 3. 安全更新 CSV 坐标与 ID
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在同步 CSV (绝对保护原数据)...\n")
        
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
            
            # 只处理目标页，并且 ID 是有效数字
            if sheet_val == str(target_sheet) and id_val.isdigit() and len(id_val) >= 4:
                # 切片提取逻辑行 (例如 331401 中的 '14', 90118 中的 '01')
                old_logic_code = id_val[-4:-2] 
                new_logic_code = logic_translation.get(old_logic_code)
                
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                
                # 只有完美匹配到了新的映射，才执行修改
                if new_logic_code and match:
                    new_phys_row = new_logic_to_phys.get(new_logic_code)
                    
                    if new_phys_row:
                        # 继承原本的列字母，替换为新算出的物理行
                        col_str = match.group(1)
                        new_cell = f"{col_str}{new_phys_row}"
                        
                        # 同步替换 ID 中的行号片段
                        new_id_val = id_val[:-4] + new_logic_code + id_val[-2:]
                        
                        new_rows.append([new_id_val, sheet_val, new_cell])
                        
                        if new_cell != old_cell or new_id_val != id_val:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [修正] ID: {id_val} ➔ {new_id_val} | 坐标: {old_cell} ➔ {new_cell}\n")
                        continue
                        
                # 保护机制：找不到对应项的，绝对不删，原样保留
                unmatched_count += 1
                new_rows.append(row)
            else:
                new_rows.append(row)

        # ==========================================
        # 4. 导出
        # ==========================================
        output_csv_path = csv_path.replace(".csv", "_Updated.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        shutil.copyfile(csv_path, csv_path + ".bak")
        shutil.copyfile(output_csv_path, csv_path)
        os.remove(output_csv_path)
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n行号与坐标追踪同步: {updated_count} 个\n原样保留/安全忽略: {unmatched_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 同步完成！\n\n成功追踪变更: {updated_count} 个。\n\n请在 WinMerge 中确认，修改绝对干净利落！")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

class RowSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标引擎 (V20 全景雷达·行变动追踪版)")
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
        
        tk.Label(input_frame, text="4. 需同步的目标 Sheet (如 104):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=15, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 启动全景雷达，追踪行变动", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
        process_v20_bulletproof_row_sync(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = RowSyncApp(app_root)
    app_root.mainloop()
