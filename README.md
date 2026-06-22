import os
import re
import csv
import shutil
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def process_v21_ultimate_row_anchor(csv_path, old_excel_path, new_excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/4】正在加载新老 Excel 文件...\n")
        log_widget.update()
        
        wb_old = openpyxl.load_workbook(old_excel_path, data_only=True)
        wb_new = openpyxl.load_workbook(new_excel_path, data_only=True)
        
        if target_sheet not in wb_old.sheetnames or target_sheet not in wb_new.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{target_sheet}]")
            return
            
        ws_old = wb_old[target_sheet]
        ws_new = wb_new[target_sheet]

        # ==========================================
        # 1. 终极解析器：提取文本指纹和完整的行番号
        # ==========================================
        def get_row_signatures(ws, version_name):
            id_col_start = None
            id_row_start = None
            
            # 无视排版，寻找 "行番号" 坐标
            for r in range(1, 40):
                for c in range(1, 20):
                    val = str(ws.cell(r, c).value or "").replace(" ", "").replace("　", "").replace("\n", "")
                    if "行番号" in val:
                        id_col_start = c
                        id_row_start = r
                        break
                    elif val == "行": # 兼容纵向拆分
                        v_below = str(ws.cell(r+1, c).value or "").replace(" ", "").replace("　", "").replace("\n", "")
                        if "番" in v_below:
                            id_col_start = c
                            id_row_start = r + 2
                            break
                if id_col_start: break
                
            phys_to_logic = {}
            sigs = []
            phys_rows = []
            
            if id_col_start:
                for r in range(id_row_start + 1, ws.max_row + 1):
                    # 把紧邻的 4 列的数字全部吸过来拼成完整的行号 (如 2+1+0 = 210)
                    logic_str = ""
                    for offset in range(4):
                        v = ws.cell(row=r, column=id_col_start + offset).value
                        if v is not None:
                            s = str(v).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                            if s.isdigit():
                                logic_str += s
                                
                    if len(logic_str) >= 2:
                        phys_to_logic[r] = logic_str
                        
                        # 提取这一行的汉字骨架，作为匹配依据 (跳过行号数字)
                        row_text = ""
                        for c in range(1, 25):
                            if id_col_start <= c <= id_col_start + 3:
                                continue
                            val = ws.cell(row=r, column=c).value
                            if val is not None:
                                s = str(val).replace(" ", "").replace("　", "").replace("\n", "").replace("\r", "")
                                s = re.sub(r'\d+', '', s) # 彻底抹除数字，防止新版数字变动干扰
                                row_text += s
                                
                        sigs.append(row_text)
                        phys_rows.append(r)
                        
            log_widget.insert(tk.END, f"➔ [{version_name}] 成功提取 {len(phys_to_logic)} 个有效行锚点。\n")
            return phys_to_logic, sigs, phys_rows

        log_widget.insert(tk.END, f"【2/4】正在解析行番号字典与文本骨架...\n")
        log_widget.update()
        
        old_phys_to_logic, old_sigs, old_phys_rows = get_row_signatures(ws_old, "旧版 Excel")
        new_phys_to_logic, new_sigs, new_phys_rows = get_row_signatures(ws_new, "新版 Excel")
        
        if not old_phys_to_logic or not new_phys_to_logic:
            messagebox.showerror("错误", "解析失败！未能找到行番号。")
            return

        # ==========================================
        # 2. Diff 文本骨架，锁定物理行的准确漂移
        # ==========================================
        log_widget.insert(tk.END, f"【3/4】正在比对文本，定位物理行的漂移轨迹...\n")
        sm = difflib.SequenceMatcher(None, old_sigs, new_sigs)
        old_row_to_new_row = {}
        
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag in ('equal', 'replace'):
                for old_idx, new_idx in zip(range(i1, i2), range(j1, j2)):
                    old_r = old_phys_rows[old_idx]
                    new_r = new_phys_rows[new_idx]
                    old_row_to_new_row[old_r] = new_r

        # ==========================================
        # 3. 读取 CSV，根据老坐标物理锚点进行绝杀替换
        # ==========================================
        log_widget.insert(tk.END, f"【4/4】正在更新 CSV 坐标并翻译逻辑 ID...\n")
        
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
            
            if sheet_val == str(target_sheet):
                match = re.match(r"([a-zA-Z]+)(\d+)", old_cell)
                if match:
                    col_str = match.group(1)
                    old_phys_r = int(match.group(2)) # 提取老坐标的物理行，例如 47
                    
                    # 去字典里查：以前的第 47 行，现在跑到第几行了？
                    new_phys_r = old_row_to_new_row.get(old_phys_r)
                    
                    if new_phys_r:
                        # 组装新坐标：列不变，换新行 (AK47 -> AK46)
                        new_cell = f"{col_str}{new_phys_r}"
                        
                        # 获取行号变更字典：例如老行号是 210，新行号是 200
                        old_logic = old_phys_to_logic.get(old_phys_r)
                        new_logic = new_phys_to_logic.get(new_phys_r)
                        
                        new_id_val = id_val
                        # 如果行号发生了变动
                        if old_logic and new_logic and old_logic != new_logic:
                            # 绝杀替换：把 ID 里最后的旧行号(210)精准替换成新行号(200)
                            prefix = id_val[:-2] # 截去列号
                            suffix = id_val[-2:] # 保留列号
                            if prefix.endswith(old_logic):
                                new_prefix = prefix[:-len(old_logic)] + new_logic
                                new_id_val = new_prefix + suffix
                                
                        new_rows.append([new_id_val, sheet_val, new_cell])
                        
                        if new_cell != old_cell or new_id_val != id_val:
                            updated_count += 1
                            log_widget.insert(tk.END, f" 🔄 [绝杀修正] ID: {id_val} ➔ {new_id_val} | 坐标: {old_cell} ➔ {new_cell}\n")
                        continue
                        
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
            
        log_widget.insert(tk.END, f"\n✅ 【处理成功】\n行号与坐标追踪同步: {updated_count} 个\n原样保留/安全忽略: {unmatched_count} 个\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 同步完成！\n\n成功绝杀修正: {updated_count} 个。\n\n请在 WinMerge 中确认战果！")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

class UltimateRowSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标引擎 (V21 物理锚点绝杀版)")
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
        
        tk.Label(input_frame, text="4. 需同步的目标 Sheet (如 95):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=15, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 启动坐标与 ID 绝杀修正", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
        process_v21_ultimate_row_anchor(csv_p, old_xl, new_xl, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = UltimateRowSyncApp(app_root)
    app_root.mainloop()
