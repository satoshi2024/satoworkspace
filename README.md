import os
import re
import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl.utils.cell import get_column_letter

def process_pure_append(csv_path, excel_path, target_sheet, added_col, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/3】正在读取 Excel 寻找新列 ({added_col}) 的物理坐标...\n")
        log_widget.update()
        
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if target_sheet not in wb.sheetnames:
            messagebox.showerror("错误", f"Excel 中找不到 Sheet: [{target_sheet}]")
            return
            
        ws = wb[target_sheet]
        
        # 1. 在 Excel 中寻找追加列的物理列号 (支持全角/半角括号和数字)
        target_phys_col_letter = None
        for c in range(1, ws.max_column + 15):
            for r in range(1, 50):
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    # 将全角转半角，消除空格
                    v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                    m = re.search(r'^[\(（]\s*(\d+)\s*[\)）]$', v_str)
                    if m and int(m.group(1)) == int(added_col):
                        target_phys_col_letter = get_column_letter(c)
                        break
            if target_phys_col_letter:
                break
                
        if not target_phys_col_letter:
            messagebox.showerror("错误", f"在 Sheet [{target_sheet}] 的前 50 行内找不到 ({added_col}) 的表头！")
            return
            
        log_widget.insert(tk.END, f"➔ 成功锁定 ({added_col}) 列的物理坐标为: {target_phys_col_letter} 列\n\n")
        
        # 2. 读取 CSV，完全不修改原有数据，仅提取物理行锚点
        log_widget.insert(tk.END, f"【2/3】正在解析 CSV 基础数据...\n")
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            csv_rows = list(csv.reader(f))
            
        base_to_phys_row = {}
        existing_ids = set()
        
        for row in csv_rows:
            if len(row) >= 3 and row[1].strip() == str(target_sheet):
                id_val, cell_val = row[0].strip(), row[2].strip()
                existing_ids.add(id_val)
                
                # 提取如 '550142' 中的 '5501' 作为基准前缀
                if id_val.isdigit() and len(id_val) >= 4:
                    base_prefix = id_val[:-2] 
                    # 提取坐标中的物理行，例如 'DR15' -> '15'
                    match = re.search(r'\d+', cell_val)
                    if match:
                        base_to_phys_row[base_prefix] = match.group(0)
                        
        if not base_to_phys_row:
            messagebox.showerror("错误", f"CSV 中没有任何关于 Sheet [{target_sheet}] 的老数据，无法推算新数据的行号！")
            return
            
        # 3. 构造追加数据，并原位插队
        log_widget.insert(tk.END, f"【3/3】开始执行纯粹追加 (不修改任何既有数据)...\n")
        
        added_col_padded = str(added_col).zfill(2)
        appended_count = 0
        
        # 为了保证不打乱原 CSV 顺序，我们需要找到对应的位置并 insert
        for base_prefix, phys_row in base_to_phys_row.items():
            new_id = f"{base_prefix}{added_col_padded}"
            
            # 如果库里确实没有，才进行追加
            if new_id not in existing_ids:
                new_cell = f"{target_phys_col_letter}{phys_row}"
                new_row_data = [new_id, str(target_sheet), new_cell]
                
                # 寻找插入点：倒序遍历，找到 CSV 中最后一个属于 base_prefix (如 5501) 的行，插在它下面
                insert_idx = len(csv_rows)
                for i in range(len(csv_rows)-1, -1, -1):
                    if len(csv_rows[i]) >= 1 and str(csv_rows[i][0]).strip().startswith(base_prefix):
                        insert_idx = i + 1
                        break
                        
                csv_rows.insert(insert_idx, new_row_data)
                appended_count += 1
                log_widget.insert(tk.END, f" ➕ [追加成功] {new_id} , {target_sheet} , {new_cell} (无缝插入完毕)\n")
                
        # 4. 导出
        output_csv_path = csv_path.replace(".csv", "_Updated.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(csv_rows)
            
        shutil.copyfile(csv_path, csv_path + ".bak")
        shutil.copyfile(output_csv_path, csv_path)
        os.remove(output_csv_path)
        
        log_widget.insert(tk.END, f"\n✅ 【处理完成】\n完全没有修改原 CSV 的任何一行数据。\n成功追加了 {appended_count} 个新坐标！\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"追加完成！\n\n共新增 {appended_count} 行。\n原表未做任何修改与排序。")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

# ==========================================
# 极简 GUI 界面
# ==========================================
class PureAppenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("坐标映射追加专属工具 (纯粹追加版)")
        self.root.geometry("650x500")
        
        def create_file_picker(parent, label_text, file_types):
            tk.Label(parent, text=label_text, font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15)
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True, ipady=3)
            btn = tk.Button(frame, text="浏览...", width=10, command=lambda: self.browse_file(entry, file_types))
            btn.pack(side="right", padx=5)
            return entry

        # 仅需 2 个文件
        self.csv_entry = create_file_picker(root, "1. 请选择要追加的映射 CSV 文件:", [("CSV Files", "*.csv")])
        self.new_xl_entry = create_file_picker(root, "2. 请选择包含新列的 Excel 文件:", [("Excel Files", "*.xlsm *.xlsx")])
        
        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=15, pady=15)
        
        tk.Label(input_frame, text="3. 目标 Sheet (如 59):", font=("MS Gothic", 9, "bold")).pack(side="left")
        self.sheet_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(side="left", padx=5)
        
        tk.Label(input_frame, text="4. 需追加的新列号 (如 43):", font=("MS Gothic", 9, "bold")).pack(side="left", padx=(20, 0))
        self.col_entry = tk.Entry(input_frame, width=10, font=("Calibri", 11, "bold"), fg="red")
        self.col_entry.pack(side="left", padx=5)
        
        self.btn_start = tk.Button(root, text="🚀 仅执行追加 (不修复/不排序)", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
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
        new_xl = self.new_xl_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        add_col = self.col_entry.get().strip()
        
        if not all([csv_p, new_xl, sheet_n, add_col]):
            messagebox.showwarning("提示", "4 项输入均为必填！")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_pure_append(csv_p, new_xl, sheet_n, add_col, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = PureAppenderApp(app_root)
    app_root.mainloop()
