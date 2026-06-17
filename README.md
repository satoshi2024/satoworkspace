import os
import re
import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def process_mapping(csv_path, excel_path, target_sheet, log_widget):
    try:
        log_widget.insert(tk.END, f"【开始处理】正在打开 Excel 文件: {os.path.basename(excel_path)} ...\n")
        log_widget.see(tk.END)
        log_widget.update()
        
        # 载入 Excel (data_only=True 确保读取的是值而非公式)
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if target_sheet not in wb.sheetnames:
            messagebox.showerror("错误", f"Excel 中找不到指定的 Sheet: {target_sheet}")
            return
            
        ws = wb[target_sheet]
        
        # ==========================================
        # 1. 动态扫描 Excel 建立【行号/列号 -> 物理坐标】的映射
        # ==========================================
        log_widget.insert(tk.END, f"正在扫描 Sheet [{target_sheet}] 的行番号和列标题...\n")
        row_map = {}  # '01' -> Excel行索引 (int)
        col_map = {}  # 18 -> Excel列字母 (str)
        
        # 寻找 "行番号" 所在的物理列
        id_col_idx = None
        for r in range(1, 30):
            for c in range(1, 100):
                val = ws.cell(row=r, column=c).value
                if val and "行番号" in str(val):
                    id_col_idx = c
                    break
            if id_col_idx:
                break
                
        if id_col_idx:
            # 在该列向下扫描所有的行号 (如 01, 02 ... 24)
            for r in range(1, ws.max_row + 1):
                val = ws.cell(row=r, column=id_col_idx).value
                if val is not None:
                    v_str = str(val).strip()
                    if v_str.isdigit():
                        row_map[v_str.zfill(2)] = r
        else:
            log_widget.insert(tk.END, "警告: 未明确找到 '行番号' 标题列，启用全表数字行扫描...\n")
            for r in range(1, min(ws.max_row + 1, 100)):
                for c in range(1, 15):
                    val = ws.cell(row=r, column=c).value
                    if val is not None and str(val).strip().isdigit():
                        v_str = str(val).strip().zfill(2)
                        if 1 <= int(v_str) <= 30 and v_str not in row_map:
                            row_map[v_str] = r

        # 扫描前 20 行，精确抓取经过你删除修改后的新列标题 (18, 19, 20, 21, 22)
        for r in range(1, 20):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    # 消除空格和括号干扰，进行精准匹配
                    val_str_clean = str(val).replace(" ", "").replace("　", "").replace("(", "").replace(")", "").replace("（", "").replace("）", "")
                    if val_str_clean in ["18", "19", "20", "21", "22"]:
                        num_int = int(val_str_clean)
                        col_letter = openpyxl.utils.get_column_letter(c)
                        if num_int not in col_map:
                            col_map[num_int] = col_letter

        log_widget.insert(tk.END, f"➔ Excel扫描成功！定位有效行数: {len(row_map)} 行\n")
        log_widget.insert(tk.END, f"➔ 定位新列坐标: {col_map}\n\n")
        
        if not col_map or 18 not in col_map:
            messagebox.showerror("错误", "未能成功定位到新的 (18) 列标题。\n请确保在执行前，您已经手工删除了旧的18列！")
            return

        # ==========================================
        # 2. 读取并备份 CSV 文件
        # ==========================================
        csv_backup = csv_path + ".bak"
        shutil.copyfile(csv_path, csv_backup)
        log_widget.insert(tk.END, f"【安全备份】已创建原 CSV 备份文件: {os.path.basename(csv_backup)}\n")
        
        rows = []
        # 使用 utf-8-sig 防止日文字符集在 Excel 打开时乱码
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
        new_rows = []
        deleted_count = 0
        updated_count = 0
        
        log_widget.insert(tk.END, f"正在定向更新 CSV 数据 (仅过滤 Sheet 为 [{target_sheet}] 且属于第九表的数据)...\n")
        
        for row in rows:
            if len(row) < 3:
                new_rows.append(row)
                continue
                
            id_val, sheet_val, cell_val = row[0].strip(), row[1].strip(), row[2].strip()
            
            # 核心过滤逻辑：只有当前指定的 Sheet 并且以 '9' 开头的 5 位 ID 才会触发修改流程
            if sheet_val == str(target_sheet) and id_val.startswith('9') and len(id_val) == 5:
                table = id_val[0]          # '9'
                row_num = id_val[1:3]      # 行号字符串，如 '01'
                col_num = int(id_val[3:5]) # 逻辑列号数字，如 18, 19
                
                if col_num == 18:
                    # 旧 18 列废弃，直接移除此行
                    deleted_count += 1
                    log_widget.insert(tk.END, f" ❌ [删除旧项] ID={id_val} (原坐标 {cell_val})\n")
                    continue
                elif col_num > 18:
                    # 19 变成 18，20 变成 19...
                    new_col_num = col_num - 1
                    new_id = f"{table}{row_num}{new_col_num:02d}"
                    
                    # 从 Excel 的扫描结果里动态获取物理坐标
                    excel_row = row_map.get(row_num)
                    excel_col = col_map.get(new_col_num)
                    
                    if excel_row and excel_col:
                        new_cell = f"{excel_col}{excel_row}"
                        updated_count += 1
                        log_widget.insert(tk.END, f" 🔄 [更名平移] ID: {id_val} ➔ {new_id} | 坐标: {cell_val} ➔ {new_cell}\n")
                        new_rows.append([new_id, sheet_val, new_cell])
                    else:
                        # 容错：如果在 Excel 里没扫到对应行号，保持原样不让数据丢失
                        log_widget.insert(tk.END, f" ⚠️ [无法定位] 未在Excel中找到行号 {row_num} 列号 {new_col_num} 的物理坐标，保持原值\n")
                        new_rows.append([new_id, sheet_val, cell_val])
                else:
                    # 小于 18 的列（如果有的话）不受影响，保留
                    new_rows.append(row)
            else:
                # 其他所有不属于当前指定的 Sheet，或者非第九表的数据，原样安全保留
                new_rows.append(row)
                
        # ==========================================
        # 3. 将新数据写回 CSV
        # ==========================================
        with open(csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
            
        log_widget.insert(tk.END, f"\n【处理成功】\n累计删除废弃行: {deleted_count} 行\n累计更新坐标行: {updated_count} 行\n新的映射已成功写入原 CSV。\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Sheet [{target_sheet}] 处理完成！\n删除旧18列: {deleted_count} 行\n更新平移后续列: {updated_count} 行")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"处理过程中发生异常:\n{str(e)}")

# ==========================================
# GUI 界面类
# ==========================================
class MiniUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("第九表映射坐标自动转换工具 v1.0")
        self.root.geometry("700x550")
        
        # 1. CSV 文件选择
        tk.Label(root, text="1. 请选择完整的映射 CSV 文件 (dafkazeiwrt.csv):", font=("MS Gothic", 10, "bold")).pack(anchor="w", padx=15, pady=8)
        self.csv_frame = tk.Frame(root)
        self.csv_frame.pack(fill="x", padx=15)
        self.csv_entry = tk.Entry(self.csv_frame, font=("Calibri", 10))
        self.csv_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(self.csv_frame, text="浏览...", width=10, command=self.browse_csv).pack(side="right", padx=5)
        
        # 2. Excel 文件选择
        tk.Label(root, text="2. 请选择已经手工删除完18列的 Excel 文件 (.xlsm):", font=("MS Gothic", 10, "bold")).pack(anchor="w", padx=15, pady=8)
        self.excel_frame = tk.Frame(root)
        self.excel_frame.pack(fill="x", padx=15)
        self.excel_entry = tk.Entry(self.excel_frame, font=("Calibri", 10))
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(self.excel_frame, text="浏览...", width=10, command=self.browse_excel).pack(side="right", padx=5)
        
        # 3. 指定 Sheet 名输入
        tk.Label(root, text="3. 请输入本次要更新的 Sheet 名称 (例如: 43):", font=("MS Gothic", 10, "bold")).pack(anchor="w", padx=15, pady=8)
        self.sheet_entry = tk.Entry(root, width=25, font=("Calibri", 11, "bold"), fg="blue")
        self.sheet_entry.pack(anchor="w", padx=15, pady=2)
        
        # 4. 执行按钮
        self.btn_start = tk.Button(root, text="🚀 开始读取 Excel 并更新指定 Sheet 的 CSV 映射", bg="#0078D7", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        # 5. 实时日志显示视窗
        tk.Label(root, text="运行日志控制台:", font=("MS Gothic", 9)).pack(anchor="w", padx=15)
        self.log_text = scrolledtext.ScrolledText(root, height=15, font=("Consolas", 9.5), bg="#F3F3F3")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=5)
        
    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.csv_entry.delete(0, tk.END)
            self.csv_entry.insert(0, path)
            
    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)
            
    def start_process(self):
        csv_p = self.csv_entry.get().strip()
        excel_p = self.excel_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        
        if not csv_p or not excel_p or not sheet_n:
            messagebox.showwarning("提示", "请完整填写 CSV 路径、Excel 路径以及目标 Sheet 名称！")
            return
            
        if not os.path.exists(csv_p) or not os.path.exists(excel_p):
            messagebox.showerror("错误", "选择的文件路径不存在，请重新检查。")
            return
            
        self.log_text.delete(1.0, tk.END)
        process_mapping(csv_p, excel_p, sheet_n, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = MiniUpdaterApp(app_root)
    app_root.mainloop()
