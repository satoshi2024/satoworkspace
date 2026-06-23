import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def generate_matrix_pattern_spec(csv_path, excel_path, sheet_name, scan_mode, start_idx, item_template, csv_id_template, item_type, max_chars, align, font_size, filter_invalid, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/3】正在安全加载 CSV 映射库...\n")
        log_widget.update()
        
        csv_dict = {}
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 3:
                    c_id, c_sheet, c_cell = row[0].strip(), row[1].strip(), row[2].strip()
                    csv_dict[c_id] = {"sheet": c_sheet, "cell": c_cell}
                    
        log_widget.insert(tk.END, f"➔ 成功加载 {len(csv_dict)} 条坐标数据。\n\n")

        log_widget.insert(tk.END, f"【2/3】正在启动全景雷达扫描 Excel...\n")
        log_widget.update()
        
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("错误", f"找不到 Sheet: [{sheet_name}]")
            return
        ws = wb[sheet_name]
        
        extracted_rows = [] 
        extracted_cols = []

        # ==========================================
        # 1. 扫描：行番号 (支持 3 位解析，保留全串与2位切片)
        # ==========================================
        if scan_mode in ["ROW", "MATRIX"]:
            id_col_start, id_row_start = None, None
            for r in range(1, 100):
                for c in range(1, 100):
                    val = str(ws.cell(r, c).value or "").replace(" ", "").replace("　", "").replace("\n", "")
                    if "行番号" in val and len(val) <= 10:
                        id_col_start, id_row_start = c, r
                        break
                    elif val == "行": 
                        v_b = str(ws.cell(r+1, c).value or "").replace(" ", "").replace("　", "").replace("\n", "")
                        if "番" in v_b:
                            id_col_start, id_row_start = c, r + 2
                            break
                if id_col_start: break
                
            if id_col_start:
                for r in range(id_row_start + 1, ws.max_row + 1):
                    logic_str = ""
                    for offset in range(8): 
                        v = ws.cell(row=r, column=id_col_start + offset).value
                        if v is not None:
                            s = str(v).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                            if s.isdigit(): logic_str += s
                    if len(logic_str) >= 2:
                        if logic_str not in [x["raw"] for x in extracted_rows]:
                            extracted_rows.append({
                                "raw": logic_str,       # 原样数据，如 "010"
                                "short": logic_str[:2]  # 强制切片2位，如 "01"
                            })

        # ==========================================
        # 2. 扫描：列表头 (支持折叠排版穿透)
        # ==========================================
        if scan_mode in ["COL", "MATRIX"]:
            for c in range(1, ws.max_column + 20):
                for r in range(1, 100):
                    val = ws.cell(row=r, column=c).value
                    if val is not None:
                        lines = str(val).split('\n')
                        v_str = lines[0].translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                        m = re.search(r'^[\(（](\d+)[\)）]', v_str)
                        if m:
                            num = m.group(1)
                            if num not in [x["raw"] for x in extracted_cols]:
                                extracted_cols.append({
                                    "raw": num,         # 原样数字，如 "1"
                                    "short": num.zfill(2) # 自动补零2位，如 "01"
                                })
            extracted_cols = sorted(extracted_cols, key=lambda x: int(x["raw"]))

        log_widget.insert(tk.END, f"➔ 扫到 {len(extracted_rows)} 个行号，{len(extracted_cols)} 个列号。\n\n")

        # ==========================================
        # 3. 组合降维与跨界查表
        # ==========================================
        combinations = []
        if scan_mode == "ROW":
            combinations = [{"r": r, "c": None} for r in extracted_rows]
        elif scan_mode == "COL":
            combinations = [{"r": None, "c": c} for c in extracted_cols]
        elif scan_mode == "MATRIX":
            combinations = [{"r": r, "c": c} for r in extracted_rows for c in extracted_cols]

        if not combinations:
            messagebox.showerror("错误", "未能提取到有效数据！")
            return

        log_widget.insert(tk.END, f"【3/3】正在现取坐标，拼装 Pattern 表...\n")
        output_rows = [["No", "項目名", "選択項目No", "タイプ種别", "最大文字数", "文字配置", "領域外の対処", "文字フォント/サイズ", "編集箇所", "編集方法"]]
        
        current_idx = int(start_idx)
        found_count = 0
        skipped_count = 0
        
        for item in combinations:
            r_raw = item["r"]["raw"] if item["r"] else ""
            r_short = item["r"]["short"] if item["r"] else ""
            c_raw = item["c"]["raw"] if item["c"] else ""
            c_short = item["c"]["short"] if item["c"] else ""

            # 解耦替换：项目名用原样三位 {行全}，CSV ID查询用切片两位 {行2位}
            item_name = item_template.replace("{行全}", r_raw).replace("{行2位}", r_short).replace("{列}", c_raw).replace("{列2位}", c_short)
            search_id = csv_id_template.replace("{行全}", r_raw).replace("{行2位}", r_short).replace("{列}", c_raw).replace("{列2位}", c_short)
            
            if search_id in csv_dict:
                c_sheet, c_cell = csv_dict[search_id]["sheet"], csv_dict[search_id]["cell"]
                edit_method = f"表行列: S{search_id}\nシート番号: {c_sheet}\n座標: {c_cell}"
                found_count += 1
                log_widget.insert(tk.END, f" 🎯 完美转换 ➔ 项目名: {item_name} | 查表ID: {search_id} ➔ 坐标: {c_cell}\n")
            else:
                if filter_invalid:
                    skipped_count += 1
                    continue 
                edit_method = f"表行列: S{search_id}\nシート番号: 未知\n座標: 未知"
                log_widget.insert(tk.END, f" ⚠️ 未找到: [{search_id}] (字典中不存在)\n")
            
            output_rows.append([str(current_idx), item_name, "", item_type, max_chars, align, "出力しない", font_size, "-", edit_method])
            current_idx += 1

        output_csv_path = os.path.join(os.path.dirname(excel_path), f"Pattern_Spec_Sheet{sheet_name}_{scan_mode}.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerows(output_rows)
            
        log_widget.insert(tk.END, f"\n🎉 【大获全胜】\n成功写入式样书: {found_count} 行 | 智能拦截空白格子: {skipped_count} 行\n文件已保存: {os.path.basename(output_csv_path)}\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"式样书生成完成！\n\n成功写入 {found_count} 行有效数据。\n已完美隔离 3 位行号与 2 位 ID！")
        
    except Exception as e:
        messagebox.showerror("系统错误", str(e))

class MatrixPatternApp:
    def __init__(self, root):
        self.root = root
        self.root.title("测试式样书生成器 (V9 行列解耦终极版)")
        self.root.geometry("720x800")
        
        def create_input_row(parent, label_text, default_val=""):
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15, pady=4)
            tk.Label(frame, text=label_text, width=22, anchor="w", font=("MS Gothic", 9, "bold")).pack(side="left")
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True)
            entry.insert(0, default_val)
            return entry

        tk.Label(root, text="1. 文件选择:", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(5, 2))
        csv_frame = tk.Frame(root)
        csv_frame.pack(fill="x", padx=15, pady=2)
        self.csv_entry = tk.Entry(csv_frame, font=("Calibri", 10))
        self.csv_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(csv_frame, text="映射CSV...", width=10, command=lambda: self.browse_file(self.csv_entry)).pack(side="right", padx=5)
        
        excel_frame = tk.Frame(root)
        excel_frame.pack(fill="x", padx=15, pady=2)
        self.excel_entry = tk.Entry(excel_frame, font=("Calibri", 10))
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(excel_frame, text="最新Excel...", width=10, command=lambda: self.browse_file(self.excel_entry)).pack(side="right", padx=5)

        tk.Label(root, text="2. 扫描模式:", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        mode_frame = tk.Frame(root)
        mode_frame.pack(fill="x", padx=15)
        self.mode_var = tk.StringVar(value="MATRIX") 
        tk.Radiobutton(mode_frame, text="【单扫行】", variable=self.mode_var, value="ROW", command=self.update_templates).pack(side="left", padx=5)
        tk.Radiobutton(mode_frame, text="【单扫列】", variable=self.mode_var, value="COL", command=self.update_templates).pack(side="left", padx=5)
        tk.Radiobutton(mode_frame, text="【行列交叉(2D)】", variable=self.mode_var, value="MATRIX", command=self.update_templates, fg="red", font=("MS Gothic", 9, "bold")).pack(side="left", padx=5)

        self.lbl_template_hint = tk.Label(root, text="3. 模板配置 ({行全}=原样010, {行2位}=切片01, {列}=1, {列2位}=补零01):", font=("MS Gothic", 9, "bold"), fg="blue")
        self.lbl_template_hint.pack(anchor="w", padx=15, pady=(10, 2))
        
        self.sheet_entry = create_input_row(root, "Sheet:", "104")
        self.start_idx_entry = create_input_row(root, "No (起始序号):", "1")
        
        # 预设完美的交叉匹配公式！
        self.template_entry = create_input_row(root, "項目名 模板:", "行番号({行全})_({列})")
        self.csv_id_template_entry = create_input_row(root, "CSV ID 模板:", "33{行2位}{列2位}") 
        
        self.type_entry = create_input_row(root, "タイプ種别:", "テキスト")
        self.max_char_entry = create_input_row(root, "最大文字数:", "6")
        self.align_entry = create_input_row(root, "文字配置:", "指定無し")
        self.font_entry = create_input_row(root, "フォント:", "MS Pゴシック/ 9")

        self.filter_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="🛡️ 自动过滤无效交叉点 (空格子自动拦截，只保留能查到坐标的行)", variable=self.filter_var, font=("MS Gothic", 9, "bold"), fg="#D35400").pack(anchor="w", padx=15, pady=5)

        self.btn_start = tk.Button(root, text="🚀 完美行列解耦，一键生成", bg="#28A745", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=10, ipady=5)
        
        self.log_text = scrolledtext.ScrolledText(root, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
    def update_templates(self):
        mode = self.mode_var.get()
        if mode == "ROW":
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, "市町村民税{行全}_特定親族特別控除")
            self.csv_id_template_entry.delete(0, tk.END)
            self.csv_id_template_entry.insert(0, "19{行2位}31") 
        elif mode == "COL":
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, "({列})_ラベル")
            self.csv_id_template_entry.delete(0, tk.END)
            self.csv_id_template_entry.insert(0, "1901{列2位}")
        else:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, "行番号({行全})_({列})")
            self.csv_id_template_entry.delete(0, tk.END)
            self.csv_id_template_entry.insert(0, "33{行2位}{列2位}")
            
    def browse_file(self, entry_widget):
        path = filedialog.askopenfilename()
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            
    def start_process(self):
        generate_matrix_pattern_spec(
            self.csv_entry.get().strip(), self.excel_entry.get().strip(), self.sheet_entry.get().strip(),
            self.mode_var.get(), self.start_idx_entry.get().strip(), self.template_entry.get().strip(),
            self.csv_id_template_entry.get().strip(), self.type_entry.get().strip(), self.max_char_entry.get().strip(),
            self.align_entry.get().strip(), self.font_entry.get().strip(), self.filter_var.get(), self.log_text
        )

if __name__ == "__main__":
    app_root = tk.Tk()
    app = MatrixPatternApp(app_root)
    app_root.mainloop()
