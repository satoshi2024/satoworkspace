import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def generate_dynamic_pattern_spec(csv_path, excel_path, sheet_name, scan_mode, start_idx, item_template, csv_id_template, item_type, max_chars, align, font_size, log_widget):
    try:
        log_widget.insert(tk.END, f"【1/3】正在读取 CSV 映射库...\n")
        log_widget.update()
        
        csv_dict = {}
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 3:
                    csv_dict[row[0].strip()] = {"sheet": row[1].strip(), "cell": row[2].strip()}

        log_widget.insert(tk.END, f"【2/3】正在启动全景雷达扫描 Excel...\n")
        log_widget.update()
        
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("错误", f"找不到 Sheet: [{sheet_name}]")
            return
        ws = wb[sheet_name]
        
        extracted_items = [] 

        if scan_mode == "ROW":
            id_col_start, id_row_start = None, None
            for r in range(1, 100):
                for c in range(1, 100):
                    val = str(ws.cell(r, c).value or "").replace(" ", "").replace("　", "").replace("\n", "")
                    # 绝杀：必须是纯粹的表头，不能是长篇说明文！
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
                        if logic_str not in [x["raw"] for x in extracted_items]:
                            extracted_items.append({"raw": logic_str, "short": logic_str[:2]})

        elif scan_mode == "COL":
            for c in range(1, ws.max_column + 20):
                for r in range(1, 100):
                    val = ws.cell(row=r, column=c).value
                    if val is not None:
                        # 兼容 (31)\n計 这种带换行的恶劣排版
                        lines = str(val).split('\n')
                        v_str = lines[0].translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                        m = re.search(r'^[\(（](\d+)[\)）]', v_str)
                        if m:
                            num = m.group(1)
                            if num not in [x["raw"] for x in extracted_items]:
                                extracted_items.append({"raw": num, "short": num})
            extracted_items = sorted(extracted_items, key=lambda x: int(x["raw"]))

        if not extracted_items:
            messagebox.showerror("错误", "未能提取到数据！")
            return

        log_widget.insert(tk.END, f"【3/3】正在拼装 Pattern 表...\n")
        output_rows = [["No", "項目名", "選択項目No", "タイプ種別", "最大文字数", "文字配置", "領域外の対処", "文字フォント/サイズ", "編集箇所", "編集方法"]]
        
        current_idx = int(start_idx)
        for item in extracted_items:
            raw_val, short_val = item["raw"], item["short"]
            item_name = item_template.replace("{行全}", raw_val).replace("{行2位}", short_val).replace("{列}", raw_val)
            search_id = csv_id_template.replace("{行全}", raw_val).replace("{行2位}", short_val).replace("{列}", raw_val)
            
            if search_id in csv_dict:
                c_sheet, c_cell = csv_dict[search_id]["sheet"], csv_dict[search_id]["cell"]
                edit_method = f"表行列: S{search_id}\nシート番号: {c_sheet}\n座標: {c_cell}"
                log_widget.insert(tk.END, f" 🎯 命中: [CSV ID: {search_id}] ➔ 坐标: {c_cell}\n")
            else:
                edit_method = f"表行列: S{search_id}\nシート番号: 未知\n座標: 未知"
                log_widget.insert(tk.END, f" ⚠️ 未找到: [CSV ID: {search_id}]\n")
            
            output_rows.append([str(current_idx), item_name, "", item_type, max_chars, align, "出力しない", font_size, "-", edit_method])
            current_idx += 1

        output_csv_path = os.path.join(os.path.dirname(excel_path), f"Pattern_Spec_Sheet{sheet_name}_{scan_mode}.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerows(output_rows)
            
        log_widget.insert(tk.END, f"\n✅ 生成完毕！文件已保存。\n")
        
    except Exception as e:
        messagebox.showerror("系统错误", str(e))

class DynamicPatternApp:
    def __init__(self, root):
        self.root = root
        self.root.title("式样书生成器 (V7 穿透折叠排版版)")
        self.root.geometry("700x780")
        
        def create_input_row(parent, label_text, default_val=""):
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15, pady=4)
            tk.Label(frame, text=label_text, width=22, anchor="w", font=("MS Gothic", 9, "bold")).pack(side="left")
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True)
            entry.insert(0, default_val)
            return entry

        tk.Label(root, text="1. 文件选择:", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
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
        self.mode_var = tk.StringVar(value="COL") 
        tk.Radiobutton(mode_frame, text="【行番号】", variable=self.mode_var, value="ROW").pack(side="left", padx=10)
        tk.Radiobutton(mode_frame, text="【列表头】", variable=self.mode_var, value="COL").pack(side="left", padx=10)

        tk.Label(root, text="3. 模板配置 ({行全}=013, {行2位}=01, {列}=31):", font=("MS Gothic", 9, "bold"), fg="blue").pack(anchor="w", padx=15, pady=(10, 2))
        self.sheet_entry = create_input_row(root, "Sheet:", "88")
        self.start_idx_entry = create_input_row(root, "No:", "17")
        self.template_entry = create_input_row(root, "項目名 模板:", "({列})_ラベル")
        self.csv_id_template_entry = create_input_row(root, "CSV ID 模板:", "1901{列}") 
        self.type_entry = create_input_row(root, "タイプ種別:", "ラベル")
        self.max_char_entry = create_input_row(root, "最大文字数:", "4")
        self.align_entry = create_input_row(root, "文字配置:", "指定無し")
        self.font_entry = create_input_row(root, "フォント:", "MS 明朝/ 8")

        self.btn_start = tk.Button(root, text="🚀 一键生成", bg="#28A745", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        self.log_text = scrolledtext.ScrolledText(root, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
    def browse_file(self, entry_widget):
        path = filedialog.askopenfilename()
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            
    def start_process(self):
        generate_dynamic_pattern_spec(
            self.csv_entry.get().strip(), self.excel_entry.get().strip(), self.sheet_entry.get().strip(),
            self.mode_var.get(), self.start_idx_entry.get().strip(), self.template_entry.get().strip(),
            self.csv_id_template_entry.get().strip(), self.type_entry.get().strip(), self.max_char_entry.get().strip(),
            self.align_entry.get().strip(), self.font_entry.get().strip(), self.log_text
        )

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicPatternApp(app_root)
    app_root.mainloop()
