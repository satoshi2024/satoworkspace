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
                    c_id, c_sheet, c_cell = row[0].strip(), row[1].strip(), row[2].strip()
                    csv_dict[c_id] = {"sheet": c_sheet, "cell": c_cell}
                    
        log_widget.insert(tk.END, f"➔ 成功加载 {len(csv_dict)} 条坐标映射数据。\n\n")

        log_widget.insert(tk.END, f"【2/3】正在启动全景雷达扫描 Excel...\n")
        log_widget.update()
        
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{sheet_name}]")
            return
        ws = wb[sheet_name]
        
        extracted_items = [] 

        if scan_mode == "ROW":
            id_col_start, id_row_start = None, None
            for r in range(1, 100):
                for c in range(1, 100):
                    val = ws.cell(row=r, column=c).value
                    if val:
                        clean_val = re.sub(r'\s+', '', str(val))
                        if "行番号" in clean_val:
                            id_col_start, id_row_start = c, r
                            break
                        elif clean_val == "行": 
                            v_below = ws.cell(row=r+1, column=c).value
                            if v_below and "番" in re.sub(r'\s+', '', str(v_below)):
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
                            extracted_items.append({
                                "raw": logic_str,       # 对应 {行全}
                                "short": logic_str[:2]  # 对应 {行2位}
                            })

        elif scan_mode == "COL":
            for c in range(1, ws.max_column + 10):
                for r in range(1, 50):
                    val = ws.cell(row=r, column=c).value
                    if val is not None:
                        v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                        m = re.search(r'[\(（](\d+)[\)）]', v_str)
                        if m:
                            num = m.group(1)
                            if num not in [x["raw"] for x in extracted_items]:
                                extracted_items.append({
                                    "raw": num,         # 对应 {列}
                                    "short": num 
                                })
                            break
            extracted_items = sorted(extracted_items, key=lambda x: int(x["raw"]))

        if not extracted_items:
            messagebox.showerror("错误", "未能提取到任何有效数据，请检查扫描模式或 Excel 内容。")
            return
            
        log_widget.insert(tk.END, f"➔ 成功提取到 {len(extracted_items)} 个动态变量。\n\n")

        log_widget.insert(tk.END, f"【3/3】正在现取坐标，拼装 Pattern 表...\n")
        output_rows = []
        output_rows.append(["No", "項目名", "選択項目No", "タイプ種別", "最大文字数", "文字配置", "領域外の対処", "文字フォント/サイズ", "編集箇所", "編集方法"])
        
        current_idx = int(start_idx)
        found_count = 0
        missing_count = 0
        
        for item in extracted_items:
            raw_val = item["raw"]
            short_val = item["short"]
            
            # 使用直白的中文占位符进行替换
            item_name = item_template.replace("{行全}", raw_val).replace("{行2位}", short_val).replace("{列}", raw_val)
            search_id = csv_id_template.replace("{行全}", raw_val).replace("{行2位}", short_val).replace("{列}", raw_val)
            
            # 现取坐标
            if search_id in csv_dict:
                c_sheet = csv_dict[search_id]["sheet"]
                c_cell = csv_dict[search_id]["cell"]
                edit_method = f"表行列: S{search_id}\nシート番号: {c_sheet}\n座標: {c_cell}"
                found_count += 1
                log_widget.insert(tk.END, f" 🎯 命中: [CSV ID: {search_id}] ➔ 坐标: {c_cell}\n")
            else:
                edit_method = f"表行列: S{search_id}\nシート番号: 未知\n座標: 未知"
                missing_count += 1
                log_widget.insert(tk.END, f" ⚠️ 未找到: [CSV ID: {search_id}] (字典中不存在)\n")
            
            row_data = [
                str(current_idx), item_name, "", item_type, max_chars, align, 
                "出力しない", font_size, "-", edit_method
            ]
            output_rows.append(row_data)
            current_idx += 1

        output_csv_path = os.path.join(os.path.dirname(excel_path), f"Pattern_Spec_Sheet{sheet_name}_{scan_mode}.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(output_rows)
            
        log_widget.insert(tk.END, f"\n🎉 【生成完毕】\n成功查表: {found_count} 个 | 失败: {missing_count} 个\n文件已保存: {os.path.basename(output_csv_path)}\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"式样书生成完毕！\n\n成功查询到 {found_count} 个坐标。")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

class DynamicPatternApp:
    def __init__(self, root):
        self.root = root
        self.root.title("测试式样书生成器 (V5 直白占位符版)")
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
        tk.Label(csv_frame, text="映射 CSV (查坐标用):", width=22, anchor="w").pack(side="left")
        self.csv_entry = tk.Entry(csv_frame, font=("Calibri", 10))
        self.csv_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(csv_frame, text="浏览...", width=10, command=lambda: self.browse_file(self.csv_entry, [("CSV", "*.csv")])).pack(side="right", padx=5)
        
        excel_frame = tk.Frame(root)
        excel_frame.pack(fill="x", padx=15, pady=2)
        tk.Label(excel_frame, text="最新 Excel (扫表头用):", width=22, anchor="w").pack(side="left")
        self.excel_entry = tk.Entry(excel_frame, font=("Calibri", 10))
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(excel_frame, text="浏览...", width=10, command=lambda: self.browse_file(self.excel_entry, [("Excel", "*.xlsm *.xlsx")])).pack(side="right", padx=5)

        tk.Label(root, text="2. 扫描来源模式:", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        mode_frame = tk.Frame(root)
        mode_frame.pack(fill="x", padx=15)
        self.mode_var = tk.StringVar(value="ROW") 
        
        tk.Radiobutton(mode_frame, text="按【行番号】(扫 013, 023)", variable=self.mode_var, value="ROW", font=("MS Gothic", 9), command=self.update_templates).pack(side="left", padx=10)
        tk.Radiobutton(mode_frame, text="按【列表头】(扫 25, 26)", variable=self.mode_var, value="COL", font=("MS Gothic", 9), command=self.update_templates).pack(side="left", padx=10)

        self.lbl_template_hint = tk.Label(root, text="3. 模板配置 (可用中文占位符: {行全}, {行2位}, {列}):", font=("MS Gothic", 9, "bold"), fg="blue")
        self.lbl_template_hint.pack(anchor="w", padx=15, pady=(10, 2))
        
        self.sheet_entry = create_input_row(root, "扫描目标 Sheet:", "69")
        self.start_idx_entry = create_input_row(root, "No (起始序号):", "17")
        self.template_entry = create_input_row(root, "項目名 模板:", "市町村民税{行全}_特定親族特別控除")
        self.csv_id_template_entry = create_input_row(root, "CSV ID 匹配模板:", "12{行2位}25") # 这里就是绝杀！
        self.type_entry = create_input_row(root, "タイプ種別:", "テキスト")
        self.max_char_entry = create_input_row(root, "最大文字数:", "13")
        self.align_entry = create_input_row(root, "文字配置:", "右配置")
        self.font_entry = create_input_row(root, "文字フォント/サイズ:", "MS Pゴシック/ 9")

        self.btn_start = tk.Button(root, text="🚀 现取坐标，一键生成", bg="#28A745", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        self.log_text = scrolledtext.ScrolledText(root, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
    def update_templates(self):
        mode = self.mode_var.get()
        if mode == "ROW":
            self.lbl_template_hint.config(text="3. 模板配置 (可用占位符: {行全}=013, {行2位}=01):")
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, "市町村民税{行全}_特定親族特別控除")
            self.csv_id_template_entry.delete(0, tk.END)
            self.csv_id_template_entry.insert(0, "12{行2位}25") # 完美匹配 12 + 行番 + 25
        else:
            self.lbl_template_hint.config(text="3. 模板配置 (可用占位符: {列}=28):")
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, "({列})_ラベル")
            self.csv_id_template_entry.delete(0, tk.END)
            self.csv_id_template_entry.insert(0, "1201{列}")
            
    def browse_file(self, entry_widget, f_types):
        path = filedialog.askopenfilename(filetypes=f_types)
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            
    def start_process(self):
        csv_p = self.csv_entry.get().strip()
        excel_p = self.excel_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        mode = self.mode_var.get()
        start_idx = self.start_idx_entry.get().strip()
        template = self.template_entry.get().strip()
        csv_id_template = self.csv_id_template_entry.get().strip()
        i_type = self.type_entry.get().strip()
        max_c = self.max_char_entry.get().strip()
        align = self.align_entry.get().strip()
        font_s = self.font_entry.get().strip()
        
        if not csv_p or not excel_p:
            messagebox.showwarning("提示", "请选择文件！")
            return
            
        self.log_text.delete(1.0, tk.END)
        generate_dynamic_pattern_spec(csv_p, excel_p, sheet_n, mode, start_idx, template, csv_id_template, i_type, max_c, align, font_s, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = DynamicPatternApp(app_root)
    app_root.mainloop()
