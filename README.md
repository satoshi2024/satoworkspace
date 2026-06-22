import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def generate_pattern_spec(excel_path, sheet_name, scan_mode, start_idx, item_template, item_type, max_chars, align, font_size, edit_method, log_widget):
    try:
        log_widget.insert(tk.END, f"正在加载 Excel 文件 [{os.path.basename(excel_path)}]...\n")
        log_widget.update()
        
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{sheet_name}]")
            return
            
        ws = wb[sheet_name]
        extracted_ids = []

        # ==========================================
        # 模式 A：扫描【行番号】(013, 023, 033...)
        # ==========================================
        if scan_mode == "ROW":
            log_widget.insert(tk.END, f"正在启动全景雷达，扫描行番号...\n")
            id_col_start = None
            id_row_start = None
            
            for r in range(1, 100):
                for c in range(1, 100):
                    val = ws.cell(row=r, column=c).value
                    if val:
                        clean_val = re.sub(r'\s+', '', str(val))
                        if "行番号" in clean_val:
                            id_col_start = c
                            id_row_start = r
                            break
                        elif clean_val == "行": 
                            v_below = ws.cell(row=r+1, column=c).value
                            if v_below and "番" in re.sub(r'\s+', '', str(v_below)):
                                id_col_start = c
                                id_row_start = r + 2
                                break
                if id_col_start: break
                
            if id_col_start:
                for r in range(id_row_start + 1, ws.max_row + 1):
                    logic_str = ""
                    for offset in range(8): # 横向拼接最多8个格子
                        v = ws.cell(row=r, column=id_col_start + offset).value
                        if v is not None:
                            s = str(v).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                            if s.isdigit():
                                logic_str += s
                    if len(logic_str) >= 2:
                        if logic_str not in extracted_ids:
                            extracted_ids.append(logic_str)
                            
            log_widget.insert(tk.END, f"➔ 成功提取到 {len(extracted_ids)} 个行番号 ID。\n\n")

        # ==========================================
        # 模式 B：扫描【列表头】(28, 29, 30...)
        # ==========================================
        elif scan_mode == "COL":
            log_widget.insert(tk.END, f"正在扫描前 50 行，提取括号列表头...\n")
            for c in range(1, ws.max_column + 10):
                for r in range(1, 50):
                    val = ws.cell(row=r, column=c).value
                    if val is not None:
                        v_str = str(val).translate(str.maketrans('０１２３４５６７８９', '0123456789')).replace(" ", "").replace("　", "").strip()
                        m = re.search(r'[\(（](\d+)[\)）]', v_str)
                        if m:
                            num = m.group(1)
                            if num not in extracted_ids:
                                extracted_ids.append(num)
                            break
            # 对列号进行纯数字排序
            extracted_ids = sorted(extracted_ids, key=lambda x: int(x))
            log_widget.insert(tk.END, f"➔ 成功提取到 {len(extracted_ids)} 个列表头 ID。\n\n")

        if not extracted_ids:
            messagebox.showerror("错误", "未能提取到任何有效的 ID，请检查扫描模式或 Excel 内容。")
            return

        # ==========================================
        # 生成 CSV 数据
        # ==========================================
        log_widget.insert(tk.END, f"正在拼装 패턴表 (Pattern Table) 数据...\n")
        output_rows = []
        
        # 写入完美的 10 列标准表头
        output_rows.append(["No", "項目名", "選択項目No", "タイプ種別", "最大文字数", "文字配置", "領域外の対処", "文字フォント/サイズ", "編集箇所", "編集方法"])
        
        current_idx = int(start_idx)
        
        for logic_id in extracted_ids:
            # 使用 {ID} 动态替换项目名，例如: "市町村民税{ID}_特定親族特別控除" -> "市町村民税013_特定親族特別控除"
            item_name = item_template.replace("{ID}", logic_id)
            
            row_data = [
                str(current_idx),          # No
                item_name,                 # 項目名
                "",                        # 選択項目No
                item_type,                 # タイプ種別
                max_chars,                 # 最大文字数
                align,                     # 文字配置
                "出力しない",               # 領域外の対処 (固定默认)
                font_size,                 # 文字フォント/サイズ
                "-",                       # 編集箇所 (固定默认)
                edit_method                # 編集方法
            ]
            output_rows.append(row_data)
            
            log_widget.insert(tk.END, f" ✅ 组装: {current_idx} | {item_name}\n")
            current_idx += 1

        # ==========================================
        # 导出 CSV
        # ==========================================
        output_csv_path = os.path.join(os.path.dirname(excel_path), f"Pattern_Spec_Sheet{sheet_name}_{scan_mode}.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(output_rows)
            
        log_widget.insert(tk.END, f"\n🎉 【处理成功】\n共生成 {len(extracted_ids)} 行式样数据。\n文件已保存: {os.path.basename(output_csv_path)}\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"Pattern 式样书生成完毕！\n\n共生成 {len(extracted_ids)} 行数据。\n请用 Excel 打开生成的 CSV，直接全选粘贴到你的模板中即可。")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

# ==========================================
# GUI 界面
# ==========================================
class PatternSpecGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("测试式样书自动生成器 (パターン表 专属)")
        self.root.geometry("680x720")
        
        def create_input_row(parent, label_text, default_val=""):
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15, pady=4)
            tk.Label(frame, text=label_text, width=22, anchor="w", font=("MS Gothic", 9, "bold")).pack(side="left")
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True)
            entry.insert(0, default_val)
            return entry

        # 1. 文件选择
        tk.Label(root, text="1. 请选择最新版的 Excel 文件:", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        excel_frame = tk.Frame(root)
        excel_frame.pack(fill="x", padx=15, pady=2)
        self.excel_entry = tk.Entry(excel_frame, font=("Calibri", 10))
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(excel_frame, text="浏览...", width=10, command=self.browse_file).pack(side="right", padx=5)

        # 2. 扫描模式
        tk.Label(root, text="2. 扫描来源模式 (决定 {ID} 提取什么):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        mode_frame = tk.Frame(root)
        mode_frame.pack(fill="x", padx=15)
        self.mode_var = tk.StringVar(value="ROW")
        tk.Radiobutton(mode_frame, text="按【行番号】生成 (提取 013, 023...)", variable=self.mode_var, value="ROW", font=("MS Gothic", 9)).pack(side="left", padx=10)
        tk.Radiobutton(mode_frame, text="按【列表头】生成 (提取 28, 29...)", variable=self.mode_var, value="COL", font=("MS Gothic", 9)).pack(side="left", padx=10)

        # 3. 参数配置
        tk.Label(root, text="3. 表格参数配置 (使用 {ID} 作为动态变量):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        self.sheet_entry = create_input_row(root, "目标 Sheet (如 69):", "69")
        self.start_idx_entry = create_input_row(root, "No (起始序号如 17):", "17")
        self.template_entry = create_input_row(root, "項目名 模板 (含 {ID}):", "市町村民税{ID}_特定親族特別控除")
        self.type_entry = create_input_row(root, "タイプ種別:", "テキスト")
        self.max_char_entry = create_input_row(root, "最大文字数:", "13")
        self.align_entry = create_input_row(root, "文字配置:", "右配置")
        self.font_entry = create_input_row(root, "文字フォント/サイズ:", "MS Pゴシック/ 9")
        self.edit_method_entry = create_input_row(root, "編集方法 (可留空):", "#REF!")

        # 4. 按钮与日志
        self.btn_start = tk.Button(root, text="🚀 一键生成 パターン表 (CSV)", bg="#28A745", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        self.log_text = scrolledtext.ScrolledText(root, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)
            
    def start_process(self):
        excel_p = self.excel_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        mode = self.mode_var.get()
        start_idx = self.start_idx_entry.get().strip()
        template = self.template_entry.get().strip()
        i_type = self.type_entry.get().strip()
        max_c = self.max_char_entry.get().strip()
        align = self.align_entry.get().strip()
        font_s = self.font_entry.get().strip()
        edit_m = self.edit_method_entry.get().strip()
        
        if not excel_p or "{ID}" not in template:
            messagebox.showwarning("提示", "请选择 Excel 文件，并且【項目名 模板】中必须包含 {ID} 占位符！")
            return
            
        self.log_text.delete(1.0, tk.END)
        generate_pattern_spec(excel_p, sheet_n, mode, start_idx, template, i_type, max_c, align, font_s, edit_m, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = PatternSpecGeneratorApp(app_root)
    app_root.mainloop()
