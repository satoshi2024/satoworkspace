import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def generate_ut_spec(excel_path, sheet_name, test_prefix, start_idx, biz_prefix, del_item, shift_col_id, shift_item, log_widget):
    try:
        log_widget.insert(tk.END, f"正在启动全景雷达加载 Excel 文件...\n")
        log_widget.update()
        
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("错误", f"在 Excel 中找不到指定的 Sheet: [{sheet_name}]")
            return
            
        ws = wb[sheet_name]

        # ==========================================
        # 1. 暴力提取所有的行番号 (换上最强全景雷达)
        # ==========================================
        id_col_start = None
        id_row_start = None
        
        # 扩大扫描范围，强力抹除换行和空格
        for r in range(1, 100):
            for c in range(1, 100):
                val = ws.cell(row=r, column=c).value
                if val:
                    clean_val = re.sub(r'\s+', '', str(val))
                    if "行番号" in clean_val:
                        id_col_start = c
                        id_row_start = r
                        break
                    # 兼容纵向换行拆分的表头
                    elif clean_val == "行": 
                        v_below = ws.cell(row=r+1, column=c).value
                        if v_below and "番" in re.sub(r'\s+', '', str(v_below)):
                            id_col_start = c
                            id_row_start = r + 2
                            break
            if id_col_start: break
            
        logic_list = []
        if id_col_start:
            # 开启横向吸尘器，扩大范围到 8 列，把碎成渣的数字全拼起来
            for r in range(id_row_start + 1, ws.max_row + 1):
                logic_str = ""
                for offset in range(8):
                    v = ws.cell(row=r, column=id_col_start + offset).value
                    if v is not None:
                        s = str(v).translate(str.maketrans('０１２３４５６７８９', '0123456789')).strip()
                        if s.isdigit():
                            logic_str += s
                            
                if len(logic_str) >= 2:
                    if logic_str not in logic_list:
                        logic_list.append(logic_str)

        if not logic_list:
            messagebox.showerror("错误", "全景雷达未能提取到任何有效的行番号！请检查该 Sheet 是否包含数据。")
            return
            
        log_widget.insert(tk.END, f"➔ 突破排版，成功提取到 {len(logic_list)} 个有效行番号。\n\n")

        # ==========================================
        # 2. 拼装生成 UT 式样书数据
        # ==========================================
        log_widget.insert(tk.END, f"正在拼装 UT 式样书话术...\n")
        output_rows = []
        
        # 写入表头 (方便你核对)
        output_rows.append(["No", "種別", "確認項目・条件", "確認内容", "実施", "結果1", "確認日1", "結果2", "確認日2"])
        
        current_idx = int(start_idx)
        
        for logic_num in logic_list:
            # 生成 No (如 001-011)
            test_id = f"{test_prefix}{str(current_idx).zfill(3)}"
            
            # 生成 確認項目・条件
            condition = f"{biz_prefix}{logic_num}_{del_item}"
            
            # 生成 確認内容
            content = f"{del_item}が表示されなくなり、また、それ以降にずれていた項目については、金額（{shift_col_id}列）が「{shift_item}」として表示されるように修正し、その後の項目の金額も正しく表示されること。"
            
            # 组装这一行数据
            row_data = [test_id, "13", condition, content, "対象", "", "", "", ""]
            output_rows.append(row_data)
            
            log_widget.insert(tk.END, f" ✅ 生成: {test_id} | {condition}\n")
            current_idx += 1

        # ==========================================
        # 3. 导出 CSV
        # ==========================================
        output_csv_path = os.path.join(os.path.dirname(excel_path), f"UT_Spec_Auto_Sheet{sheet_name}.csv")
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(output_rows)
            
        log_widget.insert(tk.END, f"\n🎉 【处理成功】\n共生成了 {len(logic_list)} 条测试用例。\n文件已保存在 Excel 同目录下: {os.path.basename(output_csv_path)}\n")
        log_widget.see(tk.END)
        messagebox.showinfo("成功", f"UT 式样书生成完毕！\n\n共生成 {len(logic_list)} 条测试用例。\n您可以直接用 Excel 打开生成的 CSV 文件，并复制到模板中。")
        
    except Exception as e:
        messagebox.showerror("系统错误", f"发生异常:\n{str(e)}")

# ==========================================
# GUI 界面
# ==========================================
class UTSpecGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("UT 式样书自动生成器 (全景雷达突破版)")
        self.root.geometry("650x650")
        
        def create_input_row(parent, label_text, default_val=""):
            frame = tk.Frame(parent)
            frame.pack(fill="x", padx=15, pady=5)
            tk.Label(frame, text=label_text, width=22, anchor="w", font=("MS Gothic", 9, "bold")).pack(side="left")
            entry = tk.Entry(frame, font=("Calibri", 10))
            entry.pack(side="left", fill="x", expand=True)
            entry.insert(0, default_val)
            return entry

        tk.Label(root, text="1. 请选择最新版的 Excel 文件 (用于读取行番号):", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        excel_frame = tk.Frame(root)
        excel_frame.pack(fill="x", padx=15, pady=5)
        self.excel_entry = tk.Entry(excel_frame, font=("Calibri", 10))
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(excel_frame, text="浏览...", width=10, command=self.browse_file).pack(side="right", padx=5)

        tk.Label(root, text="2. UT 参数配置:", font=("MS Gothic", 9, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        self.sheet_entry = create_input_row(root, "目标 Sheet (如 69):", "69")
        self.test_prefix_entry = create_input_row(root, "Test No 前缀:", "001-")
        self.start_idx_entry = create_input_row(root, "起始序号 (如 11):", "11")
        self.biz_prefix_entry = create_input_row(root, "业务分类 (如 市町村民税):", "市町村民税")
        self.del_item_entry = create_input_row(root, "被废弃的列名:", "定額による特別控除額")
        self.shift_col_entry = create_input_row(root, "递补上来的逻辑列号:", "18")
        self.shift_item_entry = create_input_row(root, "递补上来的列名:", "減免税額")

        self.btn_start = tk.Button(root, text="🚀 一键生成 UT 测试用例", bg="#28A745", fg="white", font=("MS Gothic", 11, "bold"), command=self.start_process)
        self.btn_start.pack(fill="x", padx=15, pady=15, ipady=5)
        
        self.log_text = scrolledtext.ScrolledText(root, height=10, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)
            
    def start_process(self):
        excel_p = self.excel_entry.get().strip()
        sheet_n = self.sheet_entry.get().strip()
        test_pref = self.test_prefix_entry.get().strip()
        start_idx = self.start_idx_entry.get().strip()
        biz_pref = self.biz_prefix_entry.get().strip()
        del_item = self.del_item_entry.get().strip()
        shift_col = self.shift_col_entry.get().strip()
        shift_item = self.shift_item_entry.get().strip()
        
        if not excel_p:
            messagebox.showwarning("提示", "请选择 Excel 文件！")
            return
            
        self.log_text.delete(1.0, tk.END)
        generate_ut_spec(excel_p, sheet_n, test_pref, start_idx, biz_pref, del_item, shift_col, shift_item, self.log_text)

if __name__ == "__main__":
    app_root = tk.Tk()
    app = UTSpecGeneratorApp(app_root)
    app_root.mainloop()
