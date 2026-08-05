先安装依赖：

py -m pip install pywin32

将下面源码保存为：

excel_text_color_gui.py

然后运行：

py excel_text_color_gui.py
from __future__ import annotations
import os
import shutil
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk
import pythoncom
import win32com.client
APP_TITLE = "Excel 指定文字变色工具"
SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
}
def rgb_to_excel_color(rgb: tuple[int, int, int]) -> int:
    """
    将普通 RGB 转换成 Excel Font.Color 使用的整数。
    """
    red, green, blue = rgb
    return red + green * 256 + blue * 65536
def find_all_positions(
    source: str,
    keyword: str,
    case_sensitive: bool = True,
) -> list[int]:
    """
    查找指定文字在字符串中的所有位置。
    返回值使用 Python 的 0 开始索引。
    """
    if not keyword:
        return []
    if case_sensitive:
        search_source = source
        search_keyword = keyword
    else:
        search_source = source.casefold()
        search_keyword = keyword.casefold()
    positions: list[int] = []
    start = 0
    while True:
        index = search_source.find(search_keyword, start)
        if index == -1:
            break
        positions.append(index)
        start = index + len(search_keyword)
    return positions
def parse_ranges(value: str) -> list[str]:
    """
    支持：
    BG3
    BG3:BG24
    BG3,BG8,BG10:BG20
    """
    normalized = (
        value.replace("，", ",")
        .replace("、", ",")
        .replace("；", ",")
        .replace(";", ",")
    )
    ranges = [
        item.strip()
        for item in normalized.split(",")
        if item.strip()
    ]
    if not ranges:
        raise ValueError(
            "请输入单元格或范围，例如 BG3 或 BG3:BG24。"
        )
    return ranges
class ExcelTextColorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x650")
        self.root.minsize(680, 580)
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.cell_range = tk.StringVar(value="BG3:BG24")
        self.keyword = tk.StringVar()
        self.case_sensitive = tk.BooleanVar(value=True)
        self.match_mode = tk.StringVar(value="全部出现位置")
        self.occurrence_number = tk.IntVar(value=1)
        self.status = tk.StringVar(
            value="请选择 Excel 文件。"
        )
        self.selected_rgb = (255, 0, 0)
        self.selected_hex = "#FF0000"
        self.build_ui()
        self.update_color_preview()
        self.update_occurrence_state()
    def build_ui(self) -> None:
        main_frame = ttk.Frame(
            self.root,
            padding=16,
        )
        main_frame.pack(
            fill="both",
            expand=True,
        )
        # =========================
        # 文件设置
        # =========================
        file_frame = ttk.LabelFrame(
            main_frame,
            text="1. 文件设置",
            padding=12,
        )
        file_frame.pack(
            fill="x",
        )
        ttk.Label(
            file_frame,
            text="Excel 文件",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        ttk.Entry(
            file_frame,
            textvariable=self.input_file,
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            pady=6,
        )
        ttk.Button(
            file_frame,
            text="选择文件",
            command=self.choose_input_file,
        ).grid(
            row=0,
            column=2,
            padx=(8, 0),
            pady=6,
        )
        ttk.Label(
            file_frame,
            text="输出文件",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        ttk.Entry(
            file_frame,
            textvariable=self.output_file,
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            pady=6,
        )
        ttk.Button(
            file_frame,
            text="选择位置",
            command=self.choose_output_file,
        ).grid(
            row=1,
            column=2,
            padx=(8, 0),
            pady=6,
        )
        file_frame.columnconfigure(
            1,
            weight=1,
        )
        # =========================
        # 指定位置
        # =========================
        target_frame = ttk.LabelFrame(
            main_frame,
            text="2. 指定单元格和文字",
            padding=12,
        )
        target_frame.pack(
            fill="x",
            pady=(12, 0),
        )
        ttk.Label(
            target_frame,
            text="工作表",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        self.sheet_combo = ttk.Combobox(
            target_frame,
            textvariable=self.sheet_name,
            state="readonly",
        )
        self.sheet_combo.grid(
            row=0,
            column=1,
            sticky="ew",
            pady=6,
        )
        ttk.Button(
            target_frame,
            text="读取工作表",
            command=self.load_sheet_names,
        ).grid(
            row=0,
            column=2,
            padx=(8, 0),
            pady=6,
        )
        ttk.Label(
            target_frame,
            text="单元格/范围",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        ttk.Entry(
            target_frame,
            textvariable=self.cell_range,
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            pady=6,
        )
        ttk.Label(
            target_frame,
            text="例：BG3、BG3:BG24、BG3,BG8:BG15",
        ).grid(
            row=2,
            column=1,
            columnspan=2,
            sticky="w",
        )
        ttk.Label(
            target_frame,
            text="指定文字",
        ).grid(
            row=3,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=(12, 6),
        )
        ttk.Entry(
            target_frame,
            textvariable=self.keyword,
        ).grid(
            row=3,
            column=1,
            columnspan=2,
            sticky="ew",
            pady=(12, 6),
        )
        ttk.Checkbutton(
            target_frame,
            text="区分英文字母大小写",
            variable=self.case_sensitive,
        ).grid(
            row=4,
            column=1,
            columnspan=2,
            sticky="w",
            pady=6,
        )
        target_frame.columnconfigure(
            1,
            weight=1,
        )
        # =========================
        # 颜色及匹配方式
        # =========================
        color_frame = ttk.LabelFrame(
            main_frame,
            text="3. 颜色和匹配方式",
            padding=12,
        )
        color_frame.pack(
            fill="x",
            pady=(12, 0),
        )
        ttk.Label(
            color_frame,
            text="字体颜色",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        ttk.Button(
            color_frame,
            text="选择颜色",
            command=self.choose_color,
        ).grid(
            row=0,
            column=1,
            sticky="w",
            pady=6,
        )
        self.color_preview = tk.Label(
            color_frame,
            text="示例文字",
            width=12,
            relief="solid",
            borderwidth=1,
        )
        self.color_preview.grid(
            row=0,
            column=2,
            padx=(12, 8),
            pady=6,
        )
        self.color_value_label = ttk.Label(
            color_frame,
            text=self.selected_hex,
        )
        self.color_value_label.grid(
            row=0,
            column=3,
            sticky="w",
            pady=6,
        )
        ttk.Label(
            color_frame,
            text="匹配方式",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=6,
        )
        self.match_combo = ttk.Combobox(
            color_frame,
            textvariable=self.match_mode,
            state="readonly",
            values=(
                "全部出现位置",
                "只修改第一次",
                "修改指定次数",
            ),
            width=18,
        )
        self.match_combo.grid(
            row=1,
            column=1,
            sticky="w",
            pady=6,
        )
        self.match_combo.bind(
            "<<ComboboxSelected>>",
            lambda event: self.update_occurrence_state(),
        )
        ttk.Label(
            color_frame,
            text="第",
        ).grid(
            row=1,
            column=2,
            sticky="e",
            padx=(12, 4),
            pady=6,
        )
        self.occurrence_spinbox = ttk.Spinbox(
            color_frame,
            from_=1,
            to=999,
            textvariable=self.occurrence_number,
            width=8,
        )
        self.occurrence_spinbox.grid(
            row=1,
            column=3,
            sticky="w",
            pady=6,
        )
        ttk.Label(
            color_frame,
            text="次出现",
        ).grid(
            row=1,
            column=4,
            sticky="w",
            padx=(4, 0),
            pady=6,
        )
        # =========================
        # 操作按钮
        # =========================
        button_frame = ttk.Frame(
            main_frame,
        )
        button_frame.pack(
            fill="x",
            pady=(14, 0),
        )
        self.run_button = ttk.Button(
            button_frame,
            text="开始处理",
            command=self.start_processing,
        )
        self.run_button.pack(
            side="right",
        )
        ttk.Button(
            button_frame,
            text="清空日志",
            command=self.clear_log,
        ).pack(
            side="right",
            padx=(0, 8),
        )
        # =========================
        # 日志
        # =========================
        log_frame = ttk.LabelFrame(
            main_frame,
            text="处理日志",
            padding=8,
        )
        log_frame.pack(
            fill="both",
            expand=True,
            pady=(12, 0),
        )
        self.log_text = tk.Text(
            log_frame,
            height=12,
            wrap="word",
            state="disabled",
        )
        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )
        self.log_text.configure(
            yscrollcommand=scrollbar.set,
        )
        self.log_text.pack(
            side="left",
            fill="both",
            expand=True,
        )
        scrollbar.pack(
            side="right",
            fill="y",
        )
        ttk.Label(
            main_frame,
            textvariable=self.status,
            anchor="w",
        ).pack(
            fill="x",
            pady=(8, 0),
        )
    def choose_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[
                (
                    "Excel 文件",
                    "*.xlsx *.xlsm *.xls *.xlsb",
                ),
                (
                    "所有文件",
                    "*.*",
                ),
            ],
        )
        if not path:
            return
        self.input_file.set(path)
        source = Path(path)
        default_output = source.with_name(
            f"{source.stem}_文字变色后{source.suffix}"
        )
        self.output_file.set(
            str(default_output)
        )
        self.load_sheet_names()
    def choose_output_file(self) -> None:
        input_path = self.input_file.get().strip()
        if input_path:
            initial_directory = str(
                Path(input_path).parent
            )
            initial_file = (
                Path(input_path).stem
                + "_文字变色后"
                + Path(input_path).suffix
            )
        else:
            initial_directory = os.getcwd()
            initial_file = "文字变色后.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            initialdir=initial_directory,
            initialfile=initial_file,
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel 工作簿",
                    "*.xlsx",
                ),
                (
                    "启用宏的 Excel 工作簿",
                    "*.xlsm",
                ),
                (
                    "Excel 97-2003",
                    "*.xls",
                ),
                (
                    "Excel 二进制工作簿",
                    "*.xlsb",
                ),
            ],
        )
        if path:
            self.output_file.set(path)
    def choose_color(self) -> None:
        rgb_value, hex_value = colorchooser.askcolor(
            color=self.selected_hex,
            title="选择字体颜色",
        )
        if rgb_value is None or hex_value is None:
            return
        self.selected_rgb = tuple(
            round(number)
            for number in rgb_value
        )
        self.selected_hex = hex_value.upper()
        self.update_color_preview()
    def update_color_preview(self) -> None:
        red, green, blue = self.selected_rgb
        brightness = (
            red * 299
            + green * 587
            + blue * 114
        ) / 1000
        text_color = (
            "#000000"
            if brightness > 150
            else "#FFFFFF"
        )
        self.color_preview.configure(
            background=self.selected_hex,
            foreground=text_color,
        )
        self.color_value_label.configure(
            text=self.selected_hex,
        )
    def update_occurrence_state(self) -> None:
        if self.match_mode.get() == "修改指定次数":
            self.occurrence_spinbox.configure(
                state="normal"
            )
        else:
            self.occurrence_spinbox.configure(
                state="disabled"
            )
    def load_sheet_names(self) -> None:
        file_path = self.input_file.get().strip()
        if not file_path:
            messagebox.showwarning(
                APP_TITLE,
                "请先选择 Excel 文件。",
            )
            return
        if not Path(file_path).exists():
            messagebox.showwarning(
                APP_TITLE,
                "Excel 文件不存在。",
            )
            return
        self.set_running(True)
        self.status.set(
            "正在读取工作表……"
        )
        thread = threading.Thread(
            target=self.load_sheet_names_worker,
            args=(file_path,),
            daemon=True,
        )
        thread.start()
    def load_sheet_names_worker(
        self,
        file_path: str,
    ) -> None:
        excel = None
        workbook = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx(
                "Excel.Application"
            )
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(
                os.path.abspath(file_path),
                UpdateLinks=0,
                ReadOnly=True,
            )
            sheet_names = [
                sheet.Name
                for sheet in workbook.Worksheets
            ]
            self.root.after(
                0,
                lambda: self.set_sheet_names(
                    sheet_names
                ),
            )
        except Exception as error:
            details = traceback.format_exc()
            self.root.after(
                0,
                lambda: self.show_error(
                    f"读取工作表失败：{error}",
                    details,
                ),
            )
        finally:
            try:
                if workbook is not None:
                    workbook.Close(
                        SaveChanges=False
                    )
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
            self.root.after(
                0,
                lambda: self.set_running(False),
            )
    def set_sheet_names(
        self,
        sheet_names: list[str],
    ) -> None:
        self.sheet_combo["values"] = sheet_names
        if sheet_names:
            self.sheet_name.set(
                sheet_names[0]
            )
            self.status.set(
                f"读取完成，共 {len(sheet_names)} 个工作表。"
            )
        else:
            self.sheet_name.set("")
            self.status.set(
                "没有读取到工作表。"
            )
    def validate_settings(self) -> dict:
        input_file = Path(
            self.input_file.get().strip()
        )
        output_file = Path(
            self.output_file.get().strip()
        )
        if not input_file.exists():
            raise ValueError(
                "请选择有效的 Excel 文件。"
            )
        if input_file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                "只支持 xlsx、xlsm、xls、xlsb 文件。"
            )
        if not self.output_file.get().strip():
            raise ValueError(
                "请选择输出文件。"
            )
        if input_file.resolve() == output_file.resolve():
            raise ValueError(
                "输出文件不能和原文件相同。"
            )
        if (
            input_file.suffix.lower()
            != output_file.suffix.lower()
        ):
            raise ValueError(
                "输出文件扩展名必须与原文件相同。"
            )
        sheet_name = self.sheet_name.get().strip()
        if not sheet_name:
            raise ValueError(
                "请选择工作表。"
            )
        ranges = parse_ranges(
            self.cell_range.get()
        )
        keyword = self.keyword.get()
        if not keyword:
            raise ValueError(
                "请输入需要变色的文字。"
            )
        occurrence = int(
            self.occurrence_number.get()
        )
        if occurrence < 1:
            raise ValueError(
                "出现次数必须大于等于 1。"
            )
        return {
            "input_file": input_file,
            "output_file": output_file,
            "sheet_name": sheet_name,
            "ranges": ranges,
            "keyword": keyword,
            "case_sensitive": self.case_sensitive.get(),
            "match_mode": self.match_mode.get(),
            "occurrence": occurrence,
            "color": self.selected_rgb,
        }
    def start_processing(self) -> None:
        try:
            settings = self.validate_settings()
        except ValueError as error:
            messagebox.showwarning(
                APP_TITLE,
                str(error),
            )
            return
        output_file: Path = settings["output_file"]
        if output_file.exists():
            confirm = messagebox.askyesno(
                APP_TITLE,
                f"输出文件已经存在：\n\n"
                f"{output_file}\n\n"
                f"是否覆盖？",
            )
            if not confirm:
                return
        self.clear_log()
        self.append_log(
            "开始处理。\n"
            "处理时请不要打开目标 Excel 文件。\n\n"
        )
        self.status.set(
            "正在处理……"
        )
        self.set_running(True)
        thread = threading.Thread(
            target=self.process_worker,
            args=(settings,),
            daemon=True,
        )
        thread.start()
    def process_worker(
        self,
        settings: dict,
    ) -> None:
        excel = None
        workbook = None
        try:
            pythoncom.CoInitialize()
            input_file: Path = settings["input_file"]
            output_file: Path = settings["output_file"]
            output_file.parent.mkdir(
                parents=True,
                exist_ok=True,
            )
            shutil.copy2(
                input_file,
                output_file,
            )
            self.append_log_threadsafe(
                f"已复制原文件：\n"
                f"{output_file}\n\n"
            )
            excel = win32com.client.DispatchEx(
                "Excel.Application"
            )
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(
                os.path.abspath(output_file),
                UpdateLinks=0,
                ReadOnly=False,
            )
            try:
                worksheet = workbook.Worksheets(
                    settings["sheet_name"]
                )
            except Exception as error:
                raise ValueError(
                    f"找不到工作表："
                    f"{settings['sheet_name']}"
                ) from error
            excel_color = rgb_to_excel_color(
                settings["color"]
            )
            keyword = settings["keyword"]
            total_cells = 0
            total_changed = 0
            visited_cells: set[str] = set()
            for range_address in settings["ranges"]:
                try:
                    target_range = worksheet.Range(
                        range_address
                    )
                except Exception as error:
                    raise ValueError(
                        f"无效的单元格或范围："
                        f"{range_address}"
                    ) from error
                for original_cell in target_range.Cells:
                    cell = original_cell
                    try:
                        if bool(original_cell.MergeCells):
                            cell = (
                                original_cell
                                .MergeArea
                                .Cells(1, 1)
                            )
                    except Exception:
                        cell = original_cell
                    cell_key = (
                        f"{worksheet.Name}!"
                        f"{cell.Address}"
                    )
                    if cell_key in visited_cells:
                        continue
                    visited_cells.add(cell_key)
                    total_cells += 1
                    value = cell.Value
                    if not isinstance(value, str):
                        continue
                    positions = find_all_positions(
                        source=value,
                        keyword=keyword,
                        case_sensitive=settings[
                            "case_sensitive"
                        ],
                    )
                    match_mode = settings["match_mode"]
                    if match_mode == "只修改第一次":
                        positions = positions[:1]
                    elif match_mode == "修改指定次数":
                        target_index = (
                            settings["occurrence"] - 1
                        )
                        if 0 <= target_index < len(positions):
                            positions = [
                                positions[target_index]
                            ]
                        else:
                            positions = []
                    changed_in_cell = 0
                    for position in positions:
                        try:
                            # Excel Characters 的位置从 1 开始
                            characters = cell.Characters(
                                Start=position + 1,
                                Length=len(keyword),
                            )
                            characters.Font.Color = (
                                excel_color
                            )
                            changed_in_cell += 1
                        except Exception as error:
                            self.append_log_threadsafe(
                                f"[失败] "
                                f"{worksheet.Name}!"
                                f"{cell.Address}："
                                f"{error}\n"
                            )
                    if changed_in_cell > 0:
                        total_changed += changed_in_cell
                        self.append_log_threadsafe(
                            f"[完成] "
                            f"{worksheet.Name}!"
                            f"{cell.Address}："
                            f"“{keyword}” "
                            f"修改 {changed_in_cell} 处\n"
                        )
            workbook.Save()
            self.append_log_threadsafe(
                "\n处理完成。\n"
                f"检查单元格：{total_cells}\n"
                f"变色位置：{total_changed}\n"
                f"保存位置：{output_file}\n"
            )
            self.root.after(
                0,
                lambda: self.finish_success(
                    total_changed,
                    output_file,
                ),
            )
        except Exception as error:
            details = traceback.format_exc()
            self.root.after(
                0,
                lambda: self.show_error(
                    f"处理失败：{error}",
                    details,
                ),
            )
        finally:
            try:
                if workbook is not None:
                    workbook.Close(
                        SaveChanges=False
                    )
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.EnableEvents = True
                    excel.ScreenUpdating = True
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
            self.root.after(
                0,
                lambda: self.set_running(False),
            )
    def finish_success(
        self,
        changed_count: int,
        output_file: Path,
    ) -> None:
        self.status.set(
            f"处理完成，共修改 "
            f"{changed_count} 处文字颜色。"
        )
        messagebox.showinfo(
            APP_TITLE,
            f"处理完成。\n\n"
            f"共修改 {changed_count} 处文字颜色。\n\n"
            f"保存位置：\n"
            f"{output_file}",
        )
    def show_error(
        self,
        message: str,
        details: str,
    ) -> None:
        self.status.set(
            "处理失败。"
        )
        self.append_log(
            f"\n{message}\n\n"
            f"{details}\n"
        )
        messagebox.showerror(
            APP_TITLE,
            message,
        )
    def set_running(
        self,
        running: bool,
    ) -> None:
        if running:
            self.run_button.configure(
                state="disabled"
            )
        else:
            self.run_button.configure(
                state="normal"
            )
    def clear_log(self) -> None:
        self.log_text.configure(
            state="normal"
        )
        self.log_text.delete(
            "1.0",
            "end",
        )
        self.log_text.configure(
            state="disabled"
        )
    def append_log(
        self,
        text: str,
    ) -> None:
        self.log_text.configure(
            state="normal"
        )
        self.log_text.insert(
            "end",
            text,
        )
        self.log_text.see(
            "end"
        )
        self.log_text.configure(
            state="disabled"
        )
    def append_log_threadsafe(
        self,
        text: str,
    ) -> None:
        self.root.after(
            0,
            lambda: self.append_log(text),
        )
def main() -> None:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    ExcelTextColorApp(root)
    root.mainloop()
if __name__ == "__main__":
    main()

这个程序需要在 Windows 并且安装了桌面版 Microsoft Excel 的电脑上运行。默认会复制一个新文件，不直接修改原文件。