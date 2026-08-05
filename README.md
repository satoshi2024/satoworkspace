from __future__ import annotations

import os
import traceback
from contextlib import contextmanager
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

import pythoncom
import win32com.client


APP_TITLE = "Excel 指定文字变色工具（直接修改原文件）"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}

# Excel Range.Characters 的 COM 编号
DISPID_CHARACTERS = 0x25B


def rgb_to_excel_color(
    rgb: tuple[int, int, int],
) -> int:
    """
    将普通 RGB 转换成 Excel Font.Color 使用的整数。
    """
    red, green, blue = rgb

    return (
        red
        + green * 256
        + blue * 65536
    )


def parse_ranges(
    text: str,
) -> list[str]:
    """
    支持以下输入：

    BG7
    BG7:BG28
    BG7,BJ7:BJ28
    BG7，BJ7:BJ28
    """
    text = (
        text
        .replace("，", ",")
        .replace("、", ",")
        .replace("；", ",")
        .replace(";", ",")
    )

    ranges = [
        item.strip()
        for item in text.split(",")
        if item.strip()
    ]

    if not ranges:
        raise ValueError(
            "请输入单元格或范围，"
            "例如 BG7 或 BG7:BG28。"
        )

    return ranges


def find_positions(
    source: str,
    keyword: str,
    case_sensitive: bool,
) -> list[int]:
    """
    查找关键词在单元格文字中的所有起始位置。

    返回的位置从 0 开始。
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
        index = search_source.find(
            search_keyword,
            start,
        )

        if index < 0:
            break

        positions.append(index)

        start = (
            index
            + len(search_keyword)
        )

    return positions


def get_anchor_cell(
    cell,
):
    """
    如果是合并单元格，只处理合并区域左上角单元格。
    """
    try:
        if bool(cell.MergeCells):
            return cell.MergeArea.Cells.Item(
                1,
                1,
            )

    except Exception:
        pass

    return cell


def get_cell_text(
    cell,
) -> str | None:
    """
    读取单元格中的文字。
    """
    for attribute_name in (
        "Value2",
        "Value",
        "Text",
    ):
        try:
            value = getattr(
                cell,
                attribute_name,
            )

        except Exception:
            continue

        if isinstance(value, str):
            return value

    return None


def get_cell_address(
    cell,
) -> str:
    """
    获取单元格地址。
    """
    try:
        return str(
            cell.Address
        )

    except Exception:
        return "(未知地址)"


def is_formula_cell(
    cell,
) -> bool:
    """
    判断是否为公式单元格。

    公式单元格的计算结果不能稳定保存局部字体格式，
    因此程序会跳过。
    """
    try:
        return bool(
            cell.HasFormula
        )

    except Exception:
        return False


def get_characters(
    cell,
    start_1_based: int,
    length: int,
):
    """
    取得 Excel Characters 对象。

    优先使用 GetCharacters。

    某些 Excel 或 pywin32 环境会出现
    “成员未找到”的错误，此时使用底层 COM 调用。
    """
    first_error: Exception | None = None

    try:
        return cell.GetCharacters(
            start_1_based,
            length,
        )

    except Exception as error:
        first_error = error

    try:
        raw_characters = (
            cell._oleobj_.InvokeTypes(
                DISPID_CHARACTERS,
                0,
                pythoncom.DISPATCH_PROPERTYGET,
                (
                    pythoncom.VT_DISPATCH,
                    0,
                ),
                (
                    (
                        pythoncom.VT_VARIANT,
                        17,
                    ),
                    (
                        pythoncom.VT_VARIANT,
                        17,
                    ),
                ),
                start_1_based,
                length,
            )
        )

        return win32com.client.Dispatch(
            raw_characters
        )

    except Exception as error:
        raise RuntimeError(
            "无法取得 Excel Characters 对象。"
            f"GetCharacters 错误：{first_error}；"
            f"底层 COM 错误：{error}"
        ) from error


def set_partial_font_color(
    cell,
    start_zero_based: int,
    length: int,
    color: int,
) -> None:
    """
    设置单元格中指定部分文字的字体颜色。

    Python 的字符位置从 0 开始。
    Excel 的字符位置从 1 开始。
    """
    characters = get_characters(
        cell,
        start_zero_based + 1,
        length,
    )

    characters.Font.Color = int(
        color
    )


@contextmanager
def open_workbook(
    path: Path,
    read_only: bool,
):
    """
    使用独立的 Excel 进程打开工作簿。
    """
    excel = None
    workbook = None

    pythoncom.CoInitialize()

    try:
        excel = (
            win32com.client.DispatchEx(
                "Excel.Application"
            )
        )

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(
            os.path.abspath(
                str(path)
            ),
            UpdateLinks=0,
            ReadOnly=read_only,
        )

        if (
            not read_only
            and bool(workbook.ReadOnly)
        ):
            raise PermissionError(
                "文件以只读方式打开。"
                "请关闭正在打开的 Excel 文件，"
                "并确认该文件不是只读文件。"
            )

        yield workbook

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


class App:

    def __init__(
        self,
        root: tk.Tk,
    ) -> None:
        self.root = root

        self.root.title(
            APP_TITLE
        )

        self.root.geometry(
            "760x610"
        )

        self.root.minsize(
            680,
            550,
        )

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()

        self.range_text = tk.StringVar(
            value="BG7:BG28"
        )

        self.keyword = tk.StringVar()

        self.case_sensitive = tk.BooleanVar(
            value=True
        )

        self.match_mode = tk.StringVar(
            value="全部出现位置"
        )

        self.status = tk.StringVar(
            value="请选择 Excel 文件。"
        )

        self.rgb = (
            255,
            0,
            0,
        )

        self.hex_color = "#FF0000"

        self.build_ui()
        self.update_color_preview()

    def build_ui(
        self,
    ) -> None:
        main = ttk.Frame(
            self.root,
            padding=14,
        )

        main.pack(
            fill="both",
            expand=True,
        )

        # ========================================
        # 文件
        # ========================================

        file_box = ttk.LabelFrame(
            main,
            text="1. 文件",
            padding=10,
        )

        file_box.pack(
            fill="x",
        )

        ttk.Label(
            file_box,
            text="Excel 文件",
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ttk.Entry(
            file_box,
            textvariable=self.file_path,
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=8,
        )

        ttk.Button(
            file_box,
            text="选择文件",
            command=self.choose_file,
        ).grid(
            row=0,
            column=2,
        )

        ttk.Label(
            file_box,
            text=(
                "会直接保存到原文件。"
                "请先备份文件，并关闭正在打开的 Excel。"
            ),
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="w",
            padx=8,
            pady=(6, 0),
        )

        file_box.columnconfigure(
            1,
            weight=1,
        )

        # ========================================
        # 指定位置和文字
        # ========================================

        target_box = ttk.LabelFrame(
            main,
            text="2. 指定位置和文字",
            padding=10,
        )

        target_box.pack(
            fill="x",
            pady=(10, 0),
        )

        ttk.Label(
            target_box,
            text="工作表",
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        self.sheet_combo = ttk.Combobox(
            target_box,
            textvariable=self.sheet_name,
            state="readonly",
        )

        self.sheet_combo.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=8,
            pady=4,
        )

        ttk.Button(
            target_box,
            text="读取工作表",
            command=self.load_sheets,
        ).grid(
            row=0,
            column=2,
        )

        ttk.Label(
            target_box,
            text="单元格/范围",
        ).grid(
            row=1,
            column=0,
            sticky="w",
        )

        ttk.Entry(
            target_box,
            textvariable=self.range_text,
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(8, 0),
            pady=4,
        )

        ttk.Label(
            target_box,
            text="例：BG7、BG7:BG28、BG7,BJ7:BJ28",
        ).grid(
            row=2,
            column=1,
            columnspan=2,
            sticky="w",
            padx=8,
        )

        ttk.Label(
            target_box,
            text="指定文字",
        ).grid(
            row=3,
            column=0,
            sticky="w",
        )

        ttk.Entry(
            target_box,
            textvariable=self.keyword,
        ).grid(
            row=3,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(8, 0),
            pady=8,
        )

        ttk.Checkbutton(
            target_box,
            text="区分英文字母大小写",
            variable=self.case_sensitive,
        ).grid(
            row=4,
            column=1,
            sticky="w",
            padx=8,
        )

        ttk.Button(
            target_box,
            text="先检查匹配",
            command=self.preview,
        ).grid(
            row=5,
            column=2,
            sticky="e",
            pady=(8, 0),
        )

        target_box.columnconfigure(
            1,
            weight=1,
        )

        # ========================================
        # 颜色和匹配方式
        # ========================================

        style_box = ttk.LabelFrame(
            main,
            text="3. 颜色和匹配方式",
            padding=10,
        )

        style_box.pack(
            fill="x",
            pady=(10, 0),
        )

        ttk.Label(
            style_box,
            text="字体颜色",
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ttk.Button(
            style_box,
            text="选择颜色",
            command=self.choose_color,
        ).grid(
            row=0,
            column=1,
            padx=8,
        )

        self.color_preview = tk.Label(
            style_box,
            text="示例文字",
            width=12,
            relief="solid",
            borderwidth=1,
        )

        self.color_preview.grid(
            row=0,
            column=2,
            padx=8,
        )

        self.color_label = ttk.Label(
            style_box,
            text=self.hex_color,
        )

        self.color_label.grid(
            row=0,
            column=3,
            sticky="w",
        )

        ttk.Label(
            style_box,
            text="匹配方式",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(8, 0),
        )

        ttk.Combobox(
            style_box,
            textvariable=self.match_mode,
            values=(
                "全部出现位置",
                "只修改第一次",
            ),
            state="readonly",
            width=18,
        ).grid(
            row=1,
            column=1,
            sticky="w",
            padx=8,
            pady=(8, 0),
        )

        # ========================================
        # 操作按钮
        # ========================================

        action_box = ttk.Frame(
            main
        )

        action_box.pack(
            fill="x",
            pady=(12, 0),
        )

        self.run_button = ttk.Button(
            action_box,
            text="直接修改原文件",
            command=self.process,
        )

        self.run_button.pack(
            side="right",
        )

        ttk.Button(
            action_box,
            text="清空日志",
            command=self.clear_log,
        ).pack(
            side="right",
            padx=(0, 8),
        )

        # ========================================
        # 日志
        # ========================================

        log_box = ttk.LabelFrame(
            main,
            text="处理日志",
            padding=8,
        )

        log_box.pack(
            fill="both",
            expand=True,
            pady=(10, 0),
        )

        self.log = tk.Text(
            log_box,
            height=13,
            wrap="word",
            state="disabled",
        )

        scroll = ttk.Scrollbar(
            log_box,
            command=self.log.yview,
        )

        self.log.configure(
            yscrollcommand=scroll.set
        )

        self.log.pack(
            side="left",
            fill="both",
            expand=True,
        )

        scroll.pack(
            side="right",
            fill="y",
        )

        ttk.Label(
            main,
            textvariable=self.status,
            anchor="w",
        ).pack(
            fill="x",
            pady=(6, 0),
        )

    def choose_file(
        self,
    ) -> None:
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

        self.file_path.set(
            path
        )

        self.load_sheets()

    def choose_color(
        self,
    ) -> None:
        rgb, hex_value = (
            colorchooser.askcolor(
                color=self.hex_color,
                title="选择字体颜色",
            )
        )

        if (
            rgb is None
            or hex_value is None
        ):
            return

        self.rgb = tuple(
            round(value)
            for value in rgb
        )

        self.hex_color = (
            hex_value.upper()
        )

        self.update_color_preview()

    def update_color_preview(
        self,
    ) -> None:
        red, green, blue = self.rgb

        brightness = (
            red * 299
            + green * 587
            + blue * 114
        ) / 1000

        foreground = (
            "#000000"
            if brightness > 150
            else "#FFFFFF"
        )

        self.color_preview.configure(
            background=self.hex_color,
            foreground=foreground,
        )

        self.color_label.configure(
            text=self.hex_color
        )

    def set_busy(
        self,
        busy: bool,
        text: str | None = None,
    ) -> None:
        self.run_button.configure(
            state=(
                "disabled"
                if busy
                else "normal"
            )
        )

        self.root.configure(
            cursor=(
                "wait"
                if busy
                else ""
            )
        )

        if text:
            self.status.set(
                text
            )

        self.root.update_idletasks()

    def load_sheets(
        self,
    ) -> None:
        path = Path(
            self.file_path.get().strip()
        )

        if not path.exists():
            messagebox.showwarning(
                APP_TITLE,
                "请先选择有效的 Excel 文件。",
            )
            return

        self.set_busy(
            True,
            "正在读取工作表……",
        )

        try:
            with open_workbook(
                path,
                read_only=True,
            ) as workbook:

                names = [
                    sheet.Name
                    for sheet
                    in workbook.Worksheets
                ]

            self.sheet_combo[
                "values"
            ] = names

            if names:
                self.sheet_name.set(
                    names[0]
                )

            self.status.set(
                f"读取完成，共 "
                f"{len(names)} 个工作表。"
            )

        except Exception as error:
            self.show_error(
                f"读取工作表失败：{error}"
            )

        finally:
            self.set_busy(
                False
            )

    def get_settings(
        self,
    ) -> dict:
        path = Path(
            self.file_path.get().strip()
        )

        if not path.exists():
            raise ValueError(
                "请选择有效的 Excel 文件。"
            )

        if (
            path.suffix.lower()
            not in SUPPORTED_EXTENSIONS
        ):
            raise ValueError(
                "只支持 xlsx、xlsm、xls、xlsb 文件。"
            )

        sheet_name = (
            self.sheet_name.get().strip()
        )

        if not sheet_name:
            raise ValueError(
                "请选择工作表。"
            )

        keyword = self.keyword.get()

        if not keyword:
            raise ValueError(
                "请输入需要变色的文字。"
            )

        return {
            "path": path,
            "sheet": sheet_name,
            "ranges": parse_ranges(
                self.range_text.get()
            ),
            "keyword": keyword,
            "case_sensitive": (
                self.case_sensitive.get()
            ),
            "match_mode": (
                self.match_mode.get()
            ),
            "rgb": self.rgb,
        }

    def find_matches(
        self,
        worksheet,
        settings: dict,
    ):
        checked = 0
        matches = []
        visited: set[str] = set()

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

            for raw_cell in target_range.Cells:
                cell = get_anchor_cell(
                    raw_cell
                )

                address = get_cell_address(
                    cell
                )

                if address in visited:
                    continue

                visited.add(
                    address
                )

                checked += 1

                text = get_cell_text(
                    cell
                )

                if not isinstance(
                    text,
                    str,
                ):
                    continue

                positions = find_positions(
                    text,
                    settings["keyword"],
                    settings["case_sensitive"],
                )

                if (
                    settings["match_mode"]
                    == "只修改第一次"
                ):
                    positions = positions[:1]

                if positions:
                    matches.append(
                        (
                            cell,
                            positions,
                        )
                    )

        return (
            matches,
            checked,
        )

    def preview(
        self,
    ) -> None:
        try:
            settings = self.get_settings()

        except ValueError as error:
            messagebox.showwarning(
                APP_TITLE,
                str(error),
            )
            return

        self.clear_log()

        self.append_log(
            "开始检查匹配，不会修改文件。\n\n"
        )

        self.set_busy(
            True,
            "正在检查匹配……",
        )

        try:
            with open_workbook(
                settings["path"],
                read_only=True,
            ) as workbook:

                worksheet = (
                    workbook.Worksheets(
                        settings["sheet"]
                    )
                )

                matches, checked = (
                    self.find_matches(
                        worksheet,
                        settings,
                    )
                )

                total = 0

                for cell, positions in matches:
                    total += len(
                        positions
                    )

                    formula_mark = (
                        "【公式】"
                        if is_formula_cell(cell)
                        else ""
                    )

                    self.append_log(
                        f"[匹配] "
                        f"{get_cell_address(cell)} "
                        f"{formula_mark}："
                        f"{len(positions)} 处\n"
                    )

            self.append_log(
                "\n"
                f"检查单元格：{checked}\n"
                f"匹配单元格：{len(matches)}\n"
                f"匹配位置：{total}\n"
            )

            self.status.set(
                f"检查完成，找到 "
                f"{total} 个匹配位置。"
            )

            if total == 0:
                messagebox.showwarning(
                    APP_TITLE,
                    "没有找到指定文字。",
                )

        except Exception as error:
            self.show_error(
                f"检查失败：{error}"
            )

        finally:
            self.set_busy(
                False
            )

    def process(
        self,
    ) -> None:
        try:
            settings = self.get_settings()

        except ValueError as error:
            messagebox.showwarning(
                APP_TITLE,
                str(error),
            )
            return

        path: Path = settings["path"]

        confirmed = messagebox.askyesno(
            APP_TITLE,
            "将直接修改并保存原文件：\n\n"
            f"{path}\n\n"
            "不会生成新文件。建议事先备份。\n"
            "确定继续吗？",
        )

        if not confirmed:
            return

        self.clear_log()

        self.append_log(
            "开始处理。\n"
            "请先关闭正在打开的 Excel 文件。\n\n"
            f"原文件：\n{path}\n\n"
        )

        self.set_busy(
            True,
            "正在处理……",
        )

        try:
            with open_workbook(
                path,
                read_only=False,
            ) as workbook:

                worksheet = (
                    workbook.Worksheets(
                        settings["sheet"]
                    )
                )

                matches, checked = (
                    self.find_matches(
                        worksheet,
                        settings,
                    )
                )

                excel_color = (
                    rgb_to_excel_color(
                        settings["rgb"]
                    )
                )

                total = sum(
                    len(positions)
                    for _, positions
                    in matches
                )

                success = 0
                failed = 0
                formula_skipped = 0

                for cell, positions in matches:
                    address = get_cell_address(
                        cell
                    )

                    if is_formula_cell(cell):
                        formula_skipped += len(
                            positions
                        )

                        self.append_log(
                            f"[跳过] {address}："
                            f"公式单元格。\n"
                        )

                        continue

                    changed_here = 0

                    for position in positions:
                        try:
                            set_partial_font_color(
                                cell,
                                position,
                                len(
                                    settings["keyword"]
                                ),
                                excel_color,
                            )

                            success += 1
                            changed_here += 1

                        except Exception as error:
                            failed += 1

                            self.append_log(
                                f"[失败] {address}："
                                f"开始位置 {position + 1}，"
                                f"长度 "
                                f"{len(settings['keyword'])}；"
                                f"{error}\n"
                            )

                    if changed_here:
                        self.append_log(
                            f"[完成] {address}："
                            f"成功修改 "
                            f"{changed_here} 处。\n"
                        )

                # 直接保存原文件
                workbook.Save()

            self.append_log(
                "\n处理完成。\n"
                f"检查单元格：{checked}\n"
                f"匹配位置：{total}\n"
                f"成功变色：{success}\n"
                f"公式跳过：{formula_skipped}\n"
                f"失败：{failed}\n"
                f"已保存原文件：{path}\n"
            )

            self.status.set(
                f"处理完成，成功变色 "
                f"{success} 处。"
            )

            messagebox.showinfo(
                APP_TITLE,
                "处理完成。\n\n"
                f"成功变色：{success}\n"
                f"公式跳过：{formula_skipped}\n"
                f"失败：{failed}\n\n"
                f"已保存到原文件：\n"
                f"{path}",
            )

        except PermissionError as error:
            self.show_error(
                str(error)
            )

        except Exception as error:
            self.show_error(
                f"处理失败：{error}"
            )

        finally:
            self.set_busy(
                False
            )

    def clear_log(
        self,
    ) -> None:
        self.log.configure(
            state="normal"
        )

        self.log.delete(
            "1.0",
            "end",
        )

        self.log.configure(
            state="disabled"
        )

    def append_log(
        self,
        text: str,
    ) -> None:
        self.log.configure(
            state="normal"
        )

        self.log.insert(
            "end",
            text,
        )

        self.log.see(
            "end"
        )

        self.log.configure(
            state="disabled"
        )

        self.root.update_idletasks()

    def show_error(
        self,
        message: str,
    ) -> None:
        self.status.set(
            "操作失败。"
        )

        self.append_log(
            f"\n{message}\n\n"
            f"{traceback.format_exc()}\n"
        )

        messagebox.showerror(
            APP_TITLE,
            message,
        )


def main() -> None:
    root = tk.Tk()

    try:
        style = ttk.Style(
            root
        )

        if "vista" in style.theme_names():
            style.theme_use(
                "vista"
            )

    except Exception:
        pass

    App(root)

    root.mainloop()


if __name__ == "__main__":
    main()