from __future__ import annotations

import os
import shutil
import traceback
from contextlib import contextmanager
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

import pythoncom
import win32com.client


APP_TITLE = "Excel 指定文字变色工具（修正版）"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}

# Excel Range.Characters 的 COM 编号
DISPID_CHARACTERS = 0x25B


def rgb_to_excel_color(
    rgb: tuple[int, int, int],
) -> int:
    """
    将普通 RGB 转换为 Excel Font.Color 使用的整数。
    """
    red, green, blue = rgb

    return int(
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

    result = [
        item.strip()
        for item in text.split(",")
        if item.strip()
    ]

    if not result:
        raise ValueError(
            "请输入单元格或范围，"
            "例如 BG7 或 BG7:BG28。"
        )

    return result


def find_positions(
    source: str,
    keyword: str,
    case_sensitive: bool,
) -> list[int]:
    """
    查找关键词出现的所有位置。

    返回 Python 的 0 开始索引。
    """
    if not case_sensitive:
        source = source.casefold()
        keyword = keyword.casefold()

    positions: list[int] = []
    start = 0

    while True:
        index = source.find(
            keyword,
            start,
        )

        if index < 0:
            break

        positions.append(index)

        # 不处理重叠匹配
        start = index + len(keyword)

    return positions


def get_anchor_cell(
    cell,
):
    """
    如果是合并单元格，返回合并区域左上角单元格。
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
    从 Excel 单元格读取文字。
    """
    for name in (
        "Value2",
        "Value",
        "Text",
    ):
        try:
            value = getattr(
                cell,
                name,
            )
        except Exception:
            continue

        if isinstance(value, str):
            return value

    return None


def get_cell_address(
    cell,
) -> str:
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

    Excel 公式结果通常不能保存局部富文本格式。
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
    获取 Excel Characters 对象。

    第一种方式：
        使用 pywin32 早绑定生成的 GetCharacters。

    第二种方式：
        如果出现“成员未找到”，直接通过 COM 调用
        Excel Range.Characters 属性。

    start_1_based：
        Excel 使用的 1 开始索引。
    """
    first_error: Exception | None = None

    # ==========================================
    # 方法一：GetCharacters
    # ==========================================

    try:
        return cell.GetCharacters(
            start_1_based,
            length,
        )

    except Exception as error:
        first_error = error

    # ==========================================
    # 方法二：直接调用 Excel COM Characters
    # ==========================================

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
            "无法取得 Excel Characters 对象。\n"
            f"GetCharacters 错误：{first_error}\n"
            f"COM Characters 错误：{error}"
        ) from error


def set_partial_font_color(
    cell,
    start_zero_based: int,
    length: int,
    color: int,
) -> None:
    """
    设置单元格内指定部分文字的字体颜色。

    start_zero_based：
        Python 的 0 开始索引。

    Excel Characters：
        使用 1 开始索引，所以需要加 1。
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
    启动独立 Excel 进程并打开工作簿。
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

        root.title(
            APP_TITLE
        )

        root.geometry(
            "760x650"
        )

        root.minsize(
            680,
            580,
        )

        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.sheet_name = tk.StringVar()

        self.cell_range = tk.StringVar(
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
        # 1. 文件设置
        # ========================================

        file_box = ttk.LabelFrame(
            main,
            text="1. 文件设置",
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
            pady=5,
        )

        ttk.Entry(
            file_box,
            textvariable=self.input_file,
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=8,
            pady=5,
        )

        ttk.Button(
            file_box,
            text="选择文件",
            command=self.choose_input,
        ).grid(
            row=0,
            column=2,
            pady=5,
        )

        ttk.Label(
            file_box,
            text="输出文件",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )

        ttk.Entry(
            file_box,
            textvariable=self.output_file,
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=8,
            pady=5,
        )

        ttk.Button(
            file_box,
            text="选择位置",
            command=self.choose_output,
        ).grid(
            row=1,
            column=2,
            pady=5,
        )

        file_box.columnconfigure(
            1,
            weight=1,
        )

        # ========================================
        # 2. 指定单元格和文字
        # ========================================

        target_box = ttk.LabelFrame(
            main,
            text="2. 指定单元格和文字",
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
            pady=5,
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
            pady=5,
        )

        ttk.Button(
            target_box,
            text="读取工作表",
            command=self.load_sheets,
        ).grid(
            row=0,
            column=2,
            pady=5,
        )

        ttk.Label(
            target_box,
            text="单元格/范围",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )

        ttk.Entry(
            target_box,
            textvariable=self.cell_range,
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(8, 0),
            pady=5,
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
            pady=8,
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
        # 3. 颜色和匹配方式
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
            pady=5,
        )

        ttk.Button(
            style_box,
            text="选择颜色",
            command=self.choose_color,
        ).grid(
            row=0,
            column=1,
            sticky="w",
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
            pady=5,
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
        )

        # ========================================
        # 按钮
        # ========================================

        button_box = ttk.Frame(
            main
        )

        button_box.pack(
            fill="x",
            pady=(12, 0),
        )

        self.run_button = ttk.Button(
            button_box,
            text="开始处理",
            command=self.process,
        )

        self.run_button.pack(
            side="right",
        )

        ttk.Button(
            button_box,
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

    def choose_input(
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

        self.input_file.set(
            path
        )

        source = Path(path)

        output = source.with_name(
            f"{source.stem}_文字变色后"
            f"{source.suffix}"
        )

        self.output_file.set(
            str(output)
        )

        self.load_sheets()

    def choose_output(
        self,
    ) -> None:
        input_value = (
            self.input_file.get().strip()
        )

        if input_value:
            source = Path(
                input_value
            )
        else:
            source = (
                Path.cwd()
                / "文字变色后.xlsx"
            )

        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            initialdir=str(
                source.parent
            ),
            initialfile=(
                f"{source.stem}_文字变色后"
                f"{source.suffix}"
            ),
            defaultextension=(
                source.suffix
                or ".xlsx"
            ),
            filetypes=[
                (
                    "Excel 工作簿",
                    "*.xlsx",
                ),
                (
                    "启用宏的工作簿",
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
            self.output_file.set(
                path
            )

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
        red, green, blue = (
            self.rgb
        )

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
        status: str | None = None,
    ) -> None:
        if busy:
            self.run_button.configure(
                state="disabled"
            )

            self.root.configure(
                cursor="wait"
            )

        else:
            self.run_button.configure(
                state="normal"
            )

            self.root.configure(
                cursor=""
            )

        if status:
            self.status.set(
                status
            )

        self.root.update_idletasks()

    def load_sheets(
        self,
    ) -> None:
        path = Path(
            self.input_file.get().strip()
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
        require_output: bool,
    ) -> dict:
        input_file = Path(
            self.input_file.get().strip()
        )

        if not input_file.exists():
            raise ValueError(
                "请选择有效的 Excel 文件。"
            )

        if (
            input_file.suffix.lower()
            not in SUPPORTED_EXTENSIONS
        ):
            raise ValueError(
                "只支持 xlsx、xlsm、xls、xlsb 文件。"
            )

        sheet = (
            self.sheet_name.get().strip()
        )

        if not sheet:
            raise ValueError(
                "请选择工作表。"
            )

        keyword = self.keyword.get()

        if not keyword:
            raise ValueError(
                "请输入需要变色的文字。"
            )

        config = {
            "input": input_file,
            "sheet": sheet,
            "ranges": parse_ranges(
                self.cell_range.get()
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

        if require_output:
            output_text = (
                self.output_file.get().strip()
            )

            if not output_text:
                raise ValueError(
                    "请选择输出文件。"
                )

            output = Path(
                output_text
            )

            if (
                output.suffix.lower()
                != input_file.suffix.lower()
            ):
                raise ValueError(
                    "输出文件扩展名必须与原文件相同。"
                )

            if (
                output.resolve()
                == input_file.resolve()
            ):
                raise ValueError(
                    "输出文件不能与原文件相同。"
                )

            config["output"] = output

        return config

    def find_matches(
        self,
        worksheet,
        config: dict,
    ):
        checked = 0
        matches = []

        visited: set[str] = set()

        for address in config["ranges"]:
            try:
                target_range = (
                    worksheet.Range(
                        address
                    )
                )

            except Exception as error:
                raise ValueError(
                    f"无效的单元格或范围："
                    f"{address}"
                ) from error

            for raw_cell in target_range.Cells:
                cell = get_anchor_cell(
                    raw_cell
                )

                cell_address = (
                    get_cell_address(cell)
                )

                if cell_address in visited:
                    continue

                visited.add(
                    cell_address
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
                    source=text,
                    keyword=config["keyword"],
                    case_sensitive=config[
                        "case_sensitive"
                    ],
                )

                if (
                    config["match_mode"]
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
            config = self.get_settings(
                require_output=False
            )

        except ValueError as error:
            messagebox.showwarning(
                APP_TITLE,
                str(error),
            )
            return

        self.clear_log()

        self.append_log(
            "开始检查匹配，"
            "不会修改文件。\n\n"
        )

        self.set_busy(
            True,
            "正在检查匹配……",
        )

        try:
            with open_workbook(
                config["input"],
                read_only=True,
            ) as workbook:

                worksheet = (
                    workbook.Worksheets(
                        config["sheet"]
                    )
                )

                matches, checked = (
                    self.find_matches(
                        worksheet,
                        config,
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
            config = self.get_settings(
                require_output=True
            )

        except ValueError as error:
            messagebox.showwarning(
                APP_TITLE,
                str(error),
            )
            return

        output: Path = config[
            "output"
        ]

        if output.exists():
            overwrite = (
                messagebox.askyesno(
                    APP_TITLE,
                    "输出文件已经存在：\n\n"
                    f"{output}\n\n"
                    "是否覆盖？",
                )
            )

            if not overwrite:
                return

        self.clear_log()

        self.append_log(
            "开始处理。\n"
            "请先关闭原文件和输出文件。\n\n"
        )

        self.set_busy(
            True,
            "正在处理……",
        )

        try:
            output.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            shutil.copy2(
                config["input"],
                output,
            )

            self.append_log(
                "已生成副本：\n"
                f"{output}\n\n"
            )

            with open_workbook(
                output,
                read_only=False,
            ) as workbook:

                worksheet = (
                    workbook.Worksheets(
                        config["sheet"]
                    )
                )

                matches, checked = (
                    self.find_matches(
                        worksheet,
                        config,
                    )
                )

                excel_color = (
                    rgb_to_excel_color(
                        config["rgb"]
                    )
                )

                success = 0
                failed = 0
                formula_skipped = 0

                total = sum(
                    len(positions)
                    for _, positions
                    in matches
                )

                for cell, positions in matches:
                    address = (
                        get_cell_address(cell)
                    )

                    if is_formula_cell(cell):
                        formula_skipped += len(
                            positions
                        )

                        self.append_log(
                            f"[跳过] {address}："
                            "公式单元格。\n"
                        )

                        continue

                    changed_here = 0

                    for position in positions:
                        try:
                            set_partial_font_color(
                                cell=cell,
                                start_zero_based=position,
                                length=len(
                                    config["keyword"]
                                ),
                                color=excel_color,
                            )

                            success += 1
                            changed_here += 1

                        except Exception as error:
                            failed += 1

                            self.append_log(
                                f"[失败] {address}："
                                f"开始位置 {position + 1}，"
                                f"长度 "
                                f"{len(config['keyword'])}；"
                                f"{error}\n"
                            )

                    if changed_here:
                        self.append_log(
                            f"[完成] {address}："
                            f"成功修改 "
                            f"{changed_here} 处。\n"
                        )

                workbook.Save()

            self.append_log(
                "\n处理完成。\n"
                f"检查单元格：{checked}\n"
                f"匹配位置：{total}\n"
                f"成功变色：{success}\n"
                f"公式跳过：{formula_skipped}\n"
                f"失败：{failed}\n"
                f"保存位置：{output}\n"
            )

            self.status.set(
                f"处理完成，成功变色 "
                f"{success} 处。"
            )

            if (
                success > 0
                and failed == 0
            ):
                messagebox.showinfo(
                    APP_TITLE,
                    "处理完成。\n\n"
                    f"成功变色：{success} 处\n\n"
                    "保存位置：\n"
                    f"{output}",
                )

            else:
                messagebox.showwarning(
                    APP_TITLE,
                    "处理结束。\n\n"
                    f"成功：{success}\n"
                    f"公式跳过：{formula_skipped}\n"
                    f"失败：{failed}\n\n"
                    "请查看处理日志。",
                )

        except PermissionError:
            self.show_error(
                "文件无法写入。"
                "请关闭正在打开的 Excel 文件后再试。"
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