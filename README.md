from __future__ import annotations

import os
import shutil
import traceback
import unicodedata
from contextlib import contextmanager
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

import pythoncom
import win32com.client


APP_TITLE = "Excel 指定文字变色工具（修正版）"

SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
}

# Excel 常量
XL_VALUES = -4163
XL_PART = 2
XL_BY_ROWS = 1
XL_NEXT = 1

ZERO_WIDTH_CHARS = {
    "\u200b",
    "\u200c",
    "\u200d",
    "\u2060",
    "\ufeff",
}


def rgb_to_excel_color(
    rgb: tuple[int, int, int],
) -> int:
    """
    将普通 RGB 转换为 Excel Font.Color 使用的整数。
    """
    red, green, blue = rgb

    return (
        red
        + green * 256
        + blue * 65536
    )


def parse_ranges(
    value: str,
) -> list[str]:
    """
    支持：

    BG7
    BG7:BG28
    BG7,BJ7:BJ28
    BG7，BJ7:BJ28
    """
    normalized = (
        value
        .replace("，", ",")
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
            "请输入单元格或范围，"
            "例如 BG7 或 BG7:BG28。"
        )

    return ranges


def normalize_with_map(
    text: str,
    *,
    case_sensitive: bool,
    normalize_width: bool,
    ignore_whitespace: bool,
) -> tuple[str, list[int]]:
    """
    将文字规范化，同时记录规范化字符对应的原始位置。

    可以处理：
    1. 全角和半角差异
    2. 英文字母大小写差异
    3. 空格和换行
    4. 零宽字符
    """
    output: list[str] = []
    position_map: list[int] = []

    for original_index, character in enumerate(text):
        if character in ZERO_WIDTH_CHARS:
            continue

        if normalize_width:
            converted = unicodedata.normalize(
                "NFKC",
                character,
            )
        else:
            converted = character

        if not case_sensitive:
            converted = converted.casefold()

        for converted_character in converted:
            if converted_character in ZERO_WIDTH_CHARS:
                continue

            if (
                ignore_whitespace
                and converted_character.isspace()
            ):
                continue

            output.append(converted_character)
            position_map.append(original_index)

    return "".join(output), position_map


def find_original_spans(
    source: str,
    keyword: str,
    *,
    case_sensitive: bool,
    normalize_width: bool,
    ignore_whitespace: bool,
) -> list[tuple[int, int]]:
    """
    查找指定文字。

    返回：
    [
        (原始开始位置, 原始长度),
        ...
    ]

    开始位置使用 Python 的 0 开始索引。
    """
    normalized_source, source_map = normalize_with_map(
        source,
        case_sensitive=case_sensitive,
        normalize_width=normalize_width,
        ignore_whitespace=ignore_whitespace,
    )

    normalized_keyword, _ = normalize_with_map(
        keyword,
        case_sensitive=case_sensitive,
        normalize_width=normalize_width,
        ignore_whitespace=ignore_whitespace,
    )

    if not normalized_keyword:
        return []

    if not source_map:
        return []

    spans: list[tuple[int, int]] = []

    search_start = 0

    while True:
        found_position = normalized_source.find(
            normalized_keyword,
            search_start,
        )

        if found_position < 0:
            break

        normalized_end = (
            found_position
            + len(normalized_keyword)
            - 1
        )

        original_start = source_map[
            found_position
        ]

        original_end = (
            source_map[normalized_end]
            + 1
        )

        original_length = (
            original_end
            - original_start
        )

        spans.append(
            (
                original_start,
                original_length,
            )
        )

        search_start = (
            found_position
            + len(normalized_keyword)
        )

    return spans


def select_spans(
    spans: list[tuple[int, int]],
    mode: str,
    occurrence: int,
) -> list[tuple[int, int]]:
    if mode == "全部出现位置":
        return spans

    if mode == "只修改第一次":
        return spans[:1]

    if mode == "修改指定次数":
        target_index = occurrence - 1

        if 0 <= target_index < len(spans):
            return [
                spans[target_index]
            ]

        return []

    raise ValueError(
        f"未知匹配方式：{mode}"
    )


def get_anchor_cell(cell):
    """
    合并单元格时，返回合并区域左上角单元格。
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
    优先使用 Value2 读取内容。
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
    try:
        return str(cell.Address)
    except Exception:
        return "(无法读取地址)"


def has_formula(
    cell,
) -> bool:
    try:
        return bool(cell.HasFormula)
    except Exception:
        return False


def iter_unique_cells(
    target_range,
):
    """
    遍历范围中的单元格。

    合并单元格只处理一次。
    """
    visited: set[str] = set()

    for raw_cell in target_range.Cells:
        cell = get_anchor_cell(raw_cell)
        address = get_cell_address(cell)

        if address in visited:
            continue

        visited.add(address)

        yield cell


def excel_find_cells(
    scope,
    keyword: str,
    *,
    match_case: bool,
    limit: int = 5000,
) -> list:
    """
    使用 Excel 自己的 Find 功能查找单元格。

    明确设置所有查找参数，避免受到用户上一次
    Excel 查找设置的影响。
    """
    if not keyword:
        return []

    try:
        after_cell = scope.Cells.Item(
            scope.Cells.Count
        )
    except Exception:
        after_cell = scope.Cells.Item(
            1,
            1,
        )

    found_cells: list = []
    visited: set[str] = set()

    found = scope.Find(
        keyword,
        after_cell,
        XL_VALUES,
        XL_PART,
        XL_BY_ROWS,
        XL_NEXT,
        match_case,
        False,
        False,
    )

    while found is not None:
        cell = get_anchor_cell(found)
        address = get_cell_address(cell)

        if address in visited:
            break

        visited.add(address)
        found_cells.append(cell)

        if len(found_cells) >= limit:
            break

        found = scope.Find(
            keyword,
            found,
            XL_VALUES,
            XL_PART,
            XL_BY_ROWS,
            XL_NEXT,
            match_case,
            False,
            False,
        )

    return found_cells


@contextmanager
def open_excel_workbook(
    file_path: Path,
    *,
    read_only: bool,
):
    """
    安全打开和关闭 Excel。
    """
    excel = None
    workbook = None

    pythoncom.CoInitialize()

    try:
        excel = win32com.client.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(
            os.path.abspath(str(file_path)),
            UpdateLinks=0,
            ReadOnly=read_only,
        )

        yield excel, workbook

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


def locate_matching_cells(
    worksheet,
    settings: dict,
) -> list[tuple[object, list[tuple[int, int]]]]:
    """
    根据用户设置查找所有目标单元格。
    """
    keyword = settings["keyword"]

    results: list = []
    visited: set[str] = set()

    if settings["scope_mode"] == "整张工作表":
        candidate_cells = excel_find_cells(
            worksheet.UsedRange,
            keyword,
            match_case=settings[
                "case_sensitive"
            ],
        )

        for candidate in candidate_cells:
            cell = get_anchor_cell(candidate)
            address = get_cell_address(cell)

            if address in visited:
                continue

            visited.add(address)

            text = get_cell_text(cell)

            if not isinstance(text, str):
                continue

            spans = find_original_spans(
                source=text,
                keyword=keyword,
                case_sensitive=settings[
                    "case_sensitive"
                ],
                normalize_width=settings[
                    "normalize_width"
                ],
                ignore_whitespace=settings[
                    "ignore_whitespace"
                ],
            )

            spans = select_spans(
                spans,
                settings["match_mode"],
                settings["occurrence"],
            )

            if spans:
                results.append(
                    (
                        cell,
                        spans,
                    )
                )

        return results

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

        for cell in iter_unique_cells(
            target_range
        ):
            address = get_cell_address(cell)

            if address in visited:
                continue

            visited.add(address)

            text = get_cell_text(cell)

            if not isinstance(text, str):
                continue

            spans = find_original_spans(
                source=text,
                keyword=keyword,
                case_sensitive=settings[
                    "case_sensitive"
                ],
                normalize_width=settings[
                    "normalize_width"
                ],
                ignore_whitespace=settings[
                    "ignore_whitespace"
                ],
            )

            spans = select_spans(
                spans,
                settings["match_mode"],
                settings["occurrence"],
            )

            if spans:
                results.append(
                    (
                        cell,
                        spans,
                    )
                )

    return results


class ExcelTextColorApp:

    def __init__(
        self,
        root: tk.Tk,
    ) -> None:
        self.root = root

        self.root.title(APP_TITLE)
        self.root.geometry("840x760")
        self.root.minsize(760, 680)

        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.sheet_name = tk.StringVar()

        self.scope_mode = tk.StringVar(
            value="指定范围"
        )

        self.cell_range = tk.StringVar(
            value="BG7:BG28"
        )

        self.keyword = tk.StringVar()

        self.case_sensitive = tk.BooleanVar(
            value=True
        )

        self.normalize_width = tk.BooleanVar(
            value=True
        )

        self.ignore_whitespace = tk.BooleanVar(
            value=True
        )

        self.match_mode = tk.StringVar(
            value="全部出现位置"
        )

        self.occurrence_number = tk.IntVar(
            value=1
        )

        self.status = tk.StringVar(
            value="请选择 Excel 文件。"
        )

        self.selected_rgb = (
            255,
            0,
            0,
        )

        self.selected_hex = "#FF0000"

        self.build_ui()
        self.update_color_preview()
        self.update_scope_state()
        self.update_occurrence_state()

    def build_ui(
        self,
    ) -> None:
        main_frame = ttk.Frame(
            self.root,
            padding=14,
        )

        main_frame.pack(
            fill="both",
            expand=True,
        )

        # ====================================
        # 文件设置
        # ====================================

        file_frame = ttk.LabelFrame(
            main_frame,
            text="1. 文件设置",
            padding=10,
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
            pady=5,
        )

        ttk.Entry(
            file_frame,
            textvariable=self.input_file,
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            pady=5,
        )

        ttk.Button(
            file_frame,
            text="选择文件",
            command=self.choose_input_file,
        ).grid(
            row=0,
            column=2,
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            file_frame,
            text="输出文件",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )

        ttk.Entry(
            file_frame,
            textvariable=self.output_file,
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            pady=5,
        )

        ttk.Button(
            file_frame,
            text="选择位置",
            command=self.choose_output_file,
        ).grid(
            row=1,
            column=2,
            padx=(8, 0),
            pady=5,
        )

        file_frame.columnconfigure(
            1,
            weight=1,
        )

        # ====================================
        # 位置和文字
        # ====================================

        target_frame = ttk.LabelFrame(
            main_frame,
            text="2. 查找位置和文字",
            padding=10,
        )

        target_frame.pack(
            fill="x",
            pady=(10, 0),
        )

        ttk.Label(
            target_frame,
            text="工作表",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
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
            pady=5,
        )

        ttk.Button(
            target_frame,
            text="读取工作表",
            command=self.load_sheet_names,
        ).grid(
            row=0,
            column=2,
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            target_frame,
            text="处理范围",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )

        scope_combo = ttk.Combobox(
            target_frame,
            textvariable=self.scope_mode,
            values=(
                "指定范围",
                "整张工作表",
            ),
            state="readonly",
            width=18,
        )

        scope_combo.grid(
            row=1,
            column=1,
            sticky="w",
            pady=5,
        )

        scope_combo.bind(
            "<<ComboboxSelected>>",
            lambda event: self.update_scope_state(),
        )

        ttk.Label(
            target_frame,
            text="单元格/范围",
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )

        self.range_entry = ttk.Entry(
            target_frame,
            textvariable=self.cell_range,
        )

        self.range_entry.grid(
            row=2,
            column=1,
            columnspan=2,
            sticky="ew",
            pady=5,
        )

        ttk.Label(
            target_frame,
            text="例：BG7、BG7:BG28、BG7,BJ7:BJ28",
        ).grid(
            row=3,
            column=1,
            columnspan=2,
            sticky="w",
        )

        ttk.Label(
            target_frame,
            text="指定文字",
        ).grid(
            row=4,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=(10, 5),
        )

        ttk.Entry(
            target_frame,
            textvariable=self.keyword,
        ).grid(
            row=4,
            column=1,
            columnspan=2,
            sticky="ew",
            pady=(10, 5),
        )

        option_frame = ttk.Frame(
            target_frame
        )

        option_frame.grid(
            row=5,
            column=1,
            columnspan=2,
            sticky="w",
            pady=5,
        )

        ttk.Checkbutton(
            option_frame,
            text="区分大小写",
            variable=self.case_sensitive,
        ).pack(
            side="left",
        )

        ttk.Checkbutton(
            option_frame,
            text="兼容全角/半角差异",
            variable=self.normalize_width,
        ).pack(
            side="left",
            padx=(12, 0),
        )

        ttk.Checkbutton(
            option_frame,
            text="忽略空格、换行和隐藏字符",
            variable=self.ignore_whitespace,
        ).pack(
            side="left",
            padx=(12, 0),
        )

        ttk.Button(
            target_frame,
            text="先查找位置",
            command=self.preview_search,
        ).grid(
            row=6,
            column=2,
            sticky="e",
            pady=(7, 0),
        )

        target_frame.columnconfigure(
            1,
            weight=1,
        )

        # ====================================
        # 颜色和匹配
        # ====================================

        style_frame = ttk.LabelFrame(
            main_frame,
            text="3. 颜色和匹配方式",
            padding=10,
        )

        style_frame.pack(
            fill="x",
            pady=(10, 0),
        )

        ttk.Label(
            style_frame,
            text="字体颜色",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )

        ttk.Button(
            style_frame,
            text="选择颜色",
            command=self.choose_color,
        ).grid(
            row=0,
            column=1,
            sticky="w",
            pady=5,
        )

        self.color_preview = tk.Label(
            style_frame,
            text="示例文字",
            width=12,
            relief="solid",
            borderwidth=1,
        )

        self.color_preview.grid(
            row=0,
            column=2,
            padx=(12, 8),
            pady=5,
        )

        self.color_label = ttk.Label(
            style_frame,
            text=self.selected_hex,
        )

        self.color_label.grid(
            row=0,
            column=3,
            sticky="w",
            pady=5,
        )

        ttk.Label(
            style_frame,
            text="匹配方式",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=5,
        )

        match_combo = ttk.Combobox(
            style_frame,
            textvariable=self.match_mode,
            values=(
                "全部出现位置",
                "只修改第一次",
                "修改指定次数",
            ),
            state="readonly",
            width=18,
        )

        match_combo.grid(
            row=1,
            column=1,
            sticky="w",
            pady=5,
        )

        match_combo.bind(
            "<<ComboboxSelected>>",
            lambda event: (
                self.update_occurrence_state()
            ),
        )

        ttk.Label(
            style_frame,
            text="第",
        ).grid(
            row=1,
            column=2,
            sticky="e",
            padx=(12, 4),
            pady=5,
        )

        self.occurrence_spinbox = ttk.Spinbox(
            style_frame,
            from_=1,
            to=999,
            textvariable=self.occurrence_number,
            width=8,
        )

        self.occurrence_spinbox.grid(
            row=1,
            column=3,
            sticky="w",
            pady=5,
        )

        ttk.Label(
            style_frame,
            text="次出现",
        ).grid(
            row=1,
            column=4,
            sticky="w",
            padx=(4, 0),
            pady=5,
        )

        # ====================================
        # 操作按钮
        # ====================================

        action_frame = ttk.Frame(
            main_frame
        )

        action_frame.pack(
            fill="x",
            pady=(12, 0),
        )

        ttk.Button(
            action_frame,
            text="开始处理",
            command=self.process_excel,
        ).pack(
            side="right",
        )

        ttk.Button(
            action_frame,
            text="清空日志",
            command=self.clear_log,
        ).pack(
            side="right",
            padx=(0, 8),
        )

        # ====================================
        # 日志
        # ====================================

        log_frame = ttk.LabelFrame(
            main_frame,
            text="处理日志",
            padding=8,
        )

        log_frame.pack(
            fill="both",
            expand=True,
            pady=(10, 0),
        )

        self.log_text = tk.Text(
            log_frame,
            height=14,
            wrap="word",
            state="disabled",
        )

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )

        self.log_text.configure(
            yscrollcommand=scrollbar.set
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
            pady=(7, 0),
        )

    def choose_input_file(
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

        self.input_file.set(path)

        source = Path(path)

        output = source.with_name(
            f"{source.stem}_文字变色后"
            f"{source.suffix}"
        )

        self.output_file.set(
            str(output)
        )

        self.load_sheet_names()

    def choose_output_file(
        self,
    ) -> None:
        input_value = self.input_file.get().strip()

        if input_value:
            source = Path(input_value)
        else:
            source = Path.cwd() / "文字变色后.xlsx"

        output_path = filedialog.asksaveasfilename(
            title="选择输出文件",
            initialdir=str(source.parent),
            initialfile=(
                f"{source.stem}_文字变色后"
                f"{source.suffix}"
            ),
            defaultextension=(
                source.suffix or ".xlsx"
            ),
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

        if output_path:
            self.output_file.set(
                output_path
            )

    def choose_color(
        self,
    ) -> None:
        rgb_value, hex_value = colorchooser.askcolor(
            color=self.selected_hex,
            title="选择字体颜色",
        )

        if (
            rgb_value is None
            or hex_value is None
        ):
            return

        self.selected_rgb = tuple(
            round(value)
            for value in rgb_value
        )

        self.selected_hex = (
            hex_value.upper()
        )

        self.update_color_preview()

    def update_color_preview(
        self,
    ) -> None:
        red, green, blue = self.selected_rgb

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
            background=self.selected_hex,
            foreground=foreground,
        )

        self.color_label.configure(
            text=self.selected_hex
        )

    def update_scope_state(
        self,
    ) -> None:
        if self.scope_mode.get() == "指定范围":
            self.range_entry.configure(
                state="normal"
            )
        else:
            self.range_entry.configure(
                state="disabled"
            )

    def update_occurrence_state(
        self,
    ) -> None:
        if self.match_mode.get() == "修改指定次数":
            self.occurrence_spinbox.configure(
                state="normal"
            )
        else:
            self.occurrence_spinbox.configure(
                state="disabled"
            )

    def load_sheet_names(
        self,
    ) -> None:
        file_path = Path(
            self.input_file.get().strip()
        )

        if not file_path.exists():
            messagebox.showwarning(
                APP_TITLE,
                "请先选择有效的 Excel 文件。",
            )
            return

        self.status.set(
            "正在读取工作表……"
        )

        self.root.update_idletasks()

        try:
            with open_excel_workbook(
                file_path,
                read_only=True,
            ) as (_, workbook):

                sheet_names = [
                    sheet.Name
                    for sheet in workbook.Worksheets
                ]

            self.sheet_combo[
                "values"
            ] = sheet_names

            if sheet_names:
                self.sheet_name.set(
                    sheet_names[0]
                )

            self.status.set(
                f"读取完成，共 "
                f"{len(sheet_names)} 个工作表。"
            )

        except Exception as error:
            self.show_error(
                f"读取工作表失败：{error}"
            )

    def collect_settings(
        self,
        *,
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

        sheet_name = self.sheet_name.get().strip()

        if not sheet_name:
            raise ValueError(
                "请选择工作表。"
            )

        keyword = self.keyword.get()

        if not keyword:
            raise ValueError(
                "请输入需要变色的指定文字。"
            )

        ranges: list[str] = []

        if self.scope_mode.get() == "指定范围":
            ranges = parse_ranges(
                self.cell_range.get()
            )

        occurrence = int(
            self.occurrence_number.get()
        )

        if occurrence < 1:
            raise ValueError(
                "出现次数必须大于等于 1。"
            )

        settings = {
            "input_file": input_file,
            "sheet_name": sheet_name,
            "keyword": keyword,
            "scope_mode": self.scope_mode.get(),
            "ranges": ranges,
            "case_sensitive": (
                self.case_sensitive.get()
            ),
            "normalize_width": (
                self.normalize_width.get()
            ),
            "ignore_whitespace": (
                self.ignore_whitespace.get()
            ),
            "match_mode": self.match_mode.get(),
            "occurrence": occurrence,
            "rgb": self.selected_rgb,
        }

        if require_output:
            output_value = (
                self.output_file.get().strip()
            )

            if not output_value:
                raise ValueError(
                    "请选择输出文件。"
                )

            output_file = Path(
                output_value
            )

            if (
                output_file.suffix.lower()
                != input_file.suffix.lower()
            ):
                raise ValueError(
                    "输出文件扩展名必须与原文件相同。"
                )

            if (
                output_file.resolve()
                == input_file.resolve()
            ):
                raise ValueError(
                    "输出文件不能与原文件相同。"
                )

            settings[
                "output_file"
            ] = output_file

        return settings

    def preview_search(
        self,
    ) -> None:
        try:
            settings = self.collect_settings(
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
            "开始查找，不会修改文件。\n\n"
        )

        self.status.set(
            "正在查找……"
        )

        self.root.update_idletasks()

        try:
            with open_excel_workbook(
                settings["input_file"],
                read_only=True,
            ) as (_, workbook):

                worksheet = workbook.Worksheets(
                    settings["sheet_name"]
                )

                matches = locate_matching_cells(
                    worksheet,
                    settings,
                )

                if not matches:
                    self.append_log(
                        "指定范围内没有找到指定文字。\n\n"
                    )

                    other_cells = excel_find_cells(
                        worksheet.UsedRange,
                        settings["keyword"],
                        match_case=settings[
                            "case_sensitive"
                        ],
                        limit=50,
                    )

                    if other_cells:
                        self.append_log(
                            "但在工作表的其他位置找到：\n"
                        )

                        for cell in other_cells:
                            text = (
                                get_cell_text(cell)
                                or ""
                            )

                            preview = (
                                text
                                .replace("\r", " ")
                                .replace("\n", " ")
                            )

                            if len(preview) > 100:
                                preview = (
                                    preview[:100]
                                    + "…"
                                )

                            self.append_log(
                                f"{get_cell_address(cell)}："
                                f"{preview}\n"
                            )

                        self.append_log(
                            "\n请把上面的地址输入到"
                            "“单元格/范围”，或者选择"
                            "“整张工作表”。\n"
                        )

                    self.status.set(
                        "查找完成：指定范围内未找到。"
                    )

                    return

                formula_count = 0

                self.append_log(
                    f"找到 {len(matches)} 个单元格：\n\n"
                )

                for cell, spans in matches:
                    text = (
                        get_cell_text(cell)
                        or ""
                    )

                    preview = (
                        text
                        .replace("\r", " ")
                        .replace("\n", " ")
                    )

                    if len(preview) > 120:
                        preview = (
                            preview[:120]
                            + "…"
                        )

                    formula_text = ""

                    if has_formula(cell):
                        formula_count += 1
                        formula_text = "【公式单元格】"

                    self.append_log(
                        f"{get_cell_address(cell)} "
                        f"{formula_text}"
                        f"匹配 {len(spans)} 处\n"
                        f"  {preview}\n\n"
                    )

                if formula_count:
                    self.append_log(
                        "注意：存在公式单元格。"
                        "Excel 无法在保留公式的同时，"
                        "只改变公式结果中的部分文字颜色。\n"
                    )

                self.status.set(
                    f"查找完成：找到 "
                    f"{len(matches)} 个单元格。"
                )

        except Exception as error:
            self.show_error(
                f"查找失败：{error}"
            )

    def process_excel(
        self,
    ) -> None:
        try:
            settings = self.collect_settings(
                require_output=True
            )
        except ValueError as error:
            messagebox.showwarning(
                APP_TITLE,
                str(error),
            )
            return

        output_file: Path = settings[
            "output_file"
        ]

        if output_file.exists():
            overwrite = messagebox.askyesno(
                APP_TITLE,
                f"输出文件已经存在：\n\n"
                f"{output_file}\n\n"
                f"是否覆盖？",
            )

            if not overwrite:
                return

        self.clear_log()

        self.append_log(
            "开始处理。\n"
            "请先关闭原文件和输出文件。\n\n"
        )

        self.status.set(
            "正在处理……"
        )

        self.root.update_idletasks()

        try:
            output_file.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            shutil.copy2(
                settings["input_file"],
                output_file,
            )

            self.append_log(
                f"已生成副本：\n"
                f"{output_file}\n\n"
            )

            with open_excel_workbook(
                output_file,
                read_only=False,
            ) as (_, workbook):

                worksheet = workbook.Worksheets(
                    settings["sheet_name"]
                )

                matches = locate_matching_cells(
                    worksheet,
                    settings,
                )

                if not matches:
                    self.append_log(
                        "指定范围内没有找到指定文字。\n\n"
                    )

                    other_cells = excel_find_cells(
                        worksheet.UsedRange,
                        settings["keyword"],
                        match_case=settings[
                            "case_sensitive"
                        ],
                        limit=50,
                    )

                    if other_cells:
                        self.append_log(
                            "但在工作表的以下位置找到了：\n"
                        )

                        addresses = []

                        for cell in other_cells:
                            addresses.append(
                                get_cell_address(cell)
                            )

                        self.append_log(
                            ", ".join(addresses)
                            + "\n\n"
                        )

                        self.append_log(
                            "请填写正确地址，或者将"
                            "处理范围改为“整张工作表”。\n"
                        )

                    workbook.Save()

                    self.status.set(
                        "未找到指定文字。"
                    )

                    messagebox.showwarning(
                        APP_TITLE,
                        "指定范围内没有找到指定文字，"
                        "因此没有变色。\n\n"
                        "请查看处理日志中的正确地址，"
                        "或者选择“整张工作表”。",
                    )

                    return

                excel_color = rgb_to_excel_color(
                    settings["rgb"]
                )

                changed_count = 0
                formula_skipped = 0
                failed_count = 0

                for cell, spans in matches:
                    address = get_cell_address(cell)

                    if has_formula(cell):
                        formula_skipped += 1

                        self.append_log(
                            f"[跳过] {address}："
                            f"这是公式单元格，"
                            f"无法只改变部分文字颜色。\n"
                        )

                        continue

                    changed_in_cell = 0

                    for start, length in spans:
                        try:
                            # Excel 字符位置从 1 开始。
                            # 这里使用位置参数，避免
                            # pywin32 命名参数兼容问题。
                            characters = cell.Characters(
                                start + 1,
                                length,
                            )

                            characters.Font.Color = (
                                excel_color
                            )

                            changed_in_cell += 1

                        except Exception as error:
                            failed_count += 1

                            self.append_log(
                                f"[失败] {address}："
                                f"开始位置 {start + 1}，"
                                f"长度 {length}，"
                                f"错误：{error}\n"
                            )

                    if changed_in_cell:
                        changed_count += (
                            changed_in_cell
                        )

                        text = (
                            get_cell_text(cell)
                            or ""
                        )

                        preview = (
                            text
                            .replace("\r", " ")
                            .replace("\n", " ")
                        )

                        if len(preview) > 90:
                            preview = (
                                preview[:90]
                                + "…"
                            )

                        self.append_log(
                            f"[完成] {address}："
                            f"修改 {changed_in_cell} 处\n"
                            f"  {preview}\n"
                        )

                workbook.Save()

            self.append_log(
                "\n处理完成。\n"
                f"找到单元格：{len(matches)}\n"
                f"成功变色：{changed_count}\n"
                f"公式单元格跳过：{formula_skipped}\n"
                f"失败：{failed_count}\n"
                f"保存位置：{output_file}\n"
            )

            self.status.set(
                f"处理完成，成功变色 "
                f"{changed_count} 处。"
            )

            messagebox.showinfo(
                APP_TITLE,
                "处理完成。\n\n"
                f"成功变色：{changed_count} 处\n"
                f"公式单元格跳过：{formula_skipped}\n"
                f"失败：{failed_count}\n\n"
                f"保存位置：\n"
                f"{output_file}",
            )

        except Exception as error:
            self.show_error(
                f"处理失败：{error}"
            )

    def clear_log(
        self,
    ) -> None:
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

        self.root.update_idletasks()

    def show_error(
        self,
        message: str,
    ) -> None:
        self.status.set(
            "操作失败。"
        )

        details = traceback.format_exc()

        self.append_log(
            f"\n{message}\n\n"
            f"{details}\n"
        )

        messagebox.showerror(
            APP_TITLE,
            message,
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