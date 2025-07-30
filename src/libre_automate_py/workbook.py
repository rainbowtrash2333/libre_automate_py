from __future__ import annotations
import uno
from ooodev.calc import CalcDoc
from ooodev.utils.file_io import FileIO
from ooodev.loader import Lo
from ooodev.utils.gui import GUI
from ooodev.utils.type_var import PathOrStr
from ooodev.calc import CalcDoc, CalcSheet, ZoomKind, CalcSheetView
from ooodev.office.calc import Calc
from typing import Tuple
from officeLoader import OfficeLoader
from ooodev.format.calc.direct.cell.borders import BorderLineKind
from ooodev.formatters.formatter_table import FormatterTable, FormatTableItem
from ooodev.utils.color import CommonColor
from ooodev.format.calc.direct.cell.borders import Side
from myutil import *
import pandas as pd


class Workbook:
    def __init__(self, read_only: bool = False, filepath: str | None = None, visible: bool = True) -> None:
        self._read_only = read_only
        self._filepath = filepath
        self._visible = visible
        self.doc = None

        try:
            office_loader = OfficeLoader()
            loader = office_loader.get_loader()
            if self._filepath:
                self._input_fnm = FileIO.get_absolute_path(self._filepath)
                self.doc = CalcDoc.open_doc(fnm=self._input_fnm, loader=loader, visible=self._visible)
            else:
                self.doc = CalcDoc.create_doc(visible=True)

        except Exception:
            Lo.close_office()
            raise

    def save(self, save_path: str | None = None) -> None:
        if not self.doc:
            raise RuntimeError("No document to save.")

        path = save_path or self._filepath
        if not path:
            raise ValueError("No file path specified for saving.")
        out_file = FileIO.get_absolute_path(path)
        _ = FileIO.make_directory(out_file)
        out_fnm = out_file
        try:
            self.doc.save_doc(fnm=out_fnm)
        except Exception:
            Lo.close_office()
            raise

    def get_range_value(self, sheet_n: int, range_name: str) -> Tuple[Tuple, ...]:
        cell_rng = Calc.get_range_obj(range_name="A1:B2")
        return self.doc.sheets[sheet_n].get_array(range_obj=cell_rng)

    def close(self):
        self.doc.close_doc()
        return 0

    def get_used_value(self, sheet_n: int, range_name: str = None) -> Tuple[Tuple, ...]:

        used_rng = self.doc.sheets[sheet_n].find_used_range_obj()
        # start_idx = used_rng.start_row_index
        # end_idx = used_rng.end_row_index
        # start_col = used_rng.start_col_index
        # end_idx = used_rng.end_col_index
        if range_name is None:
            return self.doc.sheets[sheet_n].get_array(range_obj=used_rng)
        return self.doc.sheets[sheet_n].get_array(range_name=range_name)

    def set_array_value(self, sheet_n: int, values: Tuple[Tuple, ...], range_name: str) -> None:
        self.doc.sheets[sheet_n].set_array(values=values, name=range_name)

    def get_end_name(self, sheet_n) -> str:
        used_rng = self.doc.sheets[sheet_n].find_used_range_obj()
        end_cell = used_rng.cell_end
        return f"{end_cell.col}{end_cell.row}"

    def formatter_range(self, sheet_n, range_name: str):
        rng = self.doc.sheets[sheet_n].get_range(range_name=range_name)
        rng.style_borders(
            border_side=Side(color=CommonColor.BLACK, width=1),
            horizontal=Side(color=CommonColor.BLACK, width=1),
            vertical=Side(color=CommonColor.BLACK, width=1),
        )

        range_list = convert_range_name_to_list(range_name)
        fl = FormatterTable(format=(".2f", ">9"), idxs=(range_list[0], range_list[3]))

    def set_pandas_range(self, data: pd.DataFrame, sheet_n: int, cell_name: str) -> None:
        result = data.values.tolist()
        self.set_array_value(sheet_n, result, cell_name)
        self.formatter_range(sheet_n, f"{cell_name}:{self.get_end_name(sheet_n)}")

    # RangeObj
    # CalcCellRange

    def merge_same_cells(self, sheet_n: int, start_cell_name: str, merge_list=None) -> None:
        used_rng = self.doc.sheets[sheet_n].find_used_range_obj()
        start_list = convert_cell_name_to_list(start_cell_name)
        end_idx = used_rng.end_row_index
        col_idx = start_list[0]
        sheet = self.doc.get_sheet(idx=sheet_n)
        start_row_idx = start_list[1]
        next_row_idx = start_row_idx + 1
        idx = 0
        while next_row_idx <= end_idx + 1:
            start_cell = sheet.get_cell(col=col_idx, row=start_row_idx)
            next_cell = sheet.get_cell(col=col_idx, row=next_row_idx)
            merge_flag = True
            if merge_list is not None and idx < len(merge_list) - 1:
                merge_flag = merge_list[idx] == merge_list[idx + 1]
                idx = idx + 1
            if start_cell.value == next_cell.value and merge_flag:
                next_row_idx = next_row_idx + 1
            else:
                if next_row_idx > start_row_idx + 1:
                    sheet.get_range(col_start=col_idx, row_start=start_row_idx, col_end=col_idx,
                                    row_end=next_row_idx - 1).merge_cells(center=True)
                start_row_idx = next_row_idx
                next_row_idx = next_row_idx + 1

    def merge_cells_by_index(self, sheet_n: int, start_cell_name: str, index: []):
        merge_ranges = []
        n = len(index)
        if n == 0:
            return merge_ranges
        start_index = 0
        current_value = index[0]
        sheet = self.doc.get_sheet(idx=sheet_n)
        start_cell_name_list = convert_cell_name_to_list(start_cell_name)
        col_name = get_cell_col_name(start_cell_name)
        for j in range(1, n):
            if index[j] != current_value:
                if start_index <= j - 1:
                    # 生成合并区域（至少两个单元格才合并）
                    if (j - 1 - start_index) >= 1:
                        start_row = start_cell_name_list[1] + 1 + start_index
                        end_row = start_cell_name_list[1] + 1 + (j - 1)
                        merge_range = f"{col_name}{start_row}:{col_name}{end_row}"
                        merge_ranges.append(merge_range)
                start_index = j
                current_value = index[j]

        # 处理最后一个区间
        if start_index <= n - 1:
            if (n - 1 - start_index) >= 1:
                start_row = start_cell_name_list[1] + 1 + start_index
                end_row = start_cell_name_list[1] + 1 + (n - 1)
                merge_range = f"{col_name}{start_row}:{col_name}{end_row}"
                merge_ranges.append(merge_range)

        for m in merge_ranges:
            sheet.get_range(range_name=m).merge_cells(center=True)
            # return merge_ranges

    def sum_col(self, sheet_n: int, sum_cell_name: str, end_cell_name: None | str = None) -> None:
        sheet = self.doc.get_sheet(idx=sheet_n)
        cell = sheet.get_cell(cell_name=sum_cell_name)
        sum_cell_list = convert_cell_name_to_list(sum_cell_name)
        col_name = get_cell_col_name(cell_name=sum_cell_name)
        range_name = None
        if end_cell_name is None:
            used_rng = sheet.find_used_range_obj()
            start_idx = sum_cell_list[1] + 2
            end_idx = used_rng.end_row_index + 1
            range_name = f"{col_name}{start_idx}:{col_name}{end_idx}"
        else:
            used_list = convert_cell_name_to_list(end_cell_name)
            end_idx = used_list[1] + 1
            start_idx = sum_cell_list[1] + 2 if sum_cell_list[1] < end_idx else sum_cell_list[1]
            range_name = f"{col_name}{start_idx}:{col_name}{end_idx}"
        # print(f"=SUM({range_name})")
        cell.set_val(f"=SUM({range_name})")
