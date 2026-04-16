"""Tests for Excel parser."""

import pytest
from io import BytesIO
from pathlib import Path

from src.parsers.excel_parser import parse_excel
from src.exceptions import UnsupportedFileTypeError, WorkbookReadError

SAMPLE = Path(__file__).parent / "sample_test.xlsx"


class TestParseExcel:
    def test_parse_returns_workbook_data(self):
        wb = parse_excel(SAMPLE)
        assert wb.sourceType == "excel"
        assert wb.sourceName == "sample_test.xlsx"

    def test_sheet_count(self):
        wb = parse_excel(SAMPLE)
        assert len(wb.sheets) == 3

    def test_sheet_names(self):
        wb = parse_excel(SAMPLE)
        names = [s.name for s in wb.sheets]
        assert names[0] == "売上表"
        assert names[2] == "サマリー"

    def test_hidden_sheet(self):
        wb = parse_excel(SAMPLE)
        assert wb.sheets[1].hidden is True
        assert wb.sheets[0].hidden is False

    def test_row_height(self):
        wb = parse_excel(SAMPLE)
        s = wb.sheets[0]
        assert s.rows["1"].height == 30.0
        assert s.rows["2"].height == 22.0

    def test_hidden_row(self):
        wb = parse_excel(SAMPLE)
        assert wb.sheets[0].rows["8"].hidden is True

    def test_column_width(self):
        wb = parse_excel(SAMPLE)
        assert wb.sheets[0].columns["A"].width == 15.0
        assert wb.sheets[0].columns["B"].width == 20.0

    def test_hidden_column(self):
        wb = parse_excel(SAMPLE)
        assert wb.sheets[0].columns["F"].hidden is True

    def test_merged_ranges(self):
        wb = parse_excel(SAMPLE)
        mr = wb.sheets[0].mergedRanges
        assert "A1:E1" in mr
        assert "A10:D10" in mr

    def test_merged_cell_flag(self):
        wb = parse_excel(SAMPLE)
        cells = wb.sheets[0].cells
        assert cells["A1"].merged is True
        assert cells["B1"].merged is True
        assert cells["A3"].merged is False

    def test_cell_value(self):
        wb = parse_excel(SAMPLE)
        cells = wb.sheets[0].cells
        assert cells["C3"].value == 89800
        assert cells["D3"].value == 15

    def test_formula(self):
        wb = parse_excel(SAMPLE)
        cells = wb.sheets[0].cells
        assert cells["E3"].formula == "=C3*D3"
        assert cells["E10"].formula == "=SUM(E3:E9)"

    def test_font(self):
        wb = parse_excel(SAMPLE)
        font = wb.sheets[0].cells["A1"].font
        assert font.name == "Meiryo UI"
        assert font.size == 16.0
        assert font.bold is True

    def test_fill_color(self):
        wb = parse_excel(SAMPLE)
        cells = wb.sheets[0].cells
        assert cells["A1"].fill == "#D6E4F0"
        assert cells["A2"].fill == "#2E75B6"
        assert cells["B1"].fill is None  # empty merged cell

    def test_alignment(self):
        wb = parse_excel(SAMPLE)
        align = wb.sheets[0].cells["A1"].alignment
        assert align.horizontal == "center"
        assert align.vertical == "center"

    def test_number_format(self):
        wb = parse_excel(SAMPLE)
        assert wb.sheets[0].cells["C3"].numberFormat == '#,##0"円"'

    def test_conditional_formatting(self):
        wb = parse_excel(SAMPLE)
        cf = wb.sheets[0].conditionalFormattingRules
        assert len(cf) == 1
        assert cf[0].ranges == ["E3:E9"]
        assert cf[0].type == "cellIs"
        assert cf[0].condition.operator == "greaterThan"

    def test_all_cells_output(self):
        wb = parse_excel(SAMPLE)
        assert len(wb.sheets[0].cells) == 60  # 10 rows * 6 cols

    def test_cross_sheet_formula(self):
        wb = parse_excel(SAMPLE)
        assert "売上表" in wb.sheets[2].cells["B2"].formula

    def test_unsupported_file_type(self):
        with pytest.raises(UnsupportedFileTypeError):
            parse_excel("test.csv")

    def test_corrupted_file(self):
        with pytest.raises(WorkbookReadError):
            parse_excel(BytesIO(b"not xlsx"), source_name="bad.xlsx")

    def test_bytesio_input(self):
        with open(SAMPLE, "rb") as f:
            data = f.read()
        wb = parse_excel(BytesIO(data), source_name="from_memory.xlsx")
        assert wb.sourceName == "from_memory.xlsx"
        assert len(wb.sheets) == 3
