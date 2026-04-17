"""Tests for workbook normalizer."""

from src.models.schema import (
    CellData,
    ColumnConfig,
    FontStyle,
    Protection,
    RowConfig,
    SheetData,
    WorkbookData,
)
from src.normalizers.workbook_normalizer import normalize


def _make_workbook(**cell_overrides) -> WorkbookData:
    cell = CellData(value="test", **cell_overrides)
    sheet = SheetData(
        name="Sheet1",
        index=0,
        cells={"A1": cell},
        rows={"1": RowConfig(height=20.0)},
        columns={"A": ColumnConfig(width=10.0)},
    )
    return WorkbookData(sourceType="excel", sourceName="test.xlsx", sheets=[sheet])


class TestNormalize:
    def test_returns_workbook_data(self):
        wb = _make_workbook()
        out = normalize(wb)
        assert out.sourceType == "excel"
        assert len(out.sheets) == 1

    def test_hex_color_normalization(self):
        wb = _make_workbook(fill="ff336699")
        out = normalize(wb)
        assert out.sheets[0].cells["A1"].fill == "#336699"

    def test_short_hex(self):
        wb = _make_workbook(fill="abc")
        out = normalize(wb)
        assert out.sheets[0].cells["A1"].fill == "#AABBCC"

    def test_none_fill_stays_none(self):
        wb = _make_workbook(fill=None)
        out = normalize(wb)
        assert out.sheets[0].cells["A1"].fill is None

    def test_font_color_normalized(self):
        wb = _make_workbook(font=FontStyle(color="FF112233"))
        out = normalize(wb)
        assert out.sheets[0].cells["A1"].font.color == "#112233"

    def test_protection_preserved(self):
        wb = _make_workbook(protection=Protection(locked=True, hidden=True))
        out = normalize(wb)
        p = out.sheets[0].cells["A1"].protection
        assert p.locked is True
        assert p.hidden is True

    def test_merged_flag_preserved(self):
        wb = _make_workbook(merged=True)
        out = normalize(wb)
        assert out.sheets[0].cells["A1"].merged is True

    def test_column_key_normalization(self):
        """Numeric column keys should be converted to letters."""
        cell = CellData(value="test")
        sheet = SheetData(
            name="Sheet1",
            index=0,
            cells={"A1": cell},
            rows={},
            columns={"1": ColumnConfig(width=10.0)},
        )
        wb = WorkbookData(sourceType="excel", sourceName="test.xlsx", sheets=[sheet])
        out = normalize(wb)
        assert "A" in out.sheets[0].columns

    def test_detect_hidden_false_clears_sheet_hidden(self):
        sheet = SheetData(
            name="Sheet1",
            index=0,
            hidden=True,
            cells={"A1": CellData(value="x")},
        )
        wb = WorkbookData(sourceType="excel", sourceName="test.xlsx", sheets=[sheet])
        out = normalize(wb, detect_hidden=False)
        assert out.sheets[0].hidden is False

    def test_detect_hidden_false_clears_row_column_hidden(self):
        sheet = SheetData(
            name="Sheet1",
            index=0,
            cells={"A1": CellData(value="x")},
            rows={"1": RowConfig(height=20.0, hidden=True)},
            columns={"A": ColumnConfig(width=10.0, hidden=True)},
        )
        wb = WorkbookData(sourceType="excel", sourceName="test.xlsx", sheets=[sheet])
        out = normalize(wb, detect_hidden=False)
        assert out.sheets[0].rows["1"].hidden is False
        assert out.sheets[0].columns["A"].hidden is False

    def test_detect_hidden_false_clears_celldata_hidden(self):
        wb = _make_workbook(hidden=True)
        out = normalize(wb, detect_hidden=False)
        assert out.sheets[0].cells["A1"].hidden is False

    def test_detect_hidden_false_preserves_protection_hidden(self):
        wb = _make_workbook(protection=Protection(locked=False, hidden=True))
        out = normalize(wb, detect_hidden=False)
        p = out.sheets[0].cells["A1"].protection
        assert p is not None
        assert p.hidden is True
