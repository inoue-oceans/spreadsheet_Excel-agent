"""Tests for exporters."""

import json
from pathlib import Path

from src.parsers.excel_parser import parse_excel
from src.normalizers.workbook_normalizer import normalize
from src.exporters.json_exporter import export_json
from src.exporters.markdown_exporter import export_markdown
from src.exporters.prompt_exporter import export_prompt

SAMPLE = Path(__file__).parent / "sample_test.xlsx"


def _get_output():
    wb = parse_excel(SAMPLE)
    return normalize(wb)


class TestJsonExporter:
    def test_valid_json(self):
        output = _get_output()
        j = export_json(output)
        data = json.loads(j)
        assert "workbook" in data

    def test_all_sheets_present(self):
        output = _get_output()
        data = json.loads(export_json(output))
        assert len(data["workbook"]["sheets"]) == 3

    def test_cell_keys_present(self):
        output = _get_output()
        data = json.loads(export_json(output))
        cell = data["workbook"]["sheets"][0]["cells"]["A1"]
        for key in ["value", "displayValue", "formula", "fill", "font", "alignment", "numberFormat", "border", "protection", "hidden", "merged"]:
            assert key in cell, f"Missing key: {key}"


class TestMarkdownExporter:
    def test_has_headings(self):
        md = export_markdown(_get_output())
        assert "# Workbook export" in md
        assert "## Sheet:" in md

    def test_has_cell_table(self):
        md = export_markdown(_get_output())
        assert "| Cell |" in md
        assert "| `A1`" in md

    def test_has_conditional_formatting(self):
        md = export_markdown(_get_output())
        assert "Conditional formatting" in md
        assert "cellIs" in md

    def test_sheet_list(self):
        md = export_markdown(_get_output())
        assert "Sheet list" in md


class TestPromptExporter:
    def test_has_instructions(self):
        pr = export_prompt(_get_output())
        assert "Instructions to recreate" in pr

    def test_has_sheet_sections(self):
        pr = export_prompt(_get_output())
        assert "売上表" in pr
        assert "サマリー" in pr

    def test_has_formula_instructions(self):
        pr = export_prompt(_get_output())
        assert "=C3*D3" in pr
        assert "=SUM(E3:E9)" in pr

    def test_has_formatting_instructions(self):
        pr = export_prompt(_get_output())
        assert "Meiryo UI" in pr
        assert "#D6E4F0" in pr
