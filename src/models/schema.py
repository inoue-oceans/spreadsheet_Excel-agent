from __future__ import annotations

from typing import Literal

from pydantic import BaseModel


class FontStyle(BaseModel):
    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: str | None = None  # HEX


class Alignment(BaseModel):
    horizontal: str | None = None
    vertical: str | None = None
    wrapText: bool = False


class BorderSide(BaseModel):
    style: str | None = None
    color: str | None = None


class Border(BaseModel):
    top: BorderSide | None = None
    bottom: BorderSide | None = None
    left: BorderSide | None = None
    right: BorderSide | None = None


class Protection(BaseModel):
    locked: bool = True
    hidden: bool = False


class CellData(BaseModel):
    value: str | int | float | bool | None = None
    displayValue: str | None = None
    formula: str | None = None
    fill: str | None = None  # HEX
    font: FontStyle | None = None
    alignment: Alignment | None = None
    numberFormat: str | None = None
    border: Border | None = None
    protection: Protection | None = None
    hidden: bool = False
    merged: bool = False


class RowConfig(BaseModel):
    height: float | None = None
    hidden: bool = False


class ColumnConfig(BaseModel):
    width: float | None = None
    hidden: bool = False


class ConditionalFormatCondition(BaseModel):
    operator: str | None = None
    values: list[str] = []


class ConditionalFormatStyle(BaseModel):
    fill: str | None = None
    fontColor: str | None = None
    bold: bool | None = None


class ConditionalFormattingRule(BaseModel):
    priority: int
    ranges: list[str]
    type: str
    condition: ConditionalFormatCondition | None = None
    format: ConditionalFormatStyle | None = None


class SheetData(BaseModel):
    name: str
    index: int
    hidden: bool = False
    usedRange: str | None = None
    rows: dict[str, RowConfig] = {}
    columns: dict[str, ColumnConfig] = {}
    mergedRanges: list[str] = []
    conditionalFormattingRules: list[ConditionalFormattingRule] = []
    cells: dict[str, CellData] = {}


class WorkbookData(BaseModel):
    sourceType: Literal["excel", "google_sheets"]
    sourceName: str
    sheets: list[SheetData] = []


class WorkbookOutput(BaseModel):
    workbook: WorkbookData
