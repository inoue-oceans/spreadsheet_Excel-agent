import re

from openpyxl.utils.cell import column_index_from_string

from src.models.schema import (
    Alignment,
    CellData,
    ConditionalFormattingRule,
    FontStyle,
    SheetData,
    WorkbookOutput,
)


def _md_cell(text: object) -> str:
    s = "" if text is None else str(text)
    return s.replace("|", "\\|").replace("\n", " ").replace("\r", "")


def _font_summary(f: FontStyle | None) -> str:
    if f is None:
        return ""
    parts: list[str] = []
    if f.name:
        parts.append(f.name)
    if f.size is not None:
        parts.append(f"{f.size}pt")
    if f.bold:
        parts.append("bold")
    if f.italic:
        parts.append("italic")
    if f.underline:
        parts.append("underline")
    if f.color:
        parts.append(f"color={f.color}")
    return "; ".join(parts)


def _align_summary(a: Alignment | None) -> str:
    if a is None:
        return ""
    parts: list[str] = []
    if a.horizontal:
        parts.append(f"h={a.horizontal}")
    if a.vertical:
        parts.append(f"v={a.vertical}")
    if a.wrapText:
        parts.append("wrap")
    return "; ".join(parts)


def _cond_condition_summary(rule: ConditionalFormattingRule) -> str:
    c = rule.condition
    if c is None:
        return ""
    op = c.operator or ""
    vals = ", ".join(c.values) if c.values else ""
    return f"{op} {vals}".strip()


def _cond_format_summary(rule: ConditionalFormattingRule) -> str:
    s = rule.format
    if s is None:
        return ""
    parts: list[str] = []
    if s.fill:
        parts.append(f"fill={s.fill}")
    if s.fontColor:
        parts.append(f"font={s.fontColor}")
    if s.bold is not None:
        parts.append(f"bold={s.bold}")
    return "; ".join(parts)


def _sheet_md(sheet: SheetData) -> str:
    lines: list[str] = []
    lines.append(f"## Sheet: {_md_cell(sheet.name)}")
    lines.append("")
    lines.append(f"- **Index**: {sheet.index}")
    lines.append(f"- **Hidden**: {sheet.hidden}")
    lines.append(f"- **Used range**: {_md_cell(sheet.usedRange)}")
    lines.append("")

    lines.append("### Row configuration")
    lines.append("")
    lines.append("| Row | Height | Hidden |")
    lines.append("| --- | ------ | ------ |")
    for rk in sorted(sheet.rows.keys(), key=lambda x: int(x) if str(x).isdigit() else x):
        rc = sheet.rows[rk]
        lines.append(f"| {rk} | {_md_cell(rc.height)} | {rc.hidden} |")
    if not sheet.rows:
        lines.append("| — | — | — |")
    lines.append("")

    lines.append("### Column configuration")
    lines.append("")
    lines.append("| Column | Width | Hidden |")
    lines.append("| ------ | ----- | ------ |")
    def _col_order(k: str) -> tuple[int, str]:
        if str(k).isdigit():
            return (int(k), "")
        try:
            return (column_index_from_string(str(k).upper()), "")
        except ValueError:
            return (10_000, str(k))

    for ck in sorted(sheet.columns.keys(), key=_col_order):
        cc = sheet.columns[ck]
        lines.append(f"| {ck} | {_md_cell(cc.width)} | {cc.hidden} |")
    if not sheet.columns:
        lines.append("| — | — | — |")
    lines.append("")

    lines.append("### Merged ranges")
    lines.append("")
    if sheet.mergedRanges:
        for m in sheet.mergedRanges:
            lines.append(f"- `{_md_cell(m)}`")
    else:
        lines.append("- *(none)*")
    lines.append("")

    lines.append("### Conditional formatting rules")
    lines.append("")
    lines.append("| Priority | Ranges | Type | Condition | Format |")
    lines.append("| -------- | ------ | ---- | --------- | ------ |")
    for rule in sorted(sheet.conditionalFormattingRules, key=lambda r: r.priority):
        ranges = ", ".join(rule.ranges) if rule.ranges else ""
        lines.append(
            f"| {rule.priority} | {_md_cell(ranges)} | {_md_cell(rule.type)} | "
            f"{_md_cell(_cond_condition_summary(rule))} | {_md_cell(_cond_format_summary(rule))} |"
        )
    if not sheet.conditionalFormattingRules:
        lines.append("| — | — | — | — | — |")
    lines.append("")

    lines.append("### Cell details")
    lines.append("")
    lines.append(
        "| Cell | Value | Formula | Font | Fill | Alignment | numberFormat | merged |"
    )
    lines.append(
        "| ---- | ----- | ------- | ---- | ---- | --------- | ------------ | ------ |"
    )

    def sort_key(ref: str) -> tuple[int, int]:
        m = re.match(r"^([A-Za-z]+)(\d+)$", ref)
        if not m:
            return (0, 0)
        try:
            col = column_index_from_string(m.group(1).upper())
        except ValueError:
            col = 0
        row = int(m.group(2))
        return (row, col)

    for cref in sorted(sheet.cells.keys(), key=sort_key):
        c: CellData = sheet.cells[cref]
        lines.append(
            f"| `{cref}` | {_md_cell(c.value)} | {_md_cell(c.formula)} | "
            f"{_md_cell(_font_summary(c.font))} | {_md_cell(c.fill)} | "
            f"{_md_cell(_align_summary(c.alignment))} | {_md_cell(c.numberFormat)} | {c.merged} |"
        )
    if not sheet.cells:
        lines.append("| — | — | — | — | — | — | — | — |")
    lines.append("")
    return "\n".join(lines)


def export_markdown(output: WorkbookOutput) -> str:
    wb = output.workbook
    parts: list[str] = []
    parts.append("# Workbook export")
    parts.append("")
    parts.append("## Source")
    parts.append("")
    parts.append(f"- **Type**: {wb.sourceType}")
    parts.append(f"- **Name**: {wb.sourceName}")
    parts.append("")

    parts.append("## Sheet list")
    parts.append("")
    parts.append("| # | Name | Hidden | Used range |")
    parts.append("| - | ---- | ------ | ---------- |")
    for s in sorted(wb.sheets, key=lambda x: x.index):
        parts.append(
            f"| {s.index} | {_md_cell(s.name)} | {s.hidden} | {_md_cell(s.usedRange)} |"
        )
    parts.append("")

    for s in sorted(wb.sheets, key=lambda x: x.index):
        parts.append(_sheet_md(s))

    return "\n".join(parts).rstrip() + "\n"
