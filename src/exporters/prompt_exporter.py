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


def _describe_font(f: FontStyle | None) -> str:
    if f is None:
        return "default font"
    bits: list[str] = []
    if f.name:
        bits.append(f"font family “{f.name}”")
    if f.size is not None:
        bits.append(f"size {f.size} pt")
    styles: list[str] = []
    if f.bold:
        styles.append("bold")
    if f.italic:
        styles.append("italic")
    if f.underline:
        styles.append("underlined")
    if styles:
        bits.append(", ".join(styles))
    if f.color:
        bits.append(f"foreground {f.color}")
    return "; ".join(bits) if bits else "default font"


def _describe_alignment(a: Alignment | None) -> str:
    if a is None:
        return "default alignment"
    bits: list[str] = []
    if a.horizontal:
        bits.append(f"horizontal {a.horizontal}")
    if a.vertical:
        bits.append(f"vertical {a.vertical}")
    if a.wrapText:
        bits.append("wrap text enabled")
    return "; ".join(bits) if bits else "default alignment"


def _describe_cond_rule(rule: ConditionalFormattingRule) -> str:
    lines: list[str] = []
    lines.append(
        f"- Rule priority {rule.priority}, type “{rule.type}”, applies to ranges: "
        f"{', '.join(rule.ranges) if rule.ranges else '(none)'}"
    )
    if rule.condition:
        c = rule.condition
        op = c.operator or "condition"
        vals = ", ".join(f"“{v}”" for v in c.values) if c.values else ""
        lines.append(f"  Condition: {op} {vals}".rstrip())
    if rule.format:
        s = rule.format
        fmt_bits: list[str] = []
        if s.fill:
            fmt_bits.append(f"fill {s.fill}")
        if s.fontColor:
            fmt_bits.append(f"font color {s.fontColor}")
        if s.bold is not None:
            fmt_bits.append(f"bold={s.bold}")
        if fmt_bits:
            lines.append(f"  Format when true: {'; '.join(fmt_bits)}")
    return "\n".join(lines)


def _sheet_prompt(sheet: SheetData) -> str:
    lines: list[str] = []
    lines.append(f"### Sheet “{sheet.name}” (order index {sheet.index})")
    lines.append(
        f"This sheet is {'hidden' if sheet.hidden else 'visible'}. "
        f"The used data range is {sheet.usedRange or 'not specified'}."
    )
    lines.append("")
    lines.append("**Row heights and visibility**")
    if sheet.rows:
        for rk in sorted(sheet.rows.keys(), key=lambda x: int(x) if str(x).isdigit() else 0):
            rc = sheet.rows[rk]
            h = f"height {rc.height}" if rc.height is not None else "default height"
            vis = "hidden" if rc.hidden else "visible"
            lines.append(f"- Row {rk}: {h}; {vis}.")
    else:
        lines.append("- No explicit row metadata was captured.")
    lines.append("")
    lines.append("**Column widths and visibility**")
    if sheet.columns:
        def _col_order(k: str) -> tuple[int, str]:
            if str(k).isdigit():
                return (int(k), "")
            try:
                return (column_index_from_string(str(k).upper()), "")
            except ValueError:
                return (10_000, str(k))

        keys = sorted(sheet.columns.keys(), key=_col_order)
        for ck in keys:
            cc = sheet.columns[ck]
            w = f"width {cc.width}" if cc.width is not None else "default width"
            vis = "hidden" if cc.hidden else "visible"
            lines.append(f"- Column {ck}: {w}; {vis}.")
    else:
        lines.append("- No explicit column metadata was captured.")
    lines.append("")
    lines.append("**Merged cell ranges**")
    if sheet.mergedRanges:
        for m in sheet.mergedRanges:
            lines.append(f"- Merge cells across range {m}.")
    else:
        lines.append("- No merged ranges.")
    lines.append("")
    lines.append("**Conditional formatting**")
    if sheet.conditionalFormattingRules:
        for rule in sorted(sheet.conditionalFormattingRules, key=lambda r: r.priority):
            lines.append(_describe_cond_rule(rule))
    else:
        lines.append("- No conditional formatting rules.")
    lines.append("")
    lines.append("**Cell contents and formatting**")
    lines.append(
        "Recreate each cell below. Use the value or formula as given; apply number formats, "
        "fonts, fills, alignment, and merged flags as described."
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
        lines.append(f"- **{cref}**")
        if c.formula:
            lines.append(f"  - Enter formula: {c.formula}")
        elif c.value is not None:
            lines.append(f"  - Enter value: {c.value!r}")
        else:
            lines.append("  - Leave empty or clear contents (still apply formatting if any).")
        if c.displayValue and c.displayValue != str(c.value):
            lines.append(f"  - Expected display (if applicable): {c.displayValue!r}")
        lines.append(f"  - Number format: {c.numberFormat or 'General'}")
        lines.append(f"  - Font: {_describe_font(c.font)}.")
        lines.append(f"  - Fill: {c.fill or 'none'}.")
        lines.append(f"  - Alignment: {_describe_alignment(c.alignment)}.")
        if c.merged:
            lines.append("  - This cell participates in a merged range (see merged ranges above).")
        if c.hidden:
            lines.append("  - Row or column may be hidden.")
        lines.append("")
    return "\n".join(lines).rstrip()


def export_prompt(output: WorkbookOutput) -> str:
    wb = output.workbook
    parts: list[str] = []
    parts.append("# Instructions to recreate the spreadsheet")
    parts.append("")
    parts.append(
        f"Build a spreadsheet equivalent to the source “{wb.sourceName}” "
        f"(origin: {wb.sourceType}). Follow the structure sheet by sheet."
    )
    parts.append("")
    parts.append("## Workbook overview")
    parts.append(f"- Total sheets: {len(wb.sheets)}")
    for s in sorted(wb.sheets, key=lambda x: x.index):
        parts.append(f"- Sheet {s.index + 1}: “{s.name}” ({'hidden' if s.hidden else 'visible'})")
    parts.append("")
    for s in sorted(wb.sheets, key=lambda x: x.index):
        parts.append(_sheet_prompt(s))
        parts.append("")
    parts.append(
        "## Final checks\n"
        "- Verify used ranges, merged regions, and conditional formats match the descriptions.\n"
        "- Spot-check formulas and displayed values against the source."
    )
    return "\n".join(parts).rstrip() + "\n"
