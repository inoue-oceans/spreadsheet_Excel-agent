"""Color conversion utilities for Excel and Google Sheets."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.styles.colors import Color
    from openpyxl.workbook import Workbook

# Standard indexed colour palette (indices 0-63) used by Excel.
# Subset of the most common ones; full palette from ECMA-376.
_INDEXED_COLORS: list[str] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
]

# Default theme colours (Office theme) before tint is applied.
_DEFAULT_THEME_COLORS: list[str] = [
    "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
    "A5A5A5", "FFC000", "5B9BD5", "70AD47",
]


def _apply_tint(rgb_hex: str, tint: float) -> str:
    """Apply an Excel tint value (-1..+1) to an RGB hex string."""
    r = int(rgb_hex[0:2], 16)
    g = int(rgb_hex[2:4], 16)
    b = int(rgb_hex[4:6], 16)

    def _tint_channel(c: int) -> int:
        if tint < 0:
            return int(c * (1.0 + tint))
        return int(c + (255 - c) * tint)

    r2 = max(0, min(255, _tint_channel(r)))
    g2 = max(0, min(255, _tint_channel(g)))
    b2 = max(0, min(255, _tint_channel(b)))
    return f"{r2:02X}{g2:02X}{b2:02X}"


def theme_color_to_hex(
    theme_id: int,
    tint: float = 0.0,
    theme_colors: list[str] | None = None,
) -> str | None:
    """Resolve a theme color index + tint to #RRGGBB."""
    palette = theme_colors or _DEFAULT_THEME_COLORS
    if theme_id < 0 or theme_id >= len(palette):
        return None
    base = palette[theme_id]
    if tint:
        base = _apply_tint(base, tint)
    return f"#{base.upper()}"


def indexed_color_to_hex(index: int) -> str | None:
    """Resolve an indexed color to #RRGGBB."""
    if index == 64:
        return None  # system foreground / automatic
    if 0 <= index < len(_INDEXED_COLORS):
        return f"#{_INDEXED_COLORS[index].upper()}"
    return None


def argb_to_hex(argb_str: str) -> str | None:
    """Convert ARGB string (e.g. 'FF336699') to #RRGGBB, stripping alpha."""
    if not argb_str:
        return None
    raw = argb_str.strip().lstrip("#")
    if len(raw) == 8 and re.fullmatch(r"[0-9A-Fa-f]{8}", raw):
        return f"#{raw[2:].upper()}"
    if len(raw) == 6 and re.fullmatch(r"[0-9A-Fa-f]{6}", raw):
        return f"#{raw.upper()}"
    return None


def rgba_float_to_hex(
    red: float = 0.0,
    green: float = 0.0,
    blue: float = 0.0,
    alpha: float = 1.0,
) -> str:
    """Convert Google Sheets 0.0-1.0 RGBA floats to #RRGGBB."""
    r = max(0, min(255, int(round(red * 255))))
    g = max(0, min(255, int(round(green * 255))))
    b = max(0, min(255, int(round(blue * 255))))
    return f"#{r:02X}{g:02X}{b:02X}"


def gsheet_color_to_hex(color_dict: dict | None) -> str | None:
    """Convert a Google Sheets color object {red, green, blue, alpha} to HEX."""
    if not color_dict:
        return None
    return rgba_float_to_hex(
        red=color_dict.get("red", 0.0),
        green=color_dict.get("green", 0.0),
        blue=color_dict.get("blue", 0.0),
        alpha=color_dict.get("alpha", 1.0),
    )


def _extract_theme_colors(wb: "Workbook") -> list[str]:
    """Try to extract theme colours from an openpyxl Workbook."""
    try:
        theme = wb.theme
        if theme and hasattr(theme, "themeElements"):
            # openpyxl doesn't expose theme colours cleanly;
            # fall back to defaults if parsing fails.
            pass
    except Exception:
        pass
    return list(_DEFAULT_THEME_COLORS)


def openpyxl_color_to_hex(color: "Color | None", wb: "Workbook | None" = None) -> str | None:
    """Convert any openpyxl Color object to #RRGGBB."""
    if color is None:
        return None

    # type == 'rgb' (most common)
    if color.type == "rgb" and color.rgb and color.rgb != "00000000":
        return argb_to_hex(str(color.rgb))

    # type == 'theme'
    if color.type == "theme" and color.theme is not None:
        theme_colors = _extract_theme_colors(wb) if wb else list(_DEFAULT_THEME_COLORS)
        tint = color.tint or 0.0
        return theme_color_to_hex(color.theme, tint, theme_colors)

    # type == 'indexed'
    if color.type == "indexed" and color.indexed is not None:
        return indexed_color_to_hex(color.indexed)

    # Fallback: try raw rgb string
    if hasattr(color, "rgb") and color.rgb:
        return argb_to_hex(str(color.rgb))

    return None
