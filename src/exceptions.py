"""Custom exceptions for the Spreadsheet Agent."""


class SpreadsheetAgentError(Exception):
    """Base exception."""


class UnsupportedFileTypeError(SpreadsheetAgentError):
    """Raised when the uploaded file is not a supported format (.xlsx)."""


class InvalidSpreadsheetUrlError(SpreadsheetAgentError):
    """Raised when the Google Sheets URL is malformed or cannot be parsed."""


class GoogleAuthError(SpreadsheetAgentError):
    """Raised when Google OAuth authentication fails."""


class PermissionDeniedError(SpreadsheetAgentError):
    """Raised when the user lacks permission to access the spreadsheet."""


class WorkbookReadError(SpreadsheetAgentError):
    """Raised when the workbook/spreadsheet cannot be read or is corrupted."""


class ParseError(SpreadsheetAgentError):
    """Raised when parsing a cell or element fails unexpectedly."""
