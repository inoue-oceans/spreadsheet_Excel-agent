from src.models.schema import WorkbookOutput


def export_json(output: WorkbookOutput) -> str:
    """Serialize workbook output as JSON text (UTF-8 code points; write to files with encoding utf-8)."""
    return output.model_dump_json(indent=2, exclude_none=False)
