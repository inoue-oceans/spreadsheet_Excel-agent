"""Streamlit UI for Excel/Google Sheets Setting Table Generator."""

from __future__ import annotations

import json

import streamlit as st
from dotenv import load_dotenv

from src.auth import google_auth
from src.exceptions import (
    GoogleAuthError,
    InvalidSpreadsheetUrlError,
    ParseError,
    PermissionDeniedError,
    SpreadsheetAgentError,
    UnsupportedFileTypeError,
    WorkbookReadError,
)
from src.exporters.json_exporter import export_json
from src.exporters.markdown_exporter import export_markdown
from src.exporters.prompt_exporter import export_prompt
from src.models.schema import WorkbookOutput
from src.normalizers.workbook_normalizer import normalize
from src.parsers.excel_parser import parse_excel
from src.parsers.gsheet_parser import parse_google_sheet

load_dotenv()

st.set_page_config(page_title="Excel/Google Sheets Setting Table Generator", layout="wide")

# --- 認証ガード ---
ALLOWED_DOMAIN = "@oceans-web.co.jp"

if not st.user.is_logged_in:
    st.title("Excel/Google スプレッドシート設定表ジェネレーター")
    st.info("このアプリは社員限定です。Googleアカウントでログインしてください。")
    if st.button("Googleでログイン", type="primary"):
        st.login("google")
    st.stop()

# ドメインチェック
if not st.user.email.endswith(ALLOWED_DOMAIN):
    st.error(
        f"このアプリは oceans-web.co.jp ドメインの社員のみ利用可能です。現在: {st.user.email}"
    )
    if st.button("ログアウト"):
        st.logout()
    st.stop()

st.title("Excel/Google Sheets Setting Table Generator")

if "google_creds" not in st.session_state:
    st.session_state.google_creds = None
if "analysis_outputs" not in st.session_state:
    st.session_state.analysis_outputs = None


def _friendly_error_message(exc: SpreadsheetAgentError) -> str:
    if isinstance(exc, UnsupportedFileTypeError):
        return (
            "Unsupported file type. Please upload a single .xlsx file "
            "(Excel workbook in Office Open XML format)."
        )
    if isinstance(exc, InvalidSpreadsheetUrlError):
        return (
            "The Google Sheets URL is invalid or could not be parsed. "
            "Paste the full URL from your browser (e.g. …/spreadsheets/d/…/edit…)."
        )
    if isinstance(exc, GoogleAuthError):
        return (
            "Google authentication failed. "
            "Set GOOGLE_CREDENTIALS_PATH (and optionally GOOGLE_TOKEN_PATH) in `.env`, "
            "or place `credentials.json` in the working directory, then try again."
        )
    if isinstance(exc, PermissionDeniedError):
        return (
            "Permission denied for this spreadsheet. "
            "Ask the owner to share it with your Google account or use a URL you can open in the browser."
        )
    if isinstance(exc, WorkbookReadError):
        return (
            "Could not read the workbook or spreadsheet. "
            "The file may be corrupted, password-protected, or temporarily unavailable."
        )
    if isinstance(exc, ParseError):
        return (
            "Parsing failed while reading cells or sheet structure. "
            "If the file is very large or unusual, try a smaller range or a copy saved as .xlsx."
        )
    return (
        "Something went wrong while processing the spreadsheet. "
        f"Details: {exc!s}"
    )


input_method = st.sidebar.radio("Source", ["Excel Upload", "Google Sheets URL"])

detect_hidden = st.sidebar.checkbox(
    "表示/非表示情報を検出する",
    value=True,
    help="OFFにすると、シート・行・列・セルの非表示状態を無視して出力します",
)

st.sidebar.divider()
st.sidebar.caption(f"ログイン中: {st.user.email}")
if st.sidebar.button("ログアウト"):
    st.logout()

uploaded_file = None
sheet_url = None

if input_method == "Excel Upload":
    uploaded_file = st.file_uploader("Upload .xlsx", type=["xlsx"])
else:
    sheet_url = st.text_input(
        "Google Sheets URL",
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    try:
        creds = google_auth.get_credentials_streamlit(st)
        if creds:
            st.session_state.google_creds = creds
            st.success("Google認証済み。スプレッドシートを分析できます。")
        else:
            st.info("上の「Google Sheets に接続」ボタンから認証してください。")
    except GoogleAuthError as e:
        st.error(_friendly_error_message(e))
    except Exception as e:
        st.error(f"認証処理でエラーが発生しました: {e!s}")

col_analyze, col_json, col_md, col_prompt, _col_spacer = st.columns(
    [1, 2, 2.5, 2.5, 4]
)
with col_analyze:
    analyze_clicked = st.button("Analyze", type="primary", use_container_width=True)

outs = st.session_state.analysis_outputs
has_outputs = outs is not None

with col_json:
    st.download_button(
        label="JSONダウンロード",
        data=outs["json"].encode("utf-8") if has_outputs else b"",
        file_name="setting_table.json",
        mime="application/json",
        disabled=not has_outputs,
        use_container_width=True,
    )
with col_md:
    st.download_button(
        label="Markdownダウンロード",
        data=outs["markdown"].encode("utf-8") if has_outputs else b"",
        file_name="setting_table.md",
        mime="text/markdown",
        disabled=not has_outputs,
        use_container_width=True,
    )
with col_prompt:
    st.download_button(
        label="Promptダウンロード",
        data=outs["prompt"].encode("utf-8") if has_outputs else b"",
        file_name="setting_table_prompt.txt",
        mime="text/plain",
        disabled=not has_outputs,
        use_container_width=True,
    )

if st.session_state.pop("show_analysis_success", False):
    st.success("Analysis complete.")

if analyze_clicked:
    st.session_state.analysis_outputs = None

    if input_method == "Excel Upload" and uploaded_file is None:
        st.error("Please upload a .xlsx file.")
    elif input_method == "Google Sheets URL" and not (sheet_url or "").strip():
        st.error("Please enter a Google Sheets URL.")
    elif input_method == "Google Sheets URL" and not st.session_state.google_creds:
        st.error("Please authenticate with Google first.")
    else:
        with st.spinner("Analyzing workbook…"):
            try:
                if input_method == "Excel Upload":
                    assert uploaded_file is not None
                    workbook_data = parse_excel(uploaded_file, source_name=uploaded_file.name)
                else:
                    workbook_data = parse_google_sheet(sheet_url.strip(), st.session_state.google_creds)

                output = WorkbookOutput(workbook=normalize(workbook_data, detect_hidden=detect_hidden))
                json_str = export_json(output)
                md_str = export_markdown(output)
                prompt_str = export_prompt(output)

                st.session_state.analysis_outputs = {
                    "json": json_str,
                    "markdown": md_str,
                    "prompt": prompt_str,
                }
                st.session_state.show_analysis_success = True
                st.rerun()
            except UnsupportedFileTypeError as e:
                st.error(_friendly_error_message(e))
            except InvalidSpreadsheetUrlError as e:
                st.error(_friendly_error_message(e))
            except GoogleAuthError as e:
                st.error(_friendly_error_message(e))
            except PermissionDeniedError as e:
                st.error(_friendly_error_message(e))
            except WorkbookReadError as e:
                st.error(_friendly_error_message(e))
            except ParseError as e:
                st.error(_friendly_error_message(e))
            except SpreadsheetAgentError as e:
                st.error(_friendly_error_message(e))
            except Exception as e:
                st.error(f"Unexpected error: {e!s}")
                raise

if outs:
    tab_json, tab_md, tab_prompt = st.tabs(["JSON", "Markdown", "Prompt"])

    with tab_json:
        try:
            parsed = json.loads(outs["json"])
            st.code(json.dumps(parsed, indent=2, ensure_ascii=False), language="json")
        except json.JSONDecodeError:
            st.code(outs["json"], language="json")

    with tab_md:
        st.markdown(outs["markdown"])

    with tab_prompt:
        st.code(outs["prompt"], language="text")
