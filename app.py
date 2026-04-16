"""Streamlit UI for Spreadsheet Agent - setting table generator."""

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="Spreadsheet Setting Table Generator", layout="wide")
st.title("Excel / Google Sheets 設定表生成ツール")
st.caption("Excel や Google スプレッドシートを解析し、AI に渡せる設定表を自動生成します")

# --- 使い方ガイド ---
with st.expander("使い方ガイド", expanded=False):
    st.markdown("""
### Excel ファイルの場合

1. 左サイドバーで **「Excel Upload」** を選択
2. `.xlsx` ファイルをドラッグ＆ドロップ、または「Browse files」でアップロード
3. **「解析する」** ボタンをクリック
4. 結果が表示されたら、各タブ（JSON / Markdown / Prompt）でダウンロード

### Google スプレッドシートの場合

#### 事前準備（初回のみ）

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. プロジェクトを作成（または既存プロジェクトを選択）
3. 「APIとサービス」→「ライブラリ」で **Google Sheets API** を検索し、有効化
4. 「APIとサービス」→「認証情報」→「認証情報を作成」→「OAuth クライアント ID」を選択
5. アプリケーションの種類で「デスクトップ アプリ」を選択し作成
6. 作成された認証情報の JSON をダウンロード
7. ダウンロードしたファイルを `credentials.json` にリネームし、本ツールと同じフォルダに配置
8. 「APIとサービス」→「OAuth 同意画面」でテストユーザーに自分のメールアドレスを追加

#### 解析手順

1. 左サイドバーで **「Google Sheets URL」** を選択
2. **「Google 認証」** ボタンをクリック（ブラウザで Google アカウント認証画面が開きます）
3. 認証完了後、対象スプレッドシートの URL を入力
4. **「解析する」** ボタンをクリック
5. 結果が表示されたら、各タブでダウンロード

### 出力形式の説明

| 形式 | 用途 |
|------|------|
| **JSON** | AI エージェントへの入力用。セル単位の全情報（値・数式・書式・条件付き書式など）を構造化データとして出力 |
| **Markdown** | 人間が内容を確認・レビューする用。表形式で見やすく整理 |
| **Prompt** | AI に「この表を再現して」と指示する用。自然言語でシート構造・書式・数式を記述 |

### 対応していない要素（現在のバージョン）

以下の要素は解析対象外です：

- グラフ・チャート
- 画像・図形・SmartArt
- コメント・メモ
- ピボットテーブル
- マクロ（VBA）
- フィルタビューの完全再現
- グループ化（折りたたみ）の完全再現

### トラブルシューティング

| エラー | 原因と対処 |
|--------|-----------|
| **「Unsupported file type」** | `.xlsx` 以外のファイル（`.xls`, `.csv` 等）がアップロードされています。`.xlsx` 形式で保存し直してください |
| **「Failed to read workbook」** | ファイルが壊れているか、パスワード保護されています。保護を解除して再度お試しください |
| **「Cannot extract spreadsheet ID」** | URL の形式が正しくありません。Google スプレッドシートの共有 URL をそのまま貼り付けてください |
| **「Authentication failed」** | `credentials.json` が見つからないか、OAuth 設定に問題があります。事前準備の手順を確認してください |
| **「Permission denied」** | 対象のスプレッドシートへのアクセス権がありません。スプレッドシートの共有設定を確認してください |
| **解析が遅い** | セル数が多いファイルは時間がかかります。不要なシートや空白領域が広い場合は、範囲を整理すると改善します |
""")

# --- Sidebar: input method selection ---
input_method = st.sidebar.radio("入力方式", ["Excel Upload", "Google Sheets URL"])

# --- Session state for Google credentials ---
if "google_creds" not in st.session_state:
    st.session_state.google_creds = None

# --- Input area ---
uploaded_file = None
sheet_url = None

if input_method == "Excel Upload":
    uploaded_file = st.file_uploader(".xlsx ファイルをアップロード", type=["xlsx"])
else:
    sheet_url = st.text_input("Google Sheets URL", placeholder="https://docs.google.com/spreadsheets/d/...")

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("Google 認証"):
            try:
                from src.auth.google_auth import get_credentials
                creds = get_credentials()
                st.session_state.google_creds = creds
                st.success("認証に成功しました")
            except Exception as e:
                st.error(f"認証に失敗しました: {e}")

    with col2:
        if st.session_state.google_creds:
            st.info("認証済み")
        else:
            st.warning("未認証")

# --- Analyze button ---
st.divider()

if st.button("解析する", type="primary", use_container_width=True):
    # Validation
    if input_method == "Excel Upload" and uploaded_file is None:
        st.error(".xlsx ファイルをアップロードしてください。")
        st.stop()
    if input_method == "Google Sheets URL" and not sheet_url:
        st.error("Google スプレッドシートの URL を入力してください。")
        st.stop()
    if input_method == "Google Sheets URL" and not st.session_state.google_creds:
        st.error("先に Google 認証を行ってください。")
        st.stop()

    with st.spinner("解析中..."):
        try:
            from src.exceptions import SpreadsheetAgentError

            # Parse
            if input_method == "Excel Upload":
                from src.parsers.excel_parser import parse_excel
                workbook_data = parse_excel(uploaded_file, source_name=uploaded_file.name)
            else:
                from src.parsers.gsheet_parser import parse_google_sheet
                workbook_data = parse_google_sheet(sheet_url, st.session_state.google_creds)

            # Normalize
            from src.normalizers.workbook_normalizer import normalize
            output = normalize(workbook_data)

            # Export
            from src.exporters.json_exporter import export_json
            from src.exporters.markdown_exporter import export_markdown
            from src.exporters.prompt_exporter import export_prompt

            json_str = export_json(output)
            md_str = export_markdown(output)
            prompt_str = export_prompt(output)

            st.success("解析が完了しました！")

            # --- Results in tabs ---
            tab_json, tab_md, tab_prompt = st.tabs(["JSON", "Markdown", "Prompt"])

            with tab_json:
                st.code(json_str, language="json")
                st.download_button(
                    "JSON をダウンロード",
                    data=json_str.encode("utf-8"),
                    file_name="setting_table.json",
                    mime="application/json",
                )

            with tab_md:
                st.code(md_str, language="markdown")
                st.download_button(
                    "Markdown をダウンロード",
                    data=md_str.encode("utf-8"),
                    file_name="setting_table.md",
                    mime="text/markdown",
                )

            with tab_prompt:
                st.code(prompt_str, language="markdown")
                st.download_button(
                    "Prompt をダウンロード",
                    data=prompt_str.encode("utf-8"),
                    file_name="recreation_prompt.txt",
                    mime="text/plain",
                )

        except SpreadsheetAgentError as e:
            st.error(f"エラー: {e}")
        except Exception as e:
            st.error(f"予期しないエラー: {e}")
            raise
