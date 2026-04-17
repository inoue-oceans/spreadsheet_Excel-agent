"""Google OAuth authentication for Sheets API access."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from src.exceptions import GoogleAuthError
from src.utils.logger import get_logger

logger = get_logger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_credentials(
    credentials_path: str | Path | None = None,
    token_path: str | Path | None = None,
    scopes: list[str] | None = None,
):
    """Local development用: InstalledAppFlow で localhost にリダイレクト。

    Streamlit Cloud では動作しないので、そちらでは get_credentials_streamlit を使う。
    """
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError as e:
        raise GoogleAuthError(
            "Google auth libraries not installed. "
            "Run: pip install google-auth google-auth-oauthlib google-api-python-client"
        ) from e

    scopes = scopes or SCOPES
    credentials_path = Path(credentials_path or os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json"))
    token_path = Path(token_path or os.getenv("GOOGLE_TOKEN_PATH", "token.json"))

    creds = None

    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        except Exception:
            logger.warning("Failed to load existing token, will re-authenticate.")
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            logger.warning("Token refresh failed: %s. Re-authenticating.", e)
            creds = None

    if not creds or not creds.valid:
        if not credentials_path.exists():
            raise GoogleAuthError(
                f"Credentials file not found: {credentials_path}. "
                "Download it from Google Cloud Console."
            )
        try:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), scopes)
            creds = flow.run_local_server(port=0)
        except Exception as e:
            raise GoogleAuthError(f"Authentication failed: {e}") from e

        try:
            token_path.write_text(creds.to_json())
            try:
                os.chmod(str(token_path), 0o600)
            except OSError:
                pass
            logger.info("Token saved to %s", token_path)
        except Exception as e:
            logger.warning("Could not save token: %s", e)

    return creds


def get_credentials_streamlit(st_module: Any):
    """Streamlit Cloud 用: Web OAuth Flow で認証。

    1. session_state にキャッシュ済み creds があれば返す（必要なら refresh）
    2. URL に ?code= があればトークン交換して session_state に保存、st.rerun()
    3. 未認証ならボタンを表示して None を返す
    """
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import Flow
    except ImportError as e:
        raise GoogleAuthError(
            "Google auth libraries not installed. "
            "Run: pip install google-auth google-auth-oauthlib google-api-python-client"
        ) from e

    cached = st_module.session_state.get("sheets_creds")
    if cached:
        if cached.valid:
            return cached
        if cached.expired and cached.refresh_token:
            try:
                cached.refresh(Request())
                st_module.session_state["sheets_creds"] = cached
                return cached
            except Exception as e:
                logger.warning("Token refresh failed: %s", e)
                st_module.session_state.pop("sheets_creds", None)

    try:
        oauth_cfg = st_module.secrets["google_sheets_oauth"]
        client_id = oauth_cfg["client_id"]
        client_secret = oauth_cfg["client_secret"]
        redirect_uri = oauth_cfg["redirect_uri"]
    except (KeyError, FileNotFoundError) as e:
        raise GoogleAuthError(
            "google_sheets_oauth 設定が secrets にありません。"
            "Streamlit Cloud の Secrets 設定を確認してください。"
        ) from e

    client_config = {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }
    flow = Flow.from_client_config(client_config, scopes=SCOPES)
    flow.redirect_uri = redirect_uri

    code = st_module.query_params.get("code")
    if code:
        try:
            flow.fetch_token(code=code)
            creds: Credentials = flow.credentials
            st_module.session_state["sheets_creds"] = creds
            st_module.query_params.clear()
            st_module.rerun()
        except Exception as e:
            st_module.query_params.clear()
            raise GoogleAuthError(f"トークン取得に失敗しました: {e}") from e

    auth_url, _state = flow.authorization_url(
        access_type="offline",
        prompt="consent",
        include_granted_scopes="true",
    )
    st_module.link_button("Google Sheets に接続", auth_url, type="primary")
    return None
