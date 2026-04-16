"""Google OAuth authentication for Sheets API access."""

from __future__ import annotations

import os
from pathlib import Path

from src.exceptions import GoogleAuthError
from src.utils.logger import get_logger

logger = get_logger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_credentials(
    credentials_path: str | Path | None = None,
    token_path: str | Path | None = None,
    scopes: list[str] | None = None,
):
    """Obtain Google OAuth credentials.

    1. Check for existing token at token_path.
    2. If token is expired, refresh it.
    3. If no token, run InstalledAppFlow.

    Returns:
        google.oauth2.credentials.Credentials
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

    # Load existing token
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        except Exception:
            logger.warning("Failed to load existing token, will re-authenticate.")
            creds = None

    # Refresh or run new flow
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

        # Save token for next run (restrict file permissions)
        try:
            token_path.write_text(creds.to_json())
            try:
                import os
                os.chmod(str(token_path), 0o600)
            except OSError:
                pass  # Windows may not support chmod fully
            logger.info("Token saved to %s", token_path)
        except Exception as e:
            logger.warning("Could not save token: %s", e)

    return creds
