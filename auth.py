"""Talking Downey — YouTube OAuth, one-time setup.

Run this ONCE after completing SETUP_GOOGLE_CLOUD.md. It opens a browser,
walks you through Google's consent flow, and saves a refresh token to
`.credentials/token.json`. After that, every other KPI script uses the saved
token automatically; no further prompts.

Usage
-----
    cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
    MCP/.venv/bin/python Projects/talking_downey/kpis/auth.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except ImportError as exc:
    sys.exit(
        f"Missing dep: {exc}\n"
        f"Install: MCP/.venv/bin/python -m pip install -r "
        f"Projects/talking_downey/kpis/requirements.txt"
    )

HERE = Path(__file__).resolve().parent
CRED_DIR = HERE / ".credentials"
CLIENT_SECRETS = CRED_DIR / "client_secrets.json"
TOKEN_PATH = CRED_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/youtube.readonly",
    "https://www.googleapis.com/auth/yt-analytics.readonly",
]


def load_or_run_flow() -> Credentials:
    """Return credentials. If a refresh token exists, use it; else run the
    one-time browser consent flow."""
    creds: Credentials | None = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        print("Existing token expired — refreshing…")
        creds.refresh(Request())
        TOKEN_PATH.write_text(creds.to_json())
        print("  ✓ Refreshed.")
        return creds

    # First-time flow
    if not CLIENT_SECRETS.exists():
        sys.exit(
            f"\nclient_secrets.json not found at:\n  {CLIENT_SECRETS}\n\n"
            f"Follow SETUP_GOOGLE_CLOUD.md to download it from Google Cloud Console."
        )

    CRED_DIR.mkdir(parents=True, exist_ok=True)
    flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRETS), SCOPES)
    # `run_local_server` opens a browser, listens on a random localhost port,
    # captures the redirect, exchanges the code for tokens.
    print("Opening browser for Google auth…")
    creds = flow.run_local_server(
        port=0,
        prompt="consent",
        # Important: ask for offline access so we get a refresh token, not just
        # a 1-hour access token. Without this, the script would re-prompt
        # roughly every hour.
        access_type="offline",
    )
    TOKEN_PATH.write_text(creds.to_json())
    return creds


def whoami(creds: Credentials) -> tuple[str, str]:
    """Return (channel_title, channel_id) for the authed account."""
    yt = build("youtube", "v3", credentials=creds)
    resp = yt.channels().list(part="snippet", mine=True).execute()
    items = resp.get("items") or []
    if not items:
        sys.exit("No channel found for this account. Are you signed in as the channel owner?")
    return items[0]["snippet"]["title"], items[0]["id"]


def main() -> None:
    creds = load_or_run_flow()
    title, channel_id = whoami(creds)
    print(f"\n✓ Token saved to: {TOKEN_PATH}")
    print(f"  Channel:    {title}")
    print(f"  Channel ID: {channel_id}")
    print(f"\nAll set. You can now run:")
    print(f"  MCP/.venv/bin/python Projects/talking_downey/kpis/pull_youtube.py")

    # Stash the channel ID alongside the token for convenience.
    meta = {"channel_id": channel_id, "channel_title": title}
    (CRED_DIR / "channel.json").write_text(json.dumps(meta, indent=2))


if __name__ == "__main__":
    main()
