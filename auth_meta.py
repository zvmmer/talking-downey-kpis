"""Talking Downey — Meta (Facebook + Instagram) auth, one-time setup.

Run this ONCE after completing SETUP_META.md. It takes your short-lived user
token, exchanges it for a 60-day long-lived token, then derives per-page tokens
for every FB page you admin (and detects IG accounts linked to those pages).

Usage
-----
    cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
    MCP/.venv/bin/python Projects/talking_downey/kpis/auth_meta.py
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from getpass import getpass
from pathlib import Path

import urllib.parse
import urllib.request

HERE = Path(__file__).resolve().parent
CRED_DIR = HERE / ".credentials"
TOKEN_PATH = CRED_DIR / "meta_token.json"
GRAPH = "https://graph.facebook.com/v21.0"


def api(path: str, params: dict) -> dict:
    """GET a Meta Graph API endpoint and return the parsed JSON, raising on error."""
    qs = urllib.parse.urlencode(params)
    url = f"{GRAPH}/{path.lstrip('/')}?{qs}"
    try:
        with urllib.request.urlopen(url, timeout=20) as resp:
            data = json.load(resp)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        sys.exit(f"\n✗ Meta API error ({e.code}): {body}\n")
    if "error" in data:
        sys.exit(f"\n✗ Meta API error: {data['error']}\n")
    return data


def exchange_short_to_long(app_id: str, app_secret: str, short_token: str) -> tuple[str, int]:
    """Exchange a short-lived user token for a long-lived one (~60 days)."""
    data = api("oauth/access_token", {
        "grant_type": "fb_exchange_token",
        "client_id": app_id,
        "client_secret": app_secret,
        "fb_exchange_token": short_token,
    })
    return data["access_token"], int(data.get("expires_in", 0))


def list_pages(user_token: str) -> list[dict]:
    """Return [{id, name, access_token, ig_user_id|None, ig_username|None}, ...]."""
    pages = api("me/accounts", {
        "access_token": user_token,
        "fields": "id,name,access_token,instagram_business_account",
    }).get("data", [])
    enriched = []
    for p in pages:
        ig_id = None
        ig_user = None
        ig_block = p.get("instagram_business_account")
        if ig_block and ig_block.get("id"):
            ig_id = ig_block["id"]
            # Fetch username for display
            try:
                ig_info = api(ig_id, {
                    "access_token": p["access_token"],
                    "fields": "username",
                })
                ig_user = ig_info.get("username")
            except SystemExit:
                pass
        enriched.append({
            "page_id": p["id"],
            "page_name": p["name"],
            "page_access_token": p["access_token"],
            "ig_user_id": ig_id,
            "ig_username": ig_user,
        })
    return enriched


def main() -> None:
    CRED_DIR.mkdir(parents=True, exist_ok=True)
    print("Meta auth — one-time setup.\n")
    print("You'll need three values from the SETUP_META.md flow:")
    print("  • App ID  (Meta Dashboard → App Settings → Basic)")
    print("  • App Secret  (same screen, click 'Show')")
    print("  • Short-lived User Access Token  (Graph API Explorer)\n")

    app_id = input("App ID: ").strip()
    app_secret = getpass("App Secret (hidden): ").strip()
    short = getpass("Short-lived User Token (hidden): ").strip()

    if not (app_id and app_secret and short):
        sys.exit("All three values required.")

    print("\nExchanging short-lived → long-lived user token…")
    long_token, expires_in = exchange_short_to_long(app_id, app_secret, short)
    print(f"  ✓ Long-lived token (expires in {expires_in // 86400} days).")

    print("\nFetching pages you admin…")
    pages = list_pages(long_token)
    if not pages:
        sys.exit("No pages found. Are you an admin of any FB pages?")

    print(f"\nPages found:")
    for p in pages:
        ig = f"IG: @{p['ig_username']}" if p["ig_username"] else "IG: (not linked)"
        print(f"  ✓ {p['page_name']:<30} | FB page | {ig}")

    state = {
        "saved_at": datetime.now(timezone.utc).isoformat(),
        "app_id": app_id,
        # App Secret stored alongside — needed for re-exchange if token rotated.
        "app_secret": app_secret,
        "user_access_token": long_token,
        "user_token_expires_in_seconds_at_save": expires_in,
        "pages": pages,
    }
    TOKEN_PATH.write_text(json.dumps(state, indent=2))
    print(f"\n✓ Saved: {TOKEN_PATH}")
    print(f"  Tokens are read-only. App Secret is in this file too — .gitignore covers it.")
    print(f"\nNext: MCP/.venv/bin/python Projects/talking_downey/kpis/pull_meta.py")


if __name__ == "__main__":
    main()
