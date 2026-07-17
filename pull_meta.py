"""Talking Downey — pull Meta (FB pages + Instagram) KPIs.

Pulls insights for the most recent posts on every FB page you admin, plus
recent IG media for every IG account linked to those pages.

Usage
-----
    # Most recent post per page + most recent media per IG account:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_meta.py

    # Last N items:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_meta.py --last 5
"""
from __future__ import annotations

import argparse
import json
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
CRED_DIR = HERE / ".credentials"
TOKEN_PATH = CRED_DIR / "meta_token.json"
GRAPH = "https://graph.facebook.com/v21.0"


def load_state() -> dict:
    if not TOKEN_PATH.exists():
        sys.exit("Run auth_meta.py first.")
    return json.loads(TOKEN_PATH.read_text())


def api(path: str, params: dict) -> dict:
    qs = urllib.parse.urlencode(params)
    url = f"{GRAPH}/{path.lstrip('/')}?{qs}"
    try:
        with urllib.request.urlopen(url, timeout=20) as resp:
            return json.load(resp)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        return {"error": {"http_status": e.code, "body": body}}


def page_posts(page_id: str, token: str, n: int) -> list[dict]:
    """Return the N most-recent posts on a FB page with engagement counts."""
    data = api(f"{page_id}/posts", {
        "access_token": token,
        "limit": min(n, 25),
        "fields": "id,message,created_time,permalink_url,"
                  "reactions.summary(true).limit(0),"
                  "comments.summary(true).limit(0),"
                  "shares",
    }).get("data", [])
    out = []
    for p in data:
        out.append({
            "post_id": p["id"],
            "message": (p.get("message") or "")[:80],
            "created_time": p.get("created_time"),
            "permalink": p.get("permalink_url"),
            "reactions": p.get("reactions", {}).get("summary", {}).get("total_count", 0),
            "comments": p.get("comments", {}).get("summary", {}).get("total_count", 0),
            "shares": (p.get("shares") or {}).get("count", 0),
        })
    return out


def page_post_insights(post_id: str, token: str) -> dict:
    """Pull reach + impressions for a single FB post."""
    metrics = "post_impressions,post_impressions_unique,post_clicks,post_video_views"
    data = api(f"{post_id}/insights", {
        "access_token": token,
        "metric": metrics,
    })
    if "error" in data:
        return {"error": data["error"]}
    out = {}
    for item in data.get("data", []):
        vals = item.get("values", [])
        out[item["name"]] = vals[0].get("value") if vals else None
    return out


def ig_media(ig_user_id: str, token: str, n: int) -> list[dict]:
    """Return N most-recent IG media items + their insights."""
    media = api(f"{ig_user_id}/media", {
        "access_token": token,
        "limit": min(n, 25),
        "fields": "id,caption,media_type,media_product_type,permalink,timestamp,"
                  "like_count,comments_count",
    }).get("data", [])
    out = []
    for m in media:
        # Insights metrics vary by media_type. We request the safe-for-all set
        # and let unsupported ones error gracefully.
        is_reel = m.get("media_product_type") == "REELS"
        if is_reel:
            metrics = "reach,plays,total_interactions,saved,shares"
        elif m.get("media_type") == "VIDEO":
            metrics = "reach,impressions,saved,shares,plays"
        else:  # IMAGE / CAROUSEL_ALBUM
            metrics = "reach,impressions,saved,shares"
        insights = api(f"{m['id']}/insights", {
            "access_token": token,
            "metric": metrics,
        })
        ins = {}
        if "error" not in insights:
            for item in insights.get("data", []):
                vals = item.get("values", [])
                ins[item["name"]] = vals[0].get("value") if vals else None
        out.append({
            "media_id": m["id"],
            "media_type": m.get("media_type"),
            "media_product_type": m.get("media_product_type"),
            "caption": (m.get("caption") or "")[:80],
            "permalink": m.get("permalink"),
            "timestamp": m.get("timestamp"),
            "likes": m.get("like_count", 0),
            "comments": m.get("comments_count", 0),
            "insights": ins,
        })
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--last", type=int, default=1, help="Items per page/IG account")
    args = ap.parse_args()

    state = load_state()

    print(f"Meta KPI pull · {datetime.now(timezone.utc).isoformat()}")
    print(f"Token saved: {state.get('saved_at')}\n")

    for page in state["pages"]:
        ptok = page["page_access_token"]
        print(f"=== FB: {page['page_name']} ===")
        posts = page_posts(page["page_id"], ptok, args.last)
        for p in posts:
            ins = page_post_insights(p["post_id"], ptok)
            print(f"  [{p['created_time']}] {p['message']}…")
            print(f"    reactions: {p['reactions']:>5}  "
                  f"comments: {p['comments']:>4}  "
                  f"shares: {p['shares']:>3}")
            if "error" not in ins:
                print(f"    impressions: {ins.get('post_impressions', 0):>6}  "
                      f"unique: {ins.get('post_impressions_unique', 0):>6}  "
                      f"clicks: {ins.get('post_clicks', 0):>5}  "
                      f"video_views: {ins.get('post_video_views', 0):>6}")

        if page["ig_user_id"]:
            print(f"\n=== IG: @{page['ig_username']} ===")
            media = ig_media(page["ig_user_id"], ptok, args.last)
            for m in media:
                kind = m["media_product_type"] or m["media_type"]
                print(f"  [{m['timestamp']}] {kind:<8} {m['caption']}…")
                ins = m["insights"]
                print(f"    likes: {m['likes']:>5}  comments: {m['comments']:>4}  "
                      f"reach: {ins.get('reach', 0):>6}  plays: {ins.get('plays', 0):>6}  "
                      f"saved: {ins.get('saved', 0):>4}  shares: {ins.get('shares', 0):>4}")
        print()

    print("✓ Meta pull complete.")


if __name__ == "__main__":
    main()
