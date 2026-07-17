"""Talking Downey — pull YouTube KPIs for a single video.

This is the first end-to-end test of the auth + API pipeline. Once it works,
we layer the per-episode workflow (24h / 72h / 1wk snapshots) on top.

Usage
-----
    # Quick test — pulls latest video stats:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_youtube.py

    # Specific video:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_youtube.py --video-id Xj1A2B3c4D5

    # Last N videos:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_youtube.py --last 5
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
except ImportError as exc:
    sys.exit(f"Missing dep: {exc}. Run auth.py first.")

HERE = Path(__file__).resolve().parent
CRED_DIR = HERE / ".credentials"
TOKEN_PATH = CRED_DIR / "token.json"
CHANNEL_PATH = CRED_DIR / "channel.json"

SCOPES = [
    "https://www.googleapis.com/auth/youtube.readonly",
    "https://www.googleapis.com/auth/yt-analytics.readonly",
]


def load_creds() -> Credentials:
    if not TOKEN_PATH.exists():
        sys.exit(
            "No saved token. Run auth.py first:\n"
            "  MCP/.venv/bin/python Projects/talking_downey/kpis/auth.py"
        )
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if creds.expired and creds.refresh_token:
        from google.auth.transport.requests import Request
        creds.refresh(Request())
        TOKEN_PATH.write_text(creds.to_json())
    return creds


def load_channel_id() -> str:
    if not CHANNEL_PATH.exists():
        sys.exit("Run auth.py first to record channel_id.")
    return json.loads(CHANNEL_PATH.read_text())["channel_id"]


def get_recent_videos(yt, channel_id: str, n: int) -> list[dict]:
    """Return up to `n` most-recent videos on the channel."""
    # Channel → uploads playlist ID
    ch = yt.channels().list(part="contentDetails", id=channel_id).execute()
    uploads = ch["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]
    items = yt.playlistItems().list(
        part="snippet,contentDetails",
        playlistId=uploads,
        maxResults=min(n, 50),
    ).execute()
    return [
        {
            "video_id": it["contentDetails"]["videoId"],
            "title": it["snippet"]["title"],
            "published_at": it["contentDetails"]["videoPublishedAt"],
        }
        for it in items.get("items", [])
    ]


def video_stats(yt, video_id: str) -> dict:
    """Lifetime totals from Data API (views, likes, comments, duration)."""
    resp = yt.videos().list(part="snippet,statistics,contentDetails", id=video_id).execute()
    items = resp.get("items") or []
    if not items:
        return {"video_id": video_id, "error": "not found"}
    v = items[0]
    return {
        "video_id": video_id,
        "title": v["snippet"]["title"],
        "published_at": v["snippet"]["publishedAt"],
        "duration": v["contentDetails"]["duration"],
        "views": int(v["statistics"].get("viewCount", 0)),
        "likes": int(v["statistics"].get("likeCount", 0)),
        "comments": int(v["statistics"].get("commentCount", 0)),
    }


def video_analytics(yta, channel_id: str, video_id: str, start: str, end: str) -> dict:
    """Watch time, retention, traffic sources from Analytics API for a window."""
    base_kwargs = dict(
        ids=f"channel=={channel_id}",
        startDate=start,
        endDate=end,
        filters=f"video=={video_id}",
    )
    # Headline metrics
    headline = yta.reports().query(
        **base_kwargs,
        metrics="views,estimatedMinutesWatched,averageViewDuration,"
                "averageViewPercentage,subscribersGained",
    ).execute()
    # Traffic sources
    sources = yta.reports().query(
        **base_kwargs,
        metrics="views",
        dimensions="insightTrafficSourceType",
        sort="-views",
    ).execute()
    return {
        "window": {"start": start, "end": end},
        "headline": _rows_to_dicts(headline),
        "traffic_sources": _rows_to_dicts(sources),
    }


def _rows_to_dicts(resp: dict) -> list[dict]:
    cols = [c["name"] for c in resp.get("columnHeaders", [])]
    return [dict(zip(cols, row)) for row in resp.get("rows", [])]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--video-id", help="YouTube video ID. Default: most recent.")
    ap.add_argument("--last", type=int, default=1, help="Pull stats for the N most recent videos")
    ap.add_argument("--window-days", type=int, default=7,
                    help="Analytics window (default 7 days back from today)")
    args = ap.parse_args()

    creds = load_creds()
    channel_id = load_channel_id()
    yt = build("youtube", "v3", credentials=creds)
    yta = build("youtubeAnalytics", "v2", credentials=creds)

    if args.video_id:
        video_ids = [args.video_id]
    else:
        recent = get_recent_videos(yt, channel_id, args.last)
        video_ids = [v["video_id"] for v in recent]

    end = datetime.now(timezone.utc).date().isoformat()
    start = (datetime.now(timezone.utc).date() - timedelta(days=args.window_days)).isoformat()

    out = []
    for vid in video_ids:
        stats = video_stats(yt, vid)
        analytics = video_analytics(yta, channel_id, vid, start, end)
        record = {**stats, "analytics": analytics}
        out.append(record)
        # Pretty print one
        print(f"\n=== {stats.get('title', vid)} ===")
        print(f"  ID:       {vid}")
        print(f"  Posted:   {stats.get('published_at')}")
        print(f"  Views:    {stats.get('views'):,}")
        print(f"  Likes:    {stats.get('likes'):,}")
        print(f"  Comments: {stats.get('comments'):,}")
        if analytics["headline"]:
            h = analytics["headline"][0]
            print(f"  Last {args.window_days}d watch time: "
                  f"{h.get('estimatedMinutesWatched', 0):,.0f} min "
                  f"(avg view duration {h.get('averageViewDuration', 0):.0f}s, "
                  f"avg % viewed {h.get('averageViewPercentage', 0):.1f}%)")
        if analytics["traffic_sources"]:
            print(f"  Traffic sources (last {args.window_days}d):")
            for s in analytics["traffic_sources"][:5]:
                print(f"    {s['insightTrafficSourceType']:<24}  "
                      f"{s['views']:>6,} views")

    print(f"\n✓ Pulled {len(out)} video(s). All good — pipeline works.")


if __name__ == "__main__":
    main()
