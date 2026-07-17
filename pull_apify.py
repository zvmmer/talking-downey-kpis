"""Talking Downey — Apify-backed scraper for Instagram + Facebook.

Reads `apify_targets` from the episode manifest (profile URLs, not per-post),
calls Apify's hosted Actors, and merges the results into the same snapshot
JSON format that pull_public.py uses. report_public.py then combines
everything into the markdown report.

Why this script exists:
    Apify hosts maintained scrapers for IG + FB that work without your login
    and without Meta API auth. Costs pennies per snapshot. Used as a bridge
    while we wait for proper Meta Graph API access from the client.

Usage
-----
    # First check the token works:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_apify.py --check

    # Pull IG + FB into the same snapshot file as pull_public.py:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_apify.py \\
        --manifest Projects/talking_downey/kpis/episodes_public/EP47/manifest.yaml \\
        --tag t0
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    import yaml
    import ssl
    import urllib.request
    import urllib.parse
    import urllib.error
    import certifi
except ImportError as exc:
    sys.exit(f"Missing dep: {exc}")

# macOS framework Python 3.10 ships without bundled CA certs by default, so
# urllib.request throws CERTIFICATE_VERIFY_FAILED on HTTPS. Use certifi's bundle
# explicitly. (certifi is already installed as a googleapiclient transitive dep.)
_SSL_CTX = ssl.create_default_context(cafile=certifi.where())
_HTTPS_HANDLER = urllib.request.HTTPSHandler(context=_SSL_CTX)
_OPENER = urllib.request.build_opener(_HTTPS_HANDLER)
urllib.request.install_opener(_OPENER)

HERE = Path(__file__).resolve().parent
TOKEN_PATH = HERE / ".credentials" / "apify_token.txt"

# Apify Actor IDs (official maintained scrapers).
ACTOR_INSTAGRAM = "apify~instagram-scraper"
ACTOR_FACEBOOK = "apify~facebook-posts-scraper"
ACTOR_YOUTUBE = "streamers~youtube-scraper"
ACTOR_TIKTOK = "clockworks~tiktok-scraper"

API = "https://api.apify.com/v2"


def load_token() -> str:
    if not TOKEN_PATH.is_file():
        sys.exit(
            f"\nApify token not found at:\n  {TOKEN_PATH}\n\n"
            f"Follow SETUP_APIFY.md to create one + save it."
        )
    tok = TOKEN_PATH.read_text().strip()
    if not tok or tok.startswith("PASTE_"):
        sys.exit(f"Token file has placeholder text. Paste your real token into:\n  {TOKEN_PATH}")
    return tok


def http_post_json(url: str, body: dict, timeout: int = 180) -> dict:
    """POST JSON to an Apify endpoint. Returns parsed JSON response."""
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body_text = e.read().decode("utf-8", errors="ignore")[:300]
        sys.exit(f"Apify HTTP {e.code}: {body_text}")
    except urllib.error.URLError as e:
        sys.exit(f"Apify network error: {e.reason}")


def http_get_json(url: str) -> dict | list:
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body_text = e.read().decode("utf-8", errors="ignore")[:300]
        sys.exit(f"Apify HTTP {e.code}: {body_text}")


def check_token(token: str) -> None:
    """Verify the token works + print free tier remaining."""
    me = http_get_json(f"{API}/users/me?token={token}")
    user = me.get("data", {})
    name = user.get("username") or user.get("email") or "?"
    print(f"✓ Apify token valid. Authenticated as: {name}")
    # Usage info (best-effort — not all plans expose it the same way)
    try:
        limits = http_get_json(f"{API}/users/me/limits?token={token}")
        used_usd = limits.get("data", {}).get("current", {}).get("monthlyUsageUsd")
        limit_usd = limits.get("data", {}).get("limits", {}).get("monthlyUsageUsd")
        if used_usd is not None and limit_usd is not None:
            print(f"  Monthly usage so far: ${used_usd:.2f} / ${limit_usd:.2f}")
    except SystemExit:
        pass


def run_actor_sync(actor_id: str, token: str, input_body: dict, timeout: int = 180) -> list[dict]:
    """Run an Apify Actor synchronously and return its dataset items."""
    url = f"{API}/acts/{actor_id}/run-sync-get-dataset-items?token={token}"
    resp = http_post_json(url, input_body, timeout=timeout)
    # When using run-sync-get-dataset-items the response IS the dataset (list).
    if isinstance(resp, list):
        return resp
    # Some response shapes wrap it.
    if isinstance(resp, dict) and "items" in resp:
        return resp["items"]
    return []


def scrape_instagram(token: str, profile_url: str, results: int = 10) -> list[dict]:
    """Pull recent posts from an IG profile via the apify/instagram-scraper Actor."""
    body = {
        "directUrls": [profile_url],
        "resultsType": "posts",
        "resultsLimit": results,
        "searchType": "user",
        "addParentData": False,
    }
    raw = run_actor_sync(ACTOR_INSTAGRAM, token, body)
    out = []
    for item in raw:
        out.append({
            "platform": "instagram",
            "id": item.get("id") or item.get("shortCode"),
            "url": item.get("url"),
            "title": (item.get("caption") or "")[:120],
            "uploader": item.get("ownerUsername"),
            "upload_date": item.get("timestamp"),
            "duration_s": item.get("videoDuration"),
            "view_count": item.get("videoPlayCount") or item.get("videoViewCount"),
            "like_count": item.get("likesCount"),
            "comment_count": item.get("commentsCount"),
            "media_type": item.get("type") or item.get("productType"),
            "description_excerpt": (item.get("caption") or "")[:200],
        })
    return out


def scrape_facebook(token: str, page_url: str, results: int = 10) -> list[dict]:
    """Pull recent posts from a FB page via apify/facebook-posts-scraper."""
    body = {
        "startUrls": [{"url": page_url}],
        "resultsLimit": results,
    }
    raw = run_actor_sync(ACTOR_FACEBOOK, token, body)
    out = []
    for item in raw:
        views = item.get("viewsCount") or item.get("videoPostViewCount") or item.get("videoViewCount")
        likes = item.get("likes") or item.get("reactionLikeCount") or item.get("likesCount")
        if likes is None:
            reactions = item.get("reactions")
            if isinstance(reactions, dict):
                likes = reactions.get("likes")
        out.append({
            "platform": "facebook",
            "id": item.get("postId") or item.get("topLevelUrl"),
            "url": item.get("topLevelUrl") or item.get("url"),
            "title": (item.get("text") or "")[:120],
            "uploader": item.get("pageName") or item.get("user", {}).get("name"),
            "upload_date": item.get("time"),
            "view_count": views,
            "like_count": likes,
            "comment_count": item.get("commentsCount"),
            "repost_count": item.get("shares") or item.get("sharesCount"),
            "is_video": item.get("isVideo", False),
            "description_excerpt": (item.get("text") or "")[:200],
        })
    return out


def scrape_youtube(token: str, channel_url: str, results: int = 30,
                   since_date: str | None = None) -> list[dict]:
    """Pull recent uploads from a YT channel via streamers/youtube-scraper.

    Skips OAuth entirely — same token model as IG/FB. Returns records shaped
    like the rest of the pipeline (platform, url, title, view_count, etc.).

    `since_date` (ISO YYYY-MM-DD) filters OUT uploads before that date, e.g.
    "2026-06-08" for Talking Downey — the day the engagement began.
    Older videos exist on the channel but aren't ours to report on.
    """
    body = {
        "startUrls": [{"url": channel_url}],
        "maxResults": results,
        "maxResultsShorts": 0,
        "maxResultStreams": 0,
    }
    raw = run_actor_sync(ACTOR_YOUTUBE, token, body, timeout=300)
    out = []
    for item in raw:
        upload_date = item.get("date") or item.get("uploadDate") or item.get("publishedAt")
        if since_date and upload_date:
            if str(upload_date)[:10] < since_date:
                continue
        views = item.get("viewCount") or item.get("videoViewCount") or item.get("views")
        likes = item.get("likes") or item.get("likeCount")
        out.append({
            "platform": "youtube",
            "id": item.get("id") or item.get("videoId"),
            "url": item.get("url") or item.get("videoUrl"),
            "title": (item.get("title") or "")[:120],
            "uploader": item.get("channelName") or item.get("author"),
            "upload_date": upload_date,
            "duration_s": item.get("duration") or item.get("lengthSeconds"),
            "view_count": views,
            "like_count": likes,
            "comment_count": item.get("commentsCount") or item.get("commentCount"),
            "description_excerpt": (item.get("text") or item.get("description") or "")[:200],
        })
    return out


def scrape_tiktok(token: str, profile_url_or_handle: str, results: int = 30) -> list[dict]:
    """Pull recent posts from a TT profile via clockworks/tiktok-scraper.

    Accepts either a full profile URL or a bare handle ("@talkingdowney"
    or "talkingdowney"). Same token model as IG/FB/YT — no login flow.
    """
    handle = profile_url_or_handle
    if handle.startswith("http"):
        # extract "@user" or "user" from URL tail
        handle = handle.rstrip("/").split("/")[-1]
    handle = handle.lstrip("@")

    body = {
        "profiles": [handle],
        "resultsPerPage": results,
        "shouldDownloadCovers": False,
        "shouldDownloadVideos": False,
        "shouldDownloadSubtitles": False,
    }
    raw = run_actor_sync(ACTOR_TIKTOK, token, body, timeout=300)
    out = []
    for item in raw:
        stats = item.get("stats") or {}
        video_meta = item.get("videoMeta") or {}
        out.append({
            "platform": "tiktok",
            "id": item.get("id"),
            "url": item.get("webVideoUrl") or item.get("shareUrl") or item.get("url"),
            "title": (item.get("text") or "")[:120],
            "uploader": (item.get("authorMeta") or {}).get("name") or item.get("author") or handle,
            "upload_date": item.get("createTimeISO") or item.get("createTime"),
            "duration_s": video_meta.get("duration"),
            "view_count": stats.get("playCount") or item.get("playCount"),
            "like_count": stats.get("diggCount") or item.get("diggCount"),
            "comment_count": stats.get("commentCount") or item.get("commentCount"),
            "share_count": stats.get("shareCount") or item.get("shareCount"),
            "description_excerpt": (item.get("text") or "")[:200],
        })
    return out


def merge_into_snapshot(snap_path: Path, new_records: list[dict], episode: str, posted_at: str, tag: str) -> None:
    """Either create a fresh snapshot file or append to an existing one (deduped by url)."""
    now = datetime.now(timezone.utc).isoformat()
    if snap_path.exists():
        snap = json.loads(snap_path.read_text())
        existing_urls = {r.get("url") for r in snap.get("records", []) if r.get("url")}
        for rec in new_records:
            if rec.get("url") not in existing_urls:
                snap["records"].append(rec)
    else:
        snap = {
            "episode": episode,
            "posted_at": posted_at,
            "snapshot_tag": tag,
            "pulled_at": now,
            "records": new_records,
        }
    snap["pulled_at"] = now
    snap_path.write_text(json.dumps(snap, indent=2))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="Just verify the token works")
    ap.add_argument("--manifest", type=Path)
    ap.add_argument("--tag", help="Snapshot tag (t0, 24h, 72h, etc.)")
    ap.add_argument("--ig-results", type=int, default=30,
                    help="Max IG posts per profile (default 30 — covers ~3 weeks at 2-3 posts/day)")
    ap.add_argument("--fb-results", type=int, default=30, help="Max FB posts per page")
    ap.add_argument("--yt-results", type=int, default=30,
                    help="Max YT videos per channel (default 30 — covers recent uploads)")
    ap.add_argument("--tt-results", type=int, default=30,
                    help="Max TT videos per profile (default 30 — covers ~3 weeks)")
    args = ap.parse_args()

    token = load_token()

    if args.check:
        check_token(token)
        return

    if not args.manifest or not args.tag:
        sys.exit("Required: --manifest PATH --tag TAG  (or use --check)")
    if not args.manifest.is_file():
        sys.exit(f"Manifest not found: {args.manifest}")

    spec = yaml.safe_load(args.manifest.read_text())
    targets = spec.get("apify_targets") or {}
    if not targets:
        sys.exit(
            "No `apify_targets:` block in manifest. Add this section:\n"
            "  apify_targets:\n"
            "    instagram_profile: \"https://www.instagram.com/talkingdowney/\"\n"
            "    facebook_page:     \"https://www.facebook.com/TalkingDowney\"\n"
        )

    out_dir = args.manifest.parent / "snapshots"
    out_dir.mkdir(parents=True, exist_ok=True)
    snap_path = out_dir / f"{args.tag}.json"

    print(f"Episode: {spec.get('episode')}")
    print(f"Snapshot tag: {args.tag}\n")

    collected: list[dict] = []

    ig_url = targets.get("instagram_profile")
    if ig_url:
        print(f"  → Instagram via Apify: {ig_url}")
        ig_items = scrape_instagram(token, ig_url, results=args.ig_results)
        print(f"    ✓ {len(ig_items)} posts")
        for i, rec in enumerate(ig_items, 1):
            rec["label"] = f"instagram_post_{i}"
        collected.extend(ig_items)

    fb_url = targets.get("facebook_page")
    if fb_url:
        print(f"  → Facebook via Apify: {fb_url}")
        fb_items = scrape_facebook(token, fb_url, results=args.fb_results)
        print(f"    ✓ {len(fb_items)} posts")
        for i, rec in enumerate(fb_items, 1):
            rec["label"] = f"facebook_post_{i}"
        collected.extend(fb_items)

    yt_url = targets.get("youtube_channel")
    if yt_url:
        yt_since = targets.get("youtube_since_date")  # e.g. "2026-06-08"
        print(f"  → YouTube via Apify: {yt_url}"
              + (f" (since {yt_since})" if yt_since else ""))
        yt_items = scrape_youtube(token, yt_url, results=args.yt_results,
                                  since_date=yt_since)
        print(f"    ✓ {len(yt_items)} videos")
        for i, rec in enumerate(yt_items, 1):
            rec["label"] = f"youtube_video_{i}"
        collected.extend(yt_items)

    tt_url = targets.get("tiktok_profile")
    if tt_url:
        print(f"  → TikTok via Apify: {tt_url}")
        tt_items = scrape_tiktok(token, tt_url, results=args.tt_results)
        print(f"    ✓ {len(tt_items)} videos")
        for i, rec in enumerate(tt_items, 1):
            rec["label"] = f"tiktok_video_{i}"
        collected.extend(tt_items)

    merge_into_snapshot(
        snap_path,
        collected,
        episode=spec.get("episode", ""),
        posted_at=spec.get("posted_at", ""),
        tag=args.tag,
    )
    print(f"\n✓ Merged into: {snap_path}")
    print(f"  Added {len(collected)} records.")


if __name__ == "__main__":
    main()
