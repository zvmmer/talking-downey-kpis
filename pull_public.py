"""Talking Downey — pull PUBLIC KPIs for an episode (no auth needed).

Uses yt-dlp's metadata extraction (--skip-download) to read whatever each
platform exposes publicly. Same data anyone with the URL can see. No
developer apps, no client permission required.

Reliability per platform:
    YouTube       ✅ views, likes, comments, duration, publish date
    TikTok        ✅ views, likes, comments, shares, uploader
    Instagram     ⚠️  caption + sometimes engagement; IG blocks scraping unpredictably
    Facebook      ⚠️  very limited; mostly just title/description
    Spotify       ❌ not supported
    Apple         ❌ not supported

The point of this script is to BUILD A CASE for proper API access. The
numbers it captures are real but partial — show the team what you can
already do, propose what you could automate with proper credentials.

Usage
-----
    # First snapshot for an episode (creates the folder):
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_public.py \\
        --manifest Projects/talking_downey/kpis/episodes_public/EP47/manifest.yaml \\
        --tag t0

    # 24h later:
    MCP/.venv/bin/python Projects/talking_downey/kpis/pull_public.py \\
        --manifest .../EP47/manifest.yaml --tag 24h

    # 48h, 72h, etc. — any tag you want:
    --tag 48h
    --tag 72h
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

try:
    import yaml
except ImportError:
    sys.exit("Need PyYAML. Install: MCP/.venv/bin/python -m pip install PyYAML")

YT_DLP = "/opt/homebrew/bin/yt-dlp"


def platform_of(url: str) -> str:
    u = url.lower()
    if "youtube.com" in u or "youtu.be" in u:
        return "youtube_short" if "/shorts/" in u else "youtube"
    if "instagram.com" in u:
        if "/reel/" in u:
            return "instagram_reel"
        return "instagram"
    if "tiktok.com" in u:
        return "tiktok"
    if "facebook.com" in u or "fb.watch" in u:
        return "facebook"
    return "unknown"


def yt_dlp_dump(url: str) -> dict:
    """Return parsed metadata or {'error': '...'} on failure."""
    cmd = [
        YT_DLP, "--skip-download", "--no-warnings",
        "--dump-single-json", "--no-playlist", url,
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except subprocess.TimeoutExpired:
        return {"error": "timeout"}
    if result.returncode != 0:
        # Common case for IG: "Restricted Video" or "Login required"
        err = (result.stderr or "")[:200].strip()
        return {"error": err or f"yt-dlp exit {result.returncode}"}
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as e:
        return {"error": f"parse failed: {e}"}


def normalize(plat: str, meta: dict, url: str) -> dict:
    """Reduce the giant yt-dlp blob to the fields we care about."""
    if "error" in meta:
        return {"platform": plat, "url": url, "error": meta["error"]}
    return {
        "platform": plat,
        "url": url,
        "id": meta.get("id"),
        "title": (meta.get("title") or meta.get("description") or "")[:120],
        "uploader": meta.get("uploader") or meta.get("channel"),
        "upload_date": meta.get("upload_date"),  # YYYYMMDD
        "timestamp": meta.get("timestamp"),
        "duration_s": meta.get("duration"),
        "view_count": meta.get("view_count"),
        "like_count": meta.get("like_count"),
        "comment_count": meta.get("comment_count"),
        "repost_count": meta.get("repost_count"),  # TikTok shares
        "description_excerpt": (meta.get("description") or "")[:200],
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--manifest", required=True, type=Path,
                    help="YAML file listing episode URLs (see episodes_public/_manifest_template.yaml)")
    ap.add_argument("--tag", required=True,
                    help="Snapshot tag: t0, 24h, 48h, 72h, 1wk, etc.")
    ap.add_argument("--out-dir", type=Path,
                    help="Override output folder (defaults to manifest's parent dir / snapshots/)")
    args = ap.parse_args()

    if not args.manifest.is_file():
        sys.exit(f"Manifest not found: {args.manifest}")

    spec = yaml.safe_load(args.manifest.read_text())
    urls = spec.get("urls", {})
    if not urls:
        sys.exit("Manifest has no 'urls:' section")

    out_dir = args.out_dir or (args.manifest.parent / "snapshots")
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{args.tag}.json"

    now = datetime.now(timezone.utc).isoformat()
    print(f"Episode: {spec.get('episode', '?')}")
    print(f"Snapshot tag: {args.tag}")
    print(f"Pulled at: {now}\n")

    records = []
    for label, url in urls.items():
        if not url:
            print(f"  [skip] {label}: no URL set")
            continue
        plat = platform_of(url)
        print(f"  → {label} ({plat})…", end=" ", flush=True)
        meta = yt_dlp_dump(url)
        rec = normalize(plat, meta, url)
        rec["label"] = label
        records.append(rec)
        if "error" in rec:
            print(f"⚠️  {rec['error'][:60]}")
        else:
            v = rec.get("view_count")
            l = rec.get("like_count")
            c = rec.get("comment_count")
            print(f"✓  views={v}  likes={l}  comments={c}")

    snapshot = {
        "episode": spec.get("episode"),
        "posted_at": spec.get("posted_at"),
        "snapshot_tag": args.tag,
        "pulled_at": now,
        "records": records,
    }
    out_path.write_text(json.dumps(snapshot, indent=2))
    print(f"\n✓ Saved: {out_path}")


if __name__ == "__main__":
    main()
