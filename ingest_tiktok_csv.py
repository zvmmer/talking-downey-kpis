"""Talking Downey — TikTok CSV ingester.

Reads a CSV exported from TikTok Studio (Analytics → Download data), normalizes
to the common KPI record shape, and writes per-video JSON to the episodes
folder so reports can combine TikTok with YouTube + Meta data.

The TikTok Studio CSV format has shifted over time. This ingester is tolerant:
it auto-detects which columns map to views/likes/comments/shares/etc. by
matching common header names. If your export has different headers, pass
`--inspect` to see what was found and which mapping was used.

Usage
-----
    # Ingest the most recent CSV in the exports folder:
    MCP/.venv/bin/python Projects/talking_downey/kpis/ingest_tiktok_csv.py

    # Specify the CSV path explicitly:
    MCP/.venv/bin/python Projects/talking_downey/kpis/ingest_tiktok_csv.py \\
        --csv "/Volumes/MAIN DRIVE/Talking Downey/POSTING DATA/tiktok_exports/2026-06-15.csv"

    # Just inspect the CSV without ingesting:
    MCP/.venv/bin/python Projects/talking_downey/kpis/ingest_tiktok_csv.py --inspect
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_EXPORTS_DIR = Path(
    "/Volumes/MAIN DRIVE/Talking Downey/POSTING DATA/tiktok_exports"
)
HERE = Path(__file__).resolve().parent
EPISODES_DIR = HERE / "episodes"

# Column header → normalized field. Lowercased, stripped. Any header that
# *contains* the key (substring match) gets mapped.
HEADER_MAP = {
    "video title": "title",
    "video link": "url",
    "publish time": "posted_at",
    "post time": "posted_at",
    "date posted": "posted_at",
    "video views": "views",
    "total views": "views",
    "views": "views",
    "likes": "likes",
    "comments": "comments",
    "shares": "shares",
    "saves": "saves",
    "average watch time": "avg_watch_time_s",
    "watch time": "watch_time_s",
    "total play time": "watch_time_s",
    "completion rate": "completion_rate",
    "reach": "reach",
    "video id": "video_id",
    "id": "video_id",
}


def latest_csv(folder: Path) -> Path | None:
    if not folder.is_dir():
        return None
    csvs = sorted(
        (p for p in folder.iterdir()
         if p.suffix.lower() == ".csv" and not p.name.startswith(".")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return csvs[0] if csvs else None


def detect_mapping(headers: list[str]) -> dict[str, str]:
    """Return {original_header: normalized_field} for headers we recognize."""
    found: dict[str, str] = {}
    for h in headers:
        lo = h.strip().lower()
        for k, v in HEADER_MAP.items():
            if k in lo:
                # First match wins (don't overwrite — avoids 'video views' being
                # overwritten by 'views' later if both appear).
                if h not in found:
                    found[h] = v
                break
    return found


def parse_int(s: str) -> int:
    if not s:
        return 0
    s = s.strip().replace(",", "").replace("%", "")
    try:
        return int(float(s))
    except ValueError:
        return 0


def parse_pct(s: str) -> float:
    if not s:
        return 0.0
    s = s.strip().replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def normalize_row(row: dict, mapping: dict[str, str]) -> dict:
    out: dict = {}
    for src_col, norm in mapping.items():
        raw = row.get(src_col, "")
        if norm in {"views", "likes", "comments", "shares", "saves", "reach", "watch_time_s"}:
            out[norm] = parse_int(raw)
        elif norm == "completion_rate":
            out[norm] = parse_pct(raw)
        elif norm == "avg_watch_time_s":
            # Could be "0:12" or "12s" or "12.4" — try to coerce to seconds.
            s = raw.strip().lower().rstrip("s")
            if ":" in s:
                parts = s.split(":")
                try:
                    out[norm] = int(parts[0]) * 60 + int(float(parts[1]))
                except ValueError:
                    out[norm] = 0
            else:
                try:
                    out[norm] = float(s)
                except ValueError:
                    out[norm] = 0
        else:
            out[norm] = raw.strip()
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", type=Path, help="Path to a specific CSV (default: latest in exports folder)")
    ap.add_argument("--exports-dir", type=Path, default=DEFAULT_EXPORTS_DIR,
                    help="Folder to scan for latest CSV (default: MAIN DRIVE/Talking Downey/POSTING DATA/tiktok_exports/)")
    ap.add_argument("--inspect", action="store_true",
                    help="Print detected headers + mapping without writing JSON")
    ap.add_argument("--out-dir", type=Path, default=EPISODES_DIR / "_tiktok_inbox",
                    help="Where to write normalized JSON (default: kpis/episodes/_tiktok_inbox/)")
    args = ap.parse_args()

    csv_path: Path | None = args.csv or latest_csv(args.exports_dir)
    if not csv_path or not csv_path.exists():
        sys.exit(
            f"No CSV found. Either pass --csv PATH, or drop the latest export into:\n"
            f"  {args.exports_dir}"
        )

    print(f"Reading: {csv_path}")
    with csv_path.open(newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        headers = reader.fieldnames or []
        mapping = detect_mapping(headers)
        rows = list(reader)

    print(f"Detected {len(headers)} columns; mapped {len(mapping)} of them.")
    print(f"\nMapping:")
    for src, norm in mapping.items():
        print(f"  {src:<32}  →  {norm}")
    unmapped = [h for h in headers if h not in mapping]
    if unmapped:
        print(f"\nUnmapped (ignored): {unmapped}")

    if args.inspect:
        print(f"\n{len(rows)} rows. First row preview:")
        if rows:
            print(json.dumps(normalize_row(rows[0], mapping), indent=2))
        return

    args.out_dir.mkdir(parents=True, exist_ok=True)

    pulled_at = datetime.now(timezone.utc).isoformat()
    written = 0
    for row in rows:
        rec = normalize_row(row, mapping)
        if not rec.get("video_id") and not rec.get("url"):
            continue
        # Build a stable key for the filename
        key = rec.get("video_id") or rec.get("url", "").rsplit("/", 1)[-1] or "unknown"
        rec["_source"] = "tiktok_csv"
        rec["_csv_file"] = csv_path.name
        rec["_pulled_at"] = pulled_at
        out_file = args.out_dir / f"tiktok_{key}.json"
        out_file.write_text(json.dumps(rec, indent=2))
        written += 1

    print(f"\n✓ Wrote {written} normalized records to {args.out_dir}")


if __name__ == "__main__":
    main()
