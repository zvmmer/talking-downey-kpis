"""Talking Downey — generate a client-ready markdown report from the public
snapshots for a single episode.

Reads every snapshot JSON in `<episode>/snapshots/*.json`, sorts them by
the tag's implied time order, and produces a markdown table showing the
trajectory of each platform's KPIs across all snapshots.

Usage
-----
    MCP/.venv/bin/python Projects/talking_downey/kpis/report_public.py \\
        --episode Projects/talking_downey/kpis/episodes_public/EP47
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

# Order tags chronologically. Anything unknown sorts after these.
TAG_ORDER = ["t0", "1h", "6h", "12h", "24h", "48h", "72h", "1wk", "2wk", "1mo"]


def sort_key(tag: str) -> tuple:
    if tag in TAG_ORDER:
        return (0, TAG_ORDER.index(tag))
    return (1, tag)


def fmt_int(n) -> str:
    if n is None:
        return "—"
    try:
        return f"{int(n):,}"
    except (ValueError, TypeError):
        return "—"


def delta(curr, prev) -> str:
    """Return a +N delta string vs the previous snapshot, or '' if N/A."""
    if curr is None or prev is None:
        return ""
    try:
        d = int(curr) - int(prev)
    except (ValueError, TypeError):
        return ""
    if d == 0:
        return ""
    sign = "+" if d > 0 else ""
    return f" ({sign}{d:,})"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--episode", required=True, type=Path,
                    help="Episode folder (must contain snapshots/*.json)")
    ap.add_argument("--out", type=Path,
                    help="Output markdown path (default: <episode>/report.md)")
    args = ap.parse_args()

    snap_dir = args.episode / "snapshots"
    if not snap_dir.is_dir():
        sys.exit(f"No snapshots dir at {snap_dir}")
    snaps = []
    for f in snap_dir.glob("*.json"):
        if f.name.startswith("."):
            continue
        try:
            snaps.append(json.loads(f.read_text()))
        except json.JSONDecodeError:
            print(f"  warn: skipping malformed {f}", file=sys.stderr)
    if not snaps:
        sys.exit(f"No snapshot JSONs found in {snap_dir}")

    snaps.sort(key=lambda s: sort_key(s.get("snapshot_tag", "")))
    episode = snaps[0].get("episode", args.episode.name)
    posted_at = snaps[0].get("posted_at")
    tags = [s["snapshot_tag"] for s in snaps]

    # Build per-platform trajectory: {label: [snapshot1_rec, snapshot2_rec, ...]}
    per_label: dict[str, list] = {}
    for snap in snaps:
        for rec in snap.get("records", []):
            label = rec.get("label", "unknown")
            per_label.setdefault(label, []).append(rec)
    # Pad missing entries with None so columns line up
    for label, recs in per_label.items():
        if len(recs) < len(snaps):
            # Walk snapshots in order, fill missing with placeholders
            ordered = []
            for snap in snaps:
                match = next((r for r in snap.get("records", []) if r.get("label") == label), None)
                ordered.append(match or {"label": label, "error": "(missing)"})
            per_label[label] = ordered

    lines = []
    lines.append(f"# Talking Downey — {episode}")
    if posted_at:
        lines.append(f"*Posted: {posted_at}*")
    lines.append("")
    lines.append(f"Report generated: {datetime.now(timezone.utc).isoformat()}")
    lines.append("")
    lines.append("**Data source: public engagement counts only (no API auth yet).** "
                 "Full retention, traffic-source, and demographic data unlocks once "
                 "we have YouTube Studio + Meta Graph + TikTok Studio API access.")
    lines.append("")
    lines.append(f"## Trajectory across {len(snaps)} snapshots: {', '.join(tags)}")
    lines.append("")

    for label, recs in per_label.items():
        plat = next((r.get("platform") for r in recs if "platform" in r), "?")
        url = next((r.get("url") for r in recs if r.get("url")), "")
        title = next((r.get("title") for r in recs if r.get("title")), "")
        lines.append(f"### {label} · *{plat}*")
        if title:
            lines.append(f"**{title}**")
        if url:
            lines.append(f"[Open ↗]({url})")
        lines.append("")

        if any("error" in r and "(missing)" not in r.get("error", "") for r in recs):
            # Show errors but still try metric table
            errs = [(r.get("error", "") if "error" in r else "—") for r in recs]
            lines.append(f"_Errors: {' / '.join(errs)}_")
            lines.append("")

        # Metrics table
        header = "| Metric | " + " | ".join(tags) + " |"
        sep = "|---|" + "|".join(["---"] * len(tags)) + "|"
        lines.append(header)
        lines.append(sep)

        for metric, key in [("Views", "view_count"),
                            ("Likes", "like_count"),
                            ("Comments", "comment_count"),
                            ("Shares/Reposts", "repost_count")]:
            row = [metric]
            prev_val = None
            for r in recs:
                v = r.get(key)
                cell = fmt_int(v) + delta(v, prev_val)
                row.append(cell)
                prev_val = v
            # Skip rows where all values are dashes (no data anywhere)
            if all(c.strip() == "—" for c in row[1:]):
                continue
            lines.append("| " + " | ".join(row) + " |")
        lines.append("")

    # What's NOT in this report (build the case for API access)
    lines.append("---")
    lines.append("")
    lines.append("## What's NOT in this report (and what API access would unlock)")
    lines.append("")
    lines.append("This report uses only **publicly visible** engagement counts. With proper API access we'd add:")
    lines.append("")
    lines.append("- **YouTube**: watch time, average view duration, retention curve, traffic sources, "
                 "subscribers gained, click-through rate, audience demographics")
    lines.append("- **Instagram**: reach, impressions, profile actions, saves, plays "
                 "(IG's `like_count` is not always publicly exposed)")
    lines.append("- **Facebook**: post reach, organic vs paid, page follower growth attribution")
    lines.append("- **TikTok**: full TikTok Studio analytics (watch time, completion rate, "
                 "for-you-page reach, audience geography)")
    lines.append("")
    lines.append("Setup is one-time per platform, $0 cost, all read-only. Setup docs already drafted: "
                 "see `kpis/SETUP_GOOGLE_CLOUD.md` and `kpis/SETUP_META.md`.")
    lines.append("")

    out = args.out or (args.episode / "report.md")
    out.write_text("\n".join(lines))
    print(f"✓ Wrote: {out}")
    print(f"  {len(per_label)} platforms × {len(snaps)} snapshots")


if __name__ == "__main__":
    main()
