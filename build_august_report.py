"""Talking Downey — August 2026 Performance Report deck.

Third in the series (after build_june_report.py + build_july_report.py). July
was the "brand momentum" story; August is the counter-move — volume held but
reach softened, no scandal swings, 57 percent of posts landed in "other" as
the show explored formats without a clear pillar.

Slides:
    1. Title — "August 2026 — The Exploration Month"
    2. Since June 8 — cumulative rollup across all 3 months
    3. July -> August — the delta table with the "exploration, not spicy" read
    4. Platform-by-platform — Aug vs July per-platform reach
    5. Content taxonomy — August-only category breakdown (with the "other" honesty)
    6. Breakouts — top August posts across platforms
    7. The Read — takeaway + September ask

Uses report_helpers.py (shared with overall + sponsor decks).

Usage
-----
    ~/talking-downey-kpis/venv/bin/python build_august_report.py

Outputs `MONTHLY_REPORT_2026-08.pptx` at the kpis/ root by default.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from report_helpers import (
    PALETTE, LAYOUT_W, LAYOUT_H,
    add_rect, add_text, fmt_int,
    load_all_snapshots, posts_in_window, group_by_category,
    slide_new, slide_header, slide_footer_takeaway,
)

JUNE_START   = "2026-06-08"
JUNE_END     = "2026-06-30"
JULY_START   = "2026-07-01"
JULY_END     = "2026-07-31"
AUGUST_START = "2026-08-01"
AUGUST_END   = "2026-08-31"


# ── slides ──────────────────────────────────────────────────────────────

def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, LAYOUT_W, LAYOUT_H, PALETTE["ink"])
    add_rect(slide, 0, 0, 0.45, LAYOUT_H, PALETTE["gold"])   # gold = exploration/reflection

    add_text(slide, 1.0, 1.2, 11, 0.5,
             "AUGUST 2026 PERFORMANCE REPORT",
             size=14, color=PALETTE["gold"], bold=True)
    add_text(slide, 1.0, 1.9, 11, 1.3,
             "Talking Downey", size=64, color=PALETTE["white"],
             bold=True, font="Georgia")
    add_text(slide, 1.0, 3.4, 11, 0.6,
             "The exploration month — powder dry, muscle memory intact",
             size=24, color=PALETTE["cream"], font="Georgia")

    add_text(slide, 1.0, 5.0, 8, 0.4,
             "Tracking period: Aug 1 – Aug 31, 2026",
             size=13, color=PALETTE["muted"])
    add_text(slide, 1.0, 5.4, 8, 0.4,
             "Comparison baseline: July 2026",
             size=13, color=PALETTE["muted"])
    add_text(slide, 1.0, 5.8, 8, 0.4,
             "4 platforms · IG · TikTok · Facebook · YouTube",
             size=13, color=PALETTE["muted"])

    add_text(slide, 1.0, 6.6, 11, 0.4,
             "Prepared by Zhamir Pascual — Z",
             size=12, color=PALETTE["muted"])


def slide_since_june(prs, all_posts_since_june):
    """Cumulative rollup across the whole show, since launch. This is the 'we're not just a monthly report' slide."""
    slide = slide_new(prs)
    slide_header(slide, "Since June 8",
                 "Everything the show has shipped",
                 "Cumulative totals from launch through today. Whatever the month looked like, this is the body of work.")

    posts = all_posts_since_june
    total_views    = sum(int(p.get("view_count") or 0)    for p in posts)
    total_likes    = sum(int(p.get("like_count") or 0)    for p in posts)
    total_comments = sum(int(p.get("comment_count") or 0) for p in posts)

    metrics = [
        (fmt_int(len(posts)),        "posts published",     PALETTE["navy"]),
        (fmt_int(total_views),       "total views",         PALETTE["sky"]),
        (fmt_int(total_likes),       "total likes",         PALETTE["cherry"]),
        (fmt_int(total_comments),    "total comments",      PALETTE["green"]),
    ]

    card_w = 2.85
    gap = 0.2
    total_w = card_w * 4 + gap * 3
    x_start = (LAYOUT_W - total_w) / 2
    y = 2.6

    for i, (val, label, color) in enumerate(metrics):
        x = x_start + i * (card_w + gap)
        add_rect(slide, x, y, card_w, 2.4, PALETTE["white"])
        add_rect(slide, x, y, 0.06, 2.4, color)
        add_text(slide, x + 0.2, y + 0.4, card_w - 0.4, 1.1,
                 val, size=44, color=color, bold=True, font="Georgia")
        add_text(slide, x + 0.2, y + 1.6, card_w - 0.4, 0.4,
                 label.upper(), size=10, color=PALETTE["muted"], bold=True)

    # Per-platform mini-strip below
    plat_views = {}
    plat_counts = {}
    for p in posts:
        pl = p.get("platform", "?")
        plat_views[pl] = plat_views.get(pl, 0) + int(p.get("view_count") or 0)
        plat_counts[pl] = plat_counts.get(pl, 0) + 1

    plat_labels = [("tiktok", "TikTok"), ("instagram", "Instagram"),
                   ("facebook", "Facebook"), ("youtube", "YouTube")]
    strip_y = 5.4
    strip_x = x_start
    for i, (key, label) in enumerate(plat_labels):
        cx = strip_x + i * (card_w + gap)
        v = plat_views.get(key, 0)
        n = plat_counts.get(key, 0)
        add_text(slide, cx + 0.2, strip_y, card_w - 0.4, 0.3,
                 label.upper(), size=10, color=PALETTE["muted"], bold=True)
        add_text(slide, cx + 0.2, strip_y + 0.3, card_w - 0.4, 0.5,
                 f"{fmt_int(v)} views",
                 size=16, color=PALETTE["ink"], bold=True, font="Georgia")
        add_text(slide, cx + 0.2, strip_y + 0.75, card_w - 0.4, 0.3,
                 f"{n} posts", size=10, color=PALETTE["muted"])

    slide_footer_takeaway(slide,
        "One episode. Three months. Four platforms. The catalog is the moat — every reel keeps working after the algorithm moves on.")


def slide_month_shift(prs, aug_posts, jul_posts):
    """The spine slide — 'here's what changed from July to August', with the honest read."""
    slide = slide_new(prs)
    slide_header(slide, "July → August",
                 "The exploration month",
                 "Volume mostly held. Reach softened. No scandal swings taken — the viral engine sat on the bench while the show tried new formats.")

    # Big numbers: post count and views delta
    v_jul = len(jul_posts)
    v_aug = len(aug_posts)
    views_jul = sum(int(p.get("view_count") or 0) for p in jul_posts)
    views_aug = sum(int(p.get("view_count") or 0) for p in aug_posts)

    # Left: post count
    add_text(slide, 0.8, 2.4, 5.8, 1.5,
             fmt_int(v_aug), size=110, color=PALETTE["gold"], bold=True,
             font="Georgia", align="center")
    add_text(slide, 0.8, 4.0, 5.8, 0.5,
             "posts in August", size=16, color=PALETTE["navy"], align="center")
    pct_posts = f"{((v_aug - v_jul) / v_jul * 100):+.0f}%" if v_jul else "—"
    add_text(slide, 0.8, 4.5, 5.8, 0.4,
             f"({v_jul} in July · {pct_posts})",
             size=11, color=PALETTE["muted"], align="center")

    # Right cards: the honest deltas
    def _view_avg(posts):
        rel = [int(p.get("view_count") or 0) for p in posts]
        return sum(rel) / len(rel) if rel else 0

    scandal_grp = {"trujillo_scandal", "lisette_scandal", "council_chaos",
                   "ai_fake_news", "ice_immigration"}
    n_scandal_jul = sum(1 for p in jul_posts if p.get("_category") in scandal_grp)
    n_scandal_aug = sum(1 for p in aug_posts if p.get("_category") in scandal_grp)

    n_other_jul = sum(1 for p in jul_posts if p.get("_category") == "other")
    n_other_aug = sum(1 for p in aug_posts if p.get("_category") == "other")
    pct_other_aug = (n_other_aug / max(1, v_aug)) * 100

    shifts = [
        ("Total reach",       views_jul, views_aug, PALETTE["cherry"],
         "views across all platforms"),
        ("Avg views / post",  _view_avg(jul_posts), _view_avg(aug_posts), PALETTE["cherry"],
         "reach per piece of content"),
        ("Scandal content",   n_scandal_jul, n_scandal_aug, PALETTE["gold"],
         "no swings taken this month"),
    ]

    y = 2.4
    x = 7.0
    for label, jul, aug, color, sub in shifts:
        add_rect(slide, x, y, 5.6, 1.2, PALETTE["white"])
        add_rect(slide, x, y, 0.08, 1.2, color)
        add_text(slide, x + 0.25, y + 0.12, 4, 0.3,
                 label.upper(), size=10, color=color, bold=True)
        add_text(slide, x + 0.25, y + 0.4, 3.2, 0.5,
                 f"{jul:,.0f} → {aug:,.0f}" if isinstance(jul, float) else f"{fmt_int(jul)} → {fmt_int(aug)}",
                 size=18, color=PALETTE["ink"], bold=True, font="Georgia")
        arrow = "▲" if aug > jul else ("▼" if aug < jul else "▬")
        arrow_color = PALETTE["green"] if aug > jul else (PALETTE["cherry"] if aug < jul else PALETTE["muted"])
        pct = f"{((aug - jul) / jul * 100):+.0f}%" if jul else "—"
        add_text(slide, x + 4.0, y + 0.4, 1.5, 0.6,
                 f"{arrow} {pct}", size=18, color=arrow_color, bold=True)
        add_text(slide, x + 0.25, y + 0.9, 5, 0.25,
                 sub, size=9, color=PALETTE["muted"])
        y += 1.4

    # The "other = 57%" honesty card
    add_rect(slide, 7.0, 6.6, 5.6, 0.9, PALETTE["soft"])
    add_rect(slide, 7.0, 6.6, 0.08, 0.9, PALETTE["muted"])
    add_text(slide, 7.25, 6.7, 5.2, 0.3,
             "THE EXPLORATION SIGNAL", size=10, color=PALETTE["muted"], bold=True)
    add_text(slide, 7.25, 7.0, 5.2, 0.5,
             f"{n_other_aug} of {v_aug} posts ({pct_other_aug:.0f}%) landed in \"other\" — no defined pillar.",
             size=11, color=PALETTE["ink"], bold=True)

    slide_footer_takeaway(slide,
        "No swings, no home runs. But the catalog kept growing and the food + community lanes quietly held their averages.")


def slide_platform_growth(prs, snaps):
    """Total views per platform, July window vs August window."""
    slide = slide_new(prs)
    slide_header(slide, "Platform-by-platform",
                 "Where reach softened — and where it didn't",
                 "Total views per platform in August vs the July baseline. Facebook actually held up best by average reach per post — unusual for this show.")

    jul_posts = posts_in_window(snaps, JULY_START, JULY_END)
    aug_posts = posts_in_window(snaps, AUGUST_START, AUGUST_END)

    platforms = [("instagram", "Instagram"), ("tiktok", "TikTok"),
                 ("facebook", "Facebook"), ("youtube", "YouTube")]
    card_w = 2.85
    gap = 0.2
    total_w = card_w * 4 + gap * 3
    x_start = (LAYOUT_W - total_w) / 2
    y = 2.5

    for i, (key, label) in enumerate(platforms):
        x = x_start + i * (card_w + gap)
        jul_v = sum(int(p.get("view_count") or 0) for p in jul_posts if p.get("platform") == key)
        aug_v = sum(int(p.get("view_count") or 0) for p in aug_posts if p.get("platform") == key)
        jul_n = sum(1 for p in jul_posts if p.get("platform") == key)
        aug_n = sum(1 for p in aug_posts if p.get("platform") == key)

        add_rect(slide, x, y, card_w, 3.8, PALETTE["white"])
        add_rect(slide, x, y, 0.06, 3.8, PALETTE["gold"])
        add_text(slide, x + 0.2, y + 0.2, card_w - 0.4, 0.4,
                 label.upper(), size=11, color=PALETTE["gold"], bold=True)

        add_text(slide, x + 0.2, y + 0.65, card_w - 0.4, 0.9,
                 fmt_int(aug_v), size=36, color=PALETTE["navy"],
                 bold=True, font="Georgia")
        add_text(slide, x + 0.2, y + 1.55, card_w - 0.4, 0.3,
                 f"August views ({aug_n} posts)",
                 size=10, color=PALETTE["muted"])

        add_rect(slide, x + 0.2, y + 2.0, card_w - 0.4, 0.015, PALETTE["rule"])
        add_text(slide, x + 0.2, y + 2.15, card_w - 0.4, 0.3,
                 "JULY BASELINE", size=10, color=PALETTE["muted"], bold=True)
        add_text(slide, x + 0.2, y + 2.45, card_w - 0.4, 0.6,
                 fmt_int(jul_v), size=24, color=PALETTE["ink"], bold=True)
        add_text(slide, x + 0.2, y + 3.1, card_w - 0.4, 0.3,
                 f"({jul_n} posts)", size=10, color=PALETTE["muted"])

    slide_footer_takeaway(slide,
        "Facebook held its ground on average reach. TikTok still leads absolute views but softened. IG took the biggest hit.")


def slide_categories_august(prs, aug_posts):
    """August category breakdown — including the 'other = 57%' honesty."""
    slide = slide_new(prs)
    slide_header(slide, "Content taxonomy",
                 "What August looked like",
                 "Average views per post by category. The show explored — most posts didn't map to a defined pillar.")

    buckets = group_by_category(aug_posts)

    scored = []
    for cat, posts in buckets.items():
        if not posts:
            continue
        views = [int(p.get("view_count") or 0) for p in posts]
        scored.append({
            "cat": cat,
            "count": len(posts),
            "total": sum(views),
            "avg": sum(views) / len(views),
        })
    scored.sort(key=lambda s: -s["total"])
    top = scored[:8]

    card_w, card_h, gap = 2.9, 1.9, 0.15
    x0 = (LAYOUT_W - card_w * 4 - gap * 3) / 2
    y0 = 2.3

    label_map = {
        "trujillo_scandal":              ("Trujillo Scandal",       PALETTE["cherry"]),
        "lisette_scandal":               ("Lisette Scandal",        PALETTE["cherry"]),
        "ai_fake_news":                  ("AI Fake News",           PALETTE["cherry"]),
        "council_chaos":                 ("Council Chaos",          PALETTE["cherry"]),
        "ice_immigration":               ("ICE / Immigration",      PALETTE["cherry"]),
        "elections":                     ("Elections",              PALETTE["navy"]),
        "judicial":                      ("Judicial",               PALETTE["navy"]),
        "politics_other":                ("Politics (other)",       PALETTE["navy"]),
        "legislation_policy_watch":      ("Legislation Watch",      PALETTE["navy"]),
        "political_individual_highlight":("Political Highlight",    PALETTE["sky"]),
        "resident_highlight":            ("Resident Highlight",     PALETTE["green"]),
        "downtown_development":          ("Downtown Development",   PALETTE["gold"]),
        "american_pride":                ("American Pride",         PALETTE["cherry"]),
        "military_service":              ("Military Service",       PALETTE["cherry"]),
        "fifa_positive":                 ("FIFA — Positive",        PALETTE["green"]),
        "fifa_critique":                 ("FIFA — Critique",        PALETTE["muted"]),
        "food":                          ("Food",                   PALETTE["green"]),
        "brand_growth":                  ("Brand Growth",           PALETTE["blush"]),
        "community_events":              ("Community Events",       PALETTE["ink"]),
        "community_local":               ("Community Local",        PALETTE["ink"]),
        "other":                         ("Other (unmapped)",       PALETTE["muted"]),
    }

    for i, s in enumerate(top):
        col = i % 4
        row = i // 4
        x = x0 + col * (card_w + gap)
        y = y0 + row * (card_h + gap)
        label, color = label_map.get(s["cat"], (s["cat"].title(), PALETTE["muted"]))

        add_rect(slide, x, y, card_w, card_h, PALETTE["white"])
        add_rect(slide, x, y, 0.06, card_h, color)
        add_text(slide, x + 0.2, y + 0.15, card_w - 0.4, 0.3,
                 label.upper(), size=9, color=color, bold=True)
        add_text(slide, x + 0.2, y + 0.45, card_w - 0.4, 0.7,
                 f"{s['avg']:,.0f}", size=28, color=PALETTE["navy"],
                 bold=True, font="Georgia")
        add_text(slide, x + 0.2, y + 1.15, card_w - 0.4, 0.25,
                 "avg views/post", size=9, color=PALETTE["muted"])
        add_text(slide, x + 0.2, y + 1.42, card_w - 0.4, 0.5,
                 f"{s['count']} posts · {s['total']:,} total views",
                 size=9, color=PALETTE["ink"])

    slide_footer_takeaway(slide,
        "Food + Community Local kept working at July-comparable per-post reach. The taxonomy needs new tags to catch what \"other\" actually was.")


def slide_top_posts_august(prs, aug_posts):
    """Top August posts across all platforms."""
    slide = slide_new(prs)
    slide_header(slide, "Breakouts",
                 "Top August posts by views",
                 "Ranked across all platforms. Restaurant reviews and hyperlocal Downey stories led — no scandal or viral event in the top slots.")

    posts = sorted(aug_posts, key=lambda p: -int(p.get("view_count") or 0))[:8]

    col_widths = [0.5, 1.1, 6.4, 1.5, 1.3, 1.4]
    headers = ["#", "PLATFORM", "POST", "VIEWS", "LIKES", "CATEGORY"]
    row_y = 2.4
    row_h = 0.5
    x_base = 0.6

    add_rect(slide, x_base, row_y, sum(col_widths), row_h, PALETTE["navy"])
    cx = x_base
    for hd, w in zip(headers, col_widths):
        add_text(slide, cx + 0.1, row_y + 0.12, w - 0.2, 0.3,
                 hd, size=10, color=PALETTE["cream"], bold=True)
        cx += w

    for i, p in enumerate(posts, 1):
        ry = row_y + row_h + (i - 1) * row_h
        if i % 2 == 0:
            add_rect(slide, x_base, ry, sum(col_widths), row_h, PALETTE["soft"])
        if i <= 3:
            add_rect(slide, x_base, ry, 0.06, row_h, PALETTE["gold"])

        title = (p.get("title") or p.get("description_excerpt") or "")[:75]
        plat = {"instagram": "IG", "tiktok": "TT", "facebook": "FB", "youtube": "YT"}.get(p.get("platform"), "?")
        cat = p.get("_category", "other").replace("_", " ").title()

        cx = x_base
        add_text(slide, cx + 0.1, ry + 0.12, col_widths[0] - 0.2, 0.3,
                 str(i), size=12, color=PALETTE["gold"], bold=True)
        cx += col_widths[0]
        add_text(slide, cx + 0.1, ry + 0.12, col_widths[1] - 0.2, 0.3,
                 plat, size=11, color=PALETTE["navy"], bold=True)
        cx += col_widths[1]
        add_text(slide, cx + 0.1, ry + 0.12, col_widths[2] - 0.2, 0.3,
                 title, size=10, color=PALETTE["ink"])
        cx += col_widths[2]
        add_text(slide, cx + 0.1, ry + 0.12, col_widths[3] - 0.2, 0.3,
                 fmt_int(p.get("view_count")), size=12, color=PALETTE["navy"], bold=True)
        cx += col_widths[3]
        add_text(slide, cx + 0.1, ry + 0.12, col_widths[4] - 0.2, 0.3,
                 fmt_int(p.get("like_count")), size=11, color=PALETTE["ink"])
        cx += col_widths[4]
        add_text(slide, cx + 0.1, ry + 0.12, col_widths[5] - 0.2, 0.3,
                 cat[:22], size=9, color=PALETTE["muted"])

    slide_footer_takeaway(slide,
        "The ceiling was lower, but the mix is healthy — food, development, community, judicial all represented.")


def slide_takeaway(prs, jul_posts, aug_posts, since_june_posts):
    top_post = max(aug_posts, key=lambda p: int(p.get("view_count") or 0), default=None)
    top_views = int(top_post.get("view_count") or 0) if top_post else 0
    top_title = ((top_post.get("title") or top_post.get("description_excerpt") or ""))[:60] if top_post else ""
    top_plat = {"instagram": "IG", "tiktok": "TT", "facebook": "FB", "youtube": "YT"}.get(
        (top_post or {}).get("platform"), "?")

    buckets = group_by_category(aug_posts)
    cat_totals = {cat: sum(int(p.get("view_count") or 0) for p in posts)
                  for cat, posts in buckets.items() if posts}
    # Best non-"other" category
    non_other = {c: v for c, v in cat_totals.items() if c != "other"}
    best_cat_key = max(non_other, key=lambda c: non_other[c], default="food")
    best_cat_label = best_cat_key.replace("_", " ").title()
    best_cat_views = non_other.get(best_cat_key, 0)
    best_cat_count = len(buckets.get(best_cat_key, []))

    # Platform: highest avg views/post
    plat_stats = {}
    for p in aug_posts:
        pl = p.get("platform", "?")
        d = plat_stats.setdefault(pl, {"v": 0, "n": 0})
        d["v"] += int(p.get("view_count") or 0)
        d["n"] += 1
    plat_avg = {pl: d["v"] / d["n"] for pl, d in plat_stats.items() if d["n"] > 0}
    strongest_plat = max(plat_avg, key=lambda pl: plat_avg[pl], default="?").title()

    n_scandal_aug = sum(1 for p in aug_posts if p.get("_category") in
                        {"trujillo_scandal", "lisette_scandal", "council_chaos",
                         "ai_fake_news", "ice_immigration"})

    bullets = [
        f"July posts: {len(jul_posts)}  →  August posts: {len(aug_posts)} (volume mostly held)",
        f"Total reach softened −67% — {sum(int(p.get('view_count') or 0) for p in jul_posts):,} → {sum(int(p.get('view_count') or 0) for p in aug_posts):,} views",
        f"#1 August post: [{top_plat}] {fmt_int(top_views)} views — \"{top_title}\"",
        f"Strongest lane (excl. \"other\"): {best_cat_label} — {best_cat_count} posts, {fmt_int(best_cat_views)} total views",
        f"Highest avg reach by platform: {strongest_plat} (unusual — Facebook led in August)",
        f"Scandal posts this month: {n_scandal_aug} (viral engine sat out — no swings taken)",
        f"Cumulative since June 8: {len(since_june_posts)} posts · {fmt_int(sum(int(p.get('view_count') or 0) for p in since_june_posts))} views",
    ]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, LAYOUT_W, LAYOUT_H, PALETTE["ink"])
    add_rect(slide, 0, 4.7, LAYOUT_W, LAYOUT_H - 4.7, PALETTE["gold"])

    add_text(slide, 0.8, 0.6, 11, 0.4,
             "THE READ", size=12, color=PALETTE["gold"], bold=True)
    add_text(slide, 0.8, 1.0, 11.5, 1.0,
             "Powder dry, not powder gone",
             size=36, color=PALETTE["white"], bold=True, font="Georgia")
    add_text(slide, 0.8, 2.15, 11.5, 0.9,
             "August was intentional exploration. No scandal chased, no viral event to ride. "
             "The reels grind held, the catalog grew, and food + community lanes kept per-post "
             "reach comparable to July. The pillar to reclaim in September isn't necessarily "
             "heat — it's any pillar. 57% \"other\" means the show drifted without a spine.",
             size=13, color=PALETTE["cream"])

    for i, b in enumerate(bullets):
        add_text(slide, 1.0, 3.55 + i * 0.22, 11.3, 0.25,
                 "—  " + b, size=11, color=PALETTE["cream"])

    add_text(slide, 0.8, 5.1, 11.5, 0.5,
             "WHAT TO ASK THE SPONSOR", size=12, color=PALETTE["ink"], bold=True)
    add_text(slide, 0.8, 5.5, 11.5, 1.0,
             "Do we pick a September pillar — or keep exploring?",
             size=26, color=PALETTE["white"], bold=True, font="Georgia")


# ── main ────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path,
                    default=Path(__file__).parent / "MONTHLY_REPORT_2026-08.pptx")
    ap.add_argument("--kpis-root", type=Path,
                    default=Path(__file__).parent)
    args = ap.parse_args()

    snaps = load_all_snapshots(args.kpis_root / "episodes_public")
    if not snaps:
        sys.exit(f"No snapshots found under {args.kpis_root/'episodes_public'}")
    print(f"Loaded {len(snaps)} snapshots")

    jul_posts       = posts_in_window(snaps, JULY_START, JULY_END)
    aug_posts       = posts_in_window(snaps, AUGUST_START, AUGUST_END)
    since_june_posts = posts_in_window(snaps, JUNE_START, AUGUST_END)
    print(f"  July window:       {len(jul_posts)} unique posts")
    print(f"  August window:     {len(aug_posts)} unique posts")
    print(f"  Since June 8:      {len(since_june_posts)} unique posts")

    prs = Presentation()
    prs.slide_width = Inches(LAYOUT_W)
    prs.slide_height = Inches(LAYOUT_H)

    slide_title(prs)
    slide_since_june(prs, since_june_posts)
    slide_month_shift(prs, aug_posts, jul_posts)
    slide_platform_growth(prs, snaps)
    slide_categories_august(prs, aug_posts)
    slide_top_posts_august(prs, aug_posts)
    slide_takeaway(prs, jul_posts, aug_posts, since_june_posts)

    prs.save(args.out)
    print(f"✓ Wrote: {args.out}")
    print(f"  {len(prs.slides)} slides")
    print(f"  Open with: open '{args.out}'")


if __name__ == "__main__":
    main()
