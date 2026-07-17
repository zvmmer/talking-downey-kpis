# Talking Downey — KPI Tracker

Per-episode performance tracking across all posting platforms, against the SOP's 24h / 72h / 1-week checkpoints.

---

## Platform setup (one-time per platform)

| Platform | Status | Setup guide | Auth script | Puller |
|---|---|---|---|---|
| **YouTube** (long + Shorts) | Ready | [SETUP_GOOGLE_CLOUD.md](SETUP_GOOGLE_CLOUD.md) | `auth.py` | `pull_youtube.py` |
| **Meta** (FB pages + Instagram) | Ready | [SETUP_META.md](SETUP_META.md) | `auth_meta.py` | `pull_meta.py` |
| **TikTok** (CSV ingest) | Ready | [SETUP_TIKTOK.md](SETUP_TIKTOK.md) | — | `ingest_tiktok_csv.py` |

**All three are free** — no paid APIs, no credit card required for any of them. YouTube is the easiest, Meta is medium pain, TikTok is the lightest because the CSV path skips API approval entirely.

### Suggested order to get going

1. **YouTube first.** Smallest UI to navigate, biggest signal in the report. Run `auth.py` → run `pull_youtube.py` → see real data. ~10 min.
2. **Meta next.** 30–45 min of UI clicking, but unlocks 4 of your platforms at once (Talking Downey FB + Downey Beat FB + IG main + IG Reels). Run `auth_meta.py` → run `pull_meta.py`.
3. **TikTok last.** Just download weekly CSVs and run `ingest_tiktok_csv.py`. No auth dance.

---

## Quick reference — running each puller

### YouTube
```bash
cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
MCP/.venv/bin/python Projects/talking_downey/kpis/pull_youtube.py           # most recent video
MCP/.venv/bin/python Projects/talking_downey/kpis/pull_youtube.py --last 5  # last 5
```

### Meta (FB + IG)
```bash
MCP/.venv/bin/python Projects/talking_downey/kpis/pull_meta.py              # most recent per page
MCP/.venv/bin/python Projects/talking_downey/kpis/pull_meta.py --last 5
```

### TikTok
```bash
# 1. Download CSV from TikTok Studio → Analytics → Download data
# 2. Save to: /Volumes/MAIN DRIVE/Talking Downey/POSTING DATA/tiktok_exports/
# 3. Run:
MCP/.venv/bin/python Projects/talking_downey/kpis/ingest_tiktok_csv.py
```

---

## Files

```
kpis/
├── README.md                ← this file
├── requirements.txt         ← Python deps (Google + Meta + TikTok)
│
├── SETUP_GOOGLE_CLOUD.md    ← Step-by-step Google Cloud Console setup
├── auth.py                  ← YouTube OAuth flow
├── pull_youtube.py          ← YouTube stats puller
│
├── SETUP_META.md            ← Step-by-step Meta for Developers setup
├── auth_meta.py             ← Meta token exchange
├── pull_meta.py             ← FB pages + Instagram puller
│
├── SETUP_TIKTOK.md          ← TikTok CSV vs. API decision tree
├── ingest_tiktok_csv.py     ← TikTok CSV → normalized JSON
│
├── .credentials/            ← gitignored
│   ├── client_secrets.json      ← YouTube OAuth client
│   ├── token.json               ← YouTube refresh token
│   ├── channel.json             ← YouTube channel ID/title cache
│   └── meta_token.json          ← Meta long-lived token + page tokens
├── episodes/                ← per-episode KPI data (created by per-episode workflow)
└── reports/                 ← generated client-deliverable reports (Phase 2)
```

---

## Roadmap

- [x] YouTube auth + puller
- [x] Meta (FB + IG) auth + puller
- [x] TikTok CSV ingester
- [ ] **Phase 2:** per-episode workflow — `new_episode.py` to scaffold a folder + `meta.yaml` when you post
- [ ] **Phase 2:** snapshot triggers — pull at 24h, 72h, 1wk
- [ ] **Phase 2:** report generator → client-ready markdown combining all platforms
- [ ] **Phase 3:** LaunchAgent auto-scheduling so you don't have to remember the snapshot times
- [ ] **Phase 3:** Apple Podcasts + Spotify ingestion (optional)

---

## What's NOT in the auto-pipeline (and why)

- **RSS / Apple Podcasts / Spotify**: each has its own dashboard, very limited APIs. Easier to note the subscriber count in the weekly report by hand than build 3 more auth flows.
- **TikTok per-video API automation**: skipped in favor of weekly CSV — saves the multi-week TikTok app review and you get richer data anyway.
- **Mario's Preview**: you said skip.

If any of these become priorities, they can be added later — same pattern as the existing pullers.

---

## Cost

**$0** across all platforms. No credit card needed anywhere. The quotas are huge relative to a weekly podcast:

| Platform | Free quota | Talking Downey usage | Headroom |
|---|---|---|---|
| YouTube | 10k units/day | ~3 units/episode/snapshot × 3 snapshots/episode × 1 ep/week = ~9 units/week | 99.9% spare |
| Meta | Per-app rate limit (~200 calls/hour per app) | ~20 calls/episode/snapshot | Comfortable |
| TikTok | Manual CSV — no API | — | — |

You'd have to scale to hundreds of episodes per week to bump against any limit.
