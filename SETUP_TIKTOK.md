# TikTok setup — two paths

TikTok's API story is genuinely worse than YouTube and Meta. Two practical options:

| Path | Effort | Data quality | Recommended? |
|---|---|---|---|
| **(A) Manual CSV export** from TikTok Creator dashboard, script ingests it | 2 min/week | Excellent — full Creator analytics | ✅ Yes for now |
| **(B) TikTok Display API** for automated pulls | Hours + approval wait | Limited — basic public counts | Skip unless A becomes painful |

The honest recommendation: **start with (A)**. You already log in to TikTok every week to verify posts went up per the SOP. Adding a CSV download to that flow is 60 seconds. Automating it would save those 60 seconds but cost weeks of approval back-and-forth with TikTok.

If TikTok becomes a primary platform later (not a clip-distribution side channel), we revisit (B).

---

## Path A — Manual CSV export (recommended)

### Weekly workflow

1. **Log into <https://www.tiktok.com/tiktokstudio>** (TikTok Studio — Creator analytics dashboard).
   - This requires a **Pro account** (Creator or Business). If you're on Personal, switch in Settings → Account → Manage account.

2. **Top menu → Analytics → Overview** (or **Content → All videos** for per-video stats).

3. **Top right → Download data** → choose `Last 7 days` (for the 1-week report) → CSV.

4. **Save the CSV** to:
   ```
   /Volumes/MAIN DRIVE/Talking Downey/POSTING DATA/tiktok_exports/YYYY-MM-DD.csv
   ```
   The script picks up the most recent file by date.

5. **Run the ingester:**
   ```bash
   cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
   MCP/.venv/bin/python Projects/talking_downey/kpis/ingest_tiktok_csv.py
   ```

   It parses the CSV, normalizes to the common KPI shape, saves per-video JSON next to the YouTube + Meta data so reports can combine all three platforms.

### For the 24h/72h SOP checkpoints

TikTok's CSV export only goes back as far as the dashboard allows — usually a week is the smallest window. So for the 24h/72h snapshots, the practical move is:

- **Open TikTok Studio after 24h** and **72h** post-publish.
- **Note the video's** views, likes, comments, shares in the per-episode `meta.yaml` (created by the per-episode workflow we'll build next).
- The 1-week CSV catches everything in detail.

Yes, this is partly manual. The math: 2 minutes per checkpoint × 3 checkpoints × 4 episodes/month = 24 min/month of TikTok-specific manual work. Faster to do that than wait 3 weeks for TikTok Display API approval that might still not get you per-video analytics.

---

## Path B — TikTok Display API (skip unless you really want it)

Only worth doing if:
- TikTok becomes your primary platform (unlikely for a podcast retainer)
- You have ~4 hours to spend on the setup
- You're OK with the API only giving you **basic public counts**, not the full analytics from TikTok Studio

### Steps (high level)

1. Go to **<https://developers.tiktok.com/>** → register as a developer (free).
2. Create an app, choose **TikTok for Developers** (NOT TikTok for Business — that one requires advertiser account).
3. Add the **Display API** product (NOT Login Kit alone).
4. Configure scopes: `user.info.basic`, `video.list`.
5. Set redirect URI to `http://localhost:8765/callback`.
6. Submit for app review (yes, even for personal use). Approval: days to weeks.
7. After approval, OAuth flow works similar to YouTube: browser → consent → token → refresh.

If you want me to build the auth code for Path B as a placeholder you can use later, say the word — but I'd rather not write code that might not work for weeks. Better to commit to A now.

---

## What you get from each path

| Metric | (A) CSV from Studio | (B) Display API |
|---|---|---|
| Views | ✅ | ✅ |
| Likes | ✅ | ✅ |
| Comments | ✅ | ✅ |
| Shares | ✅ | ✅ |
| Watch time / retention | ✅ | ❌ |
| Traffic sources | ✅ | ❌ |
| Audience demographics | ✅ | ❌ |
| Profile actions | ✅ | ❌ |

Path A actually has **more** data than Path B. The only thing Path B gives you that Path A doesn't is automation. For 60 seconds of weekly clicks, you keep richer analytics.
