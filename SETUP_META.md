# One-time Meta (Facebook + Instagram) setup

Meta's Graph API covers **both Facebook pages and Instagram** in one auth — so this single setup unlocks:

- Talking Downey FB page
- Downey Beat FB page
- Talking Downey Instagram (main feed)
- Talking Downey Instagram (Reels)

Setup is ~30–45 min. UI changes monthly so steps may look slightly different — what stays constant is the **flow**: create an App → enable products → generate a User Access Token → exchange it for a long-lived (60-day) token → derive page access tokens.

---

## Prerequisites

- **Personal Facebook account** that's an admin of both Talking Downey FB + Downey Beat FB pages.
- **Instagram account** for Talking Downey must be a **Creator** or **Business** account (not Personal). Convert it in IG Settings → Account Type if needed.
- The **Instagram account must be linked to a Facebook Page**. In FB Page Settings → Instagram → Connect. (Linking to the Talking Downey FB Page is fine.)

---

## Steps

### 1. Create a Meta for Developers app

1. Go to **<https://developers.facebook.com/>**, sign in with the personal FB account.
2. Top menu → **My Apps** → **Create App**.
3. **What do you want your app to do?** → choose **Other**.
4. **Select an app type** → **Business**.
5. App name: `talking-downey-kpis`. Contact email: yours. **Business account**: leave empty if asked, or pick yours. **Create app.**

### 2. Add the products you need

You're now in the App Dashboard. In the left sidebar → **Add products to your app**.

Add these two:
- **Facebook Login for Business** → click **Set up**.
- **Instagram** → click **Set up** → choose **Instagram Graph API** (NOT Basic Display — Basic Display was deprecated).

### 3. Get your App ID + App Secret

Sidebar → **App settings** → **Basic**.

- Copy the **App ID** (visible).
- Click **Show** next to **App Secret**, enter your FB password, copy it.

Save both — you'll paste them into the script in Step 6.

### 4. Add yourself as a tester (so the app works without going Live)

Sidebar → **App Roles** → **Roles** → click **Add People** → **Tester** → enter your Facebook username/email.

This means your account can use the app while it's still in Development mode. You don't need to go through App Review for personal use.

### 5. Generate a User Access Token with the right permissions

1. Go to **<https://developers.facebook.com/tools/explorer/>** (Graph API Explorer).
2. Top right: **Meta App** dropdown → pick your `talking-downey-kpis` app.
3. **User or Page** dropdown → **User Token**.
4. Click **Add a Permission** and check ALL of these:
   - `pages_show_list`
   - `pages_read_engagement`
   - `pages_read_user_content`
   - `read_insights`
   - `instagram_basic`
   - `instagram_manage_insights`
5. Click **Generate Access Token** → log in if prompted → grant.
6. **Copy the token** (long string starting with `EAA…`).

This token is **short-lived** (~1 hour). The next step exchanges it for a 60-day token.

### 6. Run the auth script

Paste the three values (App ID, App Secret, short-lived token) when prompted:

```bash
cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
MCP/.venv/bin/python Projects/talking_downey/kpis/auth_meta.py
```

What it does:
- Exchanges your short-lived token for a **long-lived user token** (~60 days).
- Lists every FB page you admin.
- For each page, derives a **page access token** (these don't expire as long as the user token is refreshed).
- For each page, detects whether an Instagram Business/Creator account is linked.
- Saves everything to `.credentials/meta_token.json` (gitignored).

Output looks like:
```
✓ Long-lived user token saved (expires in 60 days).
Pages found:
  ✓ Talking Downey         | FB page  | IG: @talkingdowney
  ✓ Downey Beat            | FB page  | IG: (not linked)
```

### 7. Token refresh reminder

The long-lived token lasts ~60 days. The auth script auto-refreshes if there's >7 days left on the token AND a script runs. If you go more than 60 days without running any KPI script, the token expires and you'll need to re-do Step 5 + Step 6.

If that becomes friction, we can set up a LaunchAgent that pings the API once a week just to keep the token alive.

---

## Common gotchas

- **"Instagram account not linked"** — your IG must be Creator/Business AND linked to a FB page via the Page's settings. Recheck under FB Page Settings → Instagram.
- **"This object doesn't exist"** when pulling insights — Meta APIs return cryptic errors for permission gaps. Re-check Step 5 — all six permissions must be granted on the User Token.
- **"App ID/Secret mismatch"** — most likely typo. Copy directly from the dashboard, don't retype.
- **Token expired** — re-run Step 5 + Step 6 once, you're back in business.

---

## Scopes used (read-only)

- `pages_show_list` — list pages you admin
- `pages_read_engagement` — page-level engagement
- `pages_read_user_content` — read user content on your pages
- `read_insights` — page insights
- `instagram_basic` — IG account profile + media list
- `instagram_manage_insights` — IG post insights

All read-only. The scripts cannot post, edit, or delete anything.
