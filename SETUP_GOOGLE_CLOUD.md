# One-time Google Cloud Console setup

You do this **once**. Takes ~5 min. It gives you a `client_secrets.json` file that authenticates the KPI scripts against YouTube on your behalf.

After this is done, the auth script handles everything automatically — refresh tokens auto-renew, no further browser flows needed.

---

## Steps

### 1. Make a Google Cloud project

1. Go to **<https://console.cloud.google.com/>**
2. Sign in with the Google account that has access to the **Talking Downey** YouTube channel.
3. Top bar → project selector → **New Project**
4. Name it `talking-downey-kpis` (any name works). Click **Create**.
5. After it creates, make sure that project is selected in the top bar.

### 2. Enable the two APIs we need

1. Left sidebar → **APIs & Services** → **Library**
2. Search **YouTube Data API v3** → click → **Enable**
3. Back to Library, search **YouTube Analytics API** → click → **Enable**

### 3. Configure the OAuth consent screen

This is the screen you'll see during the auth flow — it says "this app wants access to your YouTube data."

1. Left sidebar → **APIs & Services** → **OAuth consent screen**
2. **User Type**: **External** → Create
3. **App information**:
   - App name: `Talking Downey KPIs`
   - User support email: your email
   - Developer contact: your email
4. **Audience**: skip
5. **Scopes**: skip (we'll request scopes at runtime)
6. **Test users**: click **+ Add Users** → add the Google account you'll use to sign in. **You must do this** or auth will fail with "app is in testing mode."
7. **Summary**: Save.

### 4. Create OAuth credentials

1. Left sidebar → **APIs & Services** → **Credentials**
2. Top button → **+ Create Credentials** → **OAuth client ID**
3. **Application type**: **Desktop app**
4. Name: `talking-downey-cli`
5. **Create**.
6. A modal appears with your client ID + secret. Click **Download JSON**.
7. Save the downloaded file as **exactly** this path:
   ```
   /Users/zhamirpascual/Desktop/Kaname Z/kaname-z/Projects/talking_downey/kpis/.credentials/client_secrets.json
   ```

   The `.credentials/` folder is gitignored — credentials never get committed.

### 5. Run the auth script

From the repo root, in Terminal:

```bash
cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
MCP/.venv/bin/python Projects/talking_downey/kpis/auth.py
```

What happens:
- A browser tab opens.
- Sign in with the Google account you added as a test user in Step 3.
- It says "Google hasn't verified this app" (because we're in test mode — that's expected). Click **Advanced** → **Go to Talking Downey KPIs (unsafe)**.
- Grant the requested scopes (read YouTube channel data + read analytics).
- Tab redirects to a localhost page that says "auth complete — you can close this tab."
- Terminal shows: `✓ Token saved. Channel: Talking Downey (UC…). All set.`

After that, `pull_youtube.py` and every other KPI script can run without prompts.

---

## Common gotchas

- **"This app isn't verified" / "access blocked"** — you skipped Step 3.6 (Test users). Go back and add yourself.
- **"client_secrets.json not found"** — the file isn't at the exact path above, OR it's named differently. Check capitalization and the folder.
- **Wrong Google account** — make sure you sign in with the account that owns/manages the Talking Downey channel.
- **Refresh token expired** — happens after ~6 months of inactivity in test mode. Re-run `auth.py` once and it re-issues.

---

## What scopes we're requesting

- `youtube.readonly` — list channels, videos, video metadata (title, duration, posted_at).
- `yt-analytics.readonly` — watch time, views, retention, traffic sources, demographics.

Both are **read-only**. The script literally cannot modify your channel even if you wanted it to.
