# Apify setup (for Instagram + Facebook scraping)

Apify hosts maintained scrapers ("Actors") for Instagram and Facebook that work without your login. They handle the anti-scraping arms race so you don't have to.

**Cost:** Apify's free tier gives you **$5 of usage per month**. Each scrape run for Talking Downey is roughly $0.01–$0.05. You'll spend pennies for the whole reporting window. Card on file is required for signup but won't be charged unless you blow past the free tier.

---

## Steps (~2 min)

### 1. Sign up

1. Go to **<https://console.apify.com/sign-up>**
2. Sign up with email (any email — doesn't need to be tied to Talking Downey).
3. Verify your email.

### 2. Get your API token

1. In the Apify Console, click your avatar (top-right) → **Settings**.
2. Sidebar → **Integrations** → **API tokens**.
3. Click **Create new token**. Name it `talking-downey-kpis`. Click **Save**.
4. Copy the token (long string starting with `apify_api_…`).

### 3. Save the token locally

```bash
mkdir -p "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z/Projects/talking_downey/kpis/.credentials"
echo "PASTE_YOUR_TOKEN_HERE" > "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z/Projects/talking_downey/kpis/.credentials/apify_token.txt"
```

The `.credentials/` folder is gitignored — token never gets committed.

### 4. Verify it works

```bash
cd "/Users/zhamirpascual/Desktop/Kaname Z/kaname-z"
MCP/.venv/bin/python Projects/talking_downey/kpis/pull_apify.py --check
```

Should print: `✓ Apify token valid. Free tier credits remaining: $X.XX`.

---

## How the script uses it

The `pull_apify.py` script reads the same episode manifest as `pull_public.py`, but pays attention to **profile-level URLs** (not individual posts):

```yaml
apify_targets:
  instagram_profile:  "https://www.instagram.com/talkingdowney/"
  facebook_page:      "https://www.facebook.com/TalkingDowney"
```

You paste the profile/page URL once. Apify pulls the N most recent posts and returns engagement data per post. No per-post URL hunting.

---

## Costs in practice

For the Wednesday pitch (t0 + 24h + 72h = 3 snapshots × 2 platforms = 6 scrape runs):

- Instagram Scraper at 10 posts each = $0.03 × 3 = **~$0.09 total**
- Facebook Posts Scraper at 10 posts each = $0.03 × 3 = **~$0.09 total**
- Grand total: **~$0.18 for the entire pitch window**

Free tier ($5) covers ~100 weeks of weekly tracking before you'd see a charge.

---

## Why this is fine vs. raw-dogging Meta APIs

- You're not creating a developer app that touches the client's accounts.
- You're not asking for any role on the Page.
- You're using publicly visible data, fetched through a third party's infrastructure.
- This is the same data Apify would return if Hans or Mario themselves used Apify.

This is closer to "I subscribed to a media monitoring service" than "I'm reading your DMs." Worth understanding the distinction if it comes up.

That said: **once Hans grants proper Meta access, retire Apify.** Meta Graph API is free, richer, and the proper long-term path. Apify is the bridge for this week.
