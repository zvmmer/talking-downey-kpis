#!/bin/bash
# Talking Downey scheduled snapshot — auto-runs Mon/Wed/Fri after 6am.
#
# Triggered by ~/Library/LaunchAgents/com.kaname.talking-downey-snapshot.plist
# on login + at 6am/9am/12pm on Mon/Wed/Fri. Idempotent guardrails prevent
# duplicate runs the same day.
#
# What it does:
#   1. Auto-increments the snapshot tag (t10 → t11 → t12 …)
#   2. Pulls fresh IG + FB + YT + TT via Apify
#   3. Rebuilds the URL registry
#   4. Regenerates July / Overall / Sponsor decks
#   5. Auto-commits + pushes to git for cross-machine sync
#   6. Notifies via macOS notification on success or failure
#
# Log: Projects/talking_downey/kpis/snapshots.log

set +e   # keep going on individual failures so we can send notifications

REPO="/Users/zhamirpascual/talking-downey-kpis"
KPIS="$REPO"
LOG="$KPIS/snapshots.log"
LOCK="/tmp/talking-downey-snapshot.lock"
LAST_RUN="$KPIS/.last_scheduled_run"
EPISODE_DIR="$KPIS/episodes_public/2026-06-08_record_straight"
MANIFEST="$EPISODE_DIR/manifest.yaml"
VENV_PY="$REPO/venv/bin/python"

log() { echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" >> "$LOG"; }

notify_ok() {
    osascript -e "display notification \"$1\" with title \"Talking Downey ✓\" sound name \"Ping\"" 2>/dev/null
}

notify_fail() {
    osascript -e "display notification \"$1 — see snapshots.log\" with title \"Talking Downey ✗ FAILED\" sound name \"Basso\"" 2>/dev/null
    log "FAILURE: $1"
    rm -f "$LOCK"
    exit 1
}

# --- Guardrails ---

# 1. Only Mon/Wed/Fri (%u returns 1-7 for Mon-Sun)
DOW=$(date +%u)
if [[ "$DOW" != "1" && "$DOW" != "3" && "$DOW" != "5" ]]; then
    exit 0   # silent — this is not an error, just not our day
fi

# 2. Only after 6am
HOUR=$(date +%H)
if (( 10#$HOUR < 6 )); then
    exit 0   # silent — too early
fi

# 3. Already ran today? Skip silently
TODAY=$(date +%Y-%m-%d)
if [[ -f "$LAST_RUN" ]] && [[ "$(cat "$LAST_RUN")" == "$TODAY" ]]; then
    exit 0
fi

# 4. Another run in progress? Skip silently (won't happen normally, but safety)
if [[ -f "$LOCK" ]]; then
    log "Skip: lock file exists ($LOCK)"
    exit 0
fi
touch "$LOCK"
trap "rm -f $LOCK" EXIT

# --- Actual work ---

log "======= Snapshot run start ($TODAY, dow=$DOW, hour=$HOUR) ======="
cd "$REPO" || notify_fail "cd to repo failed"

# Find next tag
LAST_T=$(ls "$EPISODE_DIR/snapshots/" 2>/dev/null | grep -Eo '^t[0-9]+' | sort -V | tail -1 | tr -d 't')
LAST_T=${LAST_T:-0}
NEXT_T=$((LAST_T + 1))
TAG="t$NEXT_T"
log "Next tag: $TAG (last was t$LAST_T)"

# Pull fresh data via Apify (IG + FB + YT + TT)
log "Running pull_apify.py..."
if ! "$VENV_PY" "$KPIS/pull_apify.py" --manifest "$MANIFEST" --tag "$TAG" >> "$LOG" 2>&1; then
    notify_fail "pull_apify.py failed for $TAG"
fi

# Rebuild URL registry
log "Rebuilding URL registry..."
if ! "$VENV_PY" "$KPIS/build_url_registry.py" >> "$LOG" 2>&1; then
    notify_fail "build_url_registry.py failed"
fi

# Regenerate all 3 decks
for script in build_july_report.py build_overall_report.py build_sponsor_content_report.py; do
    log "Running $script..."
    if ! "$VENV_PY" "$KPIS/$script" >> "$LOG" 2>&1; then
        notify_fail "$script failed"
    fi
done

# Git commit + push
log "Committing + pushing to origin/main..."
git add "Projects/talking_downey/kpis/" >> "$LOG" 2>&1
if git diff --cached --quiet; then
    log "No changes to commit (already up to date?)"
else
    if ! git commit -m "Scheduled snapshot $TAG ($TODAY): fresh IG/FB/YT/TT + rebuilt decks

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>" >> "$LOG" 2>&1; then
        notify_fail "git commit failed"
    fi
    if ! git push origin main >> "$LOG" 2>&1; then
        notify_fail "git push failed (auth issue? check git-credential-helper)"
    fi
fi

# Mark today done
echo "$TODAY" > "$LAST_RUN"

log "======= Snapshot run complete ($TAG) ======="
notify_ok "Snapshot $TAG pushed. Decks updated."
exit 0
