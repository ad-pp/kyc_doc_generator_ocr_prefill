#!/usr/bin/env bash
# One-command setup. Run from the repo root:  bash worker/setup.sh
#
# You will be asked for exactly three things: the Gemini, OCR.space and Groq
# API keys. Everything else — Cloudflare login, KV namespace, config edits,
# deploy, health check — is handled here.
#
# Safe to re-run: it updates in place rather than duplicating anything.

set -euo pipefail

cd "$(dirname "$0")/.."
ROOT="$PWD"

# Guard the most likely first-run mistake: cloning the default branch, which
# does not contain this directory, and getting an opaque "No such file" error.
[ -f "$ROOT/tools/apply-config.mjs" ] || {
  printf "\n\033[31mX This does not look like the right branch.\033[0m\n" >&2
  printf "  worker/ and tools/ live on claude/ocr-llm-key-security-aaa59u until PR #1 merges.\n" >&2
  printf "  Run:  git checkout claude/ocr-llm-key-security-aaa59u\n\n" >&2
  exit 1
}

say()  { printf "\n\033[1m==> %s\033[0m\n" "$1"; }
warn() { printf "\033[33m    %s\033[0m\n" "$1"; }
die()  { printf "\n\033[31mX %s\033[0m\n" "$1" >&2; exit 1; }

command -v node >/dev/null || die "Node.js 18+ is required. Install it from nodejs.org and re-run."
NODE_MAJOR="$(node -p 'process.versions.node.split(".")[0]')"
[ "$NODE_MAJOR" -ge 18 ] || die "Node.js 18+ is required (found $(node -v))."

# Overridable so the failure paths can be exercised with a stub.
WRANGLER="${WRANGLER:-npx --yes wrangler@latest}"

say "1/6  Signing in to Cloudflare"
if $WRANGLER whoami >/dev/null 2>&1; then
  echo "    Already signed in."
else
  warn "A browser window will open. Approve the login, then come back here."
  $WRANGLER login
fi

say "2/6  Creating the KV namespace (quota counters + usage log)"
cd "$ROOT/worker"
if grep -q '^\[\[kv_namespaces\]\]' wrangler.toml; then
  echo "    Already configured — skipping."
else
  # Assigning from a failing command substitution under `set -e` exits the
  # script before anything gets printed, hiding the very error we need. Run it
  # inside an `if`, which suspends `set -e`, and always show the output.
  KV_OUT=""
  if ! KV_OUT="$($WRANGLER kv namespace create USAGE 2>&1)"; then
    if ! KV_OUT="$($WRANGLER kv:namespace create USAGE 2>&1)"; then
      echo "$KV_OUT"
      warn "Could not create the KV namespace (full output above)."
      warn "The Worker runs fine without it — you lose the per-day caps and the"
      warn "central usage log, both of which can be added later."
      printf "    Continue without KV? [y/N] "
      # Prefer the terminal (stdin may be a pipe), but never hard-fail on it.
      ANSWER="n"
      read -r ANSWER </dev/tty 2>/dev/null || read -r ANSWER 2>/dev/null || ANSWER="n"
      case "$ANSWER" in
        [yY]*) KV_OUT="";;
        *) die "Stopped. Paste the output above and we can work out the cause.";;
      esac
    fi
  fi
  if [ -n "$KV_OUT" ]; then
    echo "$KV_OUT" | tail -n 5
    # The id is a 32-char hex string in the snippet wrangler prints back.
    KV_ID="$(printf '%s' "$KV_OUT" | grep -oE '[0-9a-f]{32}' | head -n 1 || true)"
    if [ -n "$KV_ID" ]; then
      node "$ROOT/tools/apply-config.mjs" --kv-id "$KV_ID"
    else
      warn "Namespace created but no id found in the output above."
      warn "Paste the id into the kv_namespaces block in worker/wrangler.toml, then re-run."
    fi
  fi
fi

say "3/6  Storing the API keys"
echo "    Paste each key when prompted. Input is hidden and stored encrypted"
echo "    in Cloudflare — never written to this repo."
# Existing secrets are left alone so a re-run doesn't ask for everything again.
EXISTING_SECRETS="$($WRANGLER secret list 2>/dev/null || true)"
for KEY in GEMINI_API_KEY OCRSPACE_API_KEY GROQ_API_KEY; do
  case "$EXISTING_SECRETS" in
    *"$KEY"*)
      echo; echo "    $KEY  already stored — skipping."
      echo "    (change it later with: npx wrangler secret put $KEY)"
      continue;;
  esac
  case "$KEY" in
    GEMINI_API_KEY)   echo; echo "    $KEY  (aistudio.google.com/apikey — primary chain)";;
    OCRSPACE_API_KEY) echo; echo "    $KEY  (ocr.space/ocrapi/freekey — fallback OCR)";;
    GROQ_API_KEY)     echo; echo "    $KEY  (console.groq.com/keys — fallback LLM)";;
  esac
  if ! $WRANGLER secret put "$KEY"; then
    warn "Storing $KEY failed. Re-run this script, or set it later with:"
    warn "  npx wrangler secret put $KEY"
  fi
done

say "4/6  Deploying the Worker"
# Stream through tee rather than capturing into a variable: the output stays
# visible as it happens, and a failure is readable without extra handling.
DEPLOY_LOG="$(mktemp "${TMPDIR:-/tmp}/docgen-deploy.XXXXXX")"
DEPLOY_OK=0
$WRANGLER deploy 2>&1 | tee "$DEPLOY_LOG" || DEPLOY_OK=1

if grep -qE 'register a workers\.dev subdomain|workers\.dev subdomain before publishing' "$DEPLOY_LOG"; then
  ACCOUNT_URL="$(grep -oE 'https://dash\.cloudflare\.com/[a-f0-9]+/workers/onboarding' "$DEPLOY_LOG" | head -n 1 || true)"
  printf "\n\033[33m    This Cloudflare account has no workers.dev subdomain yet.\033[0m\n"
  printf "    Register one (free, once, ~30 seconds):\n\n"
  printf "      %s\n\n" "${ACCOUNT_URL:-https://dash.cloudflare.com/ -> Workers & Pages -> set up a subdomain}"
  printf "    Then re-run this script. Your keys are already stored and will be kept.\n\n"
  rm -f "$DEPLOY_LOG"
  exit 1
fi
[ "$DEPLOY_OK" -eq 0 ] || { rm -f "$DEPLOY_LOG"; die "Deploy failed (full output above)."; }

API_URL="$(grep -oE 'https://[a-zA-Z0-9._-]+\.workers\.dev' "$DEPLOY_LOG" | head -n 1 || true)"
rm -f "$DEPLOY_LOG"
if [ -z "$API_URL" ]; then
  warn "The Worker uploaded but has no public URL (the dashboard shows"
  warn "\"No URLs enabled\"). Confirm workers_dev = true in wrangler.toml,"
  warn "then re-run. Enabling it needs a deploy, not just a dashboard toggle."
fi
[ -n "$API_URL" ] || die "Deploy finished but no workers.dev URL was found. Copy it from the output above into API_BASE_URL in app.js."
echo "    Live at: $API_URL"

if ! grep -q '^\[\[kv_namespaces\]\]' "$ROOT/worker/wrangler.toml"; then
  warn "No KV namespace is bound: per-day caps and the central usage log are off."
  warn "Extraction and document generation are unaffected. To add it later:"
  warn "  npx wrangler kv namespace create USAGE"
  warn "  node tools/apply-config.mjs --kv-id <id> && npx wrangler deploy"
fi

say "5/6  Checking both provider chains"
# A workers.dev route enabled for the first time needs a few seconds to
# propagate, and returns 404 until it does. Retry before believing the answer.
HEALTH=""
for ATTEMPT in 1 2 3 4 5 6; do
  HEALTH="$(curl -fsS --max-time 10 "$API_URL/api/health" 2>/dev/null || true)"
  case "$HEALTH" in
    *'"ok"'*) break;;
  esac
  [ "$ATTEMPT" -lt 6 ] || break
  echo "    Route not answering yet, retrying ($ATTEMPT/6)..."
  sleep 5
done
echo "    ${HEALTH:-<no response>}"
case "$HEALTH" in
  "") warn "No response from $API_URL/api/health."
      warn "A brand-new workers.dev route can take a minute. Check again with:"
      warn "  curl -i $API_URL/api/health";;
esac
case "$HEALTH" in
  *'"primary":true'*) ;;
  *) warn "Primary chain is NOT configured — re-run and check the Gemini key.";;
esac
case "$HEALTH" in
  *'"secondary":true'*) ;;
  *) warn "Secondary chain is NOT configured — re-run and check the OCR.space and Groq keys.";;
esac

say "6/6  Wiring the app to the Worker"
node "$ROOT/tools/apply-config.mjs" --api-url "$API_URL"

cat <<EOF

------------------------------------------------------------------
Setup complete.

  Worker : $API_URL
  App    : app.js now points at it; cache version bumped.

Last step — publish the app (this is the only git command you need):

  git add app.js index.html worker/wrangler.toml
  git commit -m "Point app at deployed API Worker"
  git push origin \$(git rev-parse --abbrev-ref HEAD)

GitHub Pages deploys from main, so merge PR #1 after pushing. About a minute
later, open the site on a phone and the deed upload card will be live.

To change a key later:  npx wrangler secret put GEMINI_API_KEY && npx wrangler deploy
------------------------------------------------------------------
EOF
