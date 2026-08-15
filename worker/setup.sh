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

WRANGLER="npx --yes wrangler@latest"

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
  KV_OUT="$($WRANGLER kv namespace create USAGE 2>&1 || $WRANGLER kv:namespace create USAGE 2>&1)"
  echo "$KV_OUT" | tail -n 5
  # The id is a 32-char hex string in the snippet wrangler prints back.
  KV_ID="$(printf '%s' "$KV_OUT" | grep -oE '[0-9a-f]{32}' | head -n 1 || true)"
  [ -n "$KV_ID" ] || die "Could not read the namespace id from wrangler's output. Copy it into worker/wrangler.toml by hand, then re-run."
  node "$ROOT/tools/apply-config.mjs" --kv-id "$KV_ID"
fi

say "3/6  Storing the API keys"
echo "    Paste each key when prompted. Input is hidden and stored encrypted"
echo "    in Cloudflare — never written to this repo."
for KEY in GEMINI_API_KEY OCRSPACE_API_KEY GROQ_API_KEY; do
  case "$KEY" in
    GEMINI_API_KEY)   echo; echo "    $KEY  (aistudio.google.com/apikey — primary chain)";;
    OCRSPACE_API_KEY) echo; echo "    $KEY  (ocr.space/ocrapi/freekey — fallback OCR)";;
    GROQ_API_KEY)     echo; echo "    $KEY  (console.groq.com/keys — fallback LLM)";;
  esac
  $WRANGLER secret put "$KEY"
done

say "4/6  Deploying the Worker"
DEPLOY_OUT="$($WRANGLER deploy 2>&1)"
echo "$DEPLOY_OUT" | tail -n 8
API_URL="$(printf '%s' "$DEPLOY_OUT" | grep -oE 'https://[a-zA-Z0-9._-]+\.workers\.dev' | head -n 1 || true)"
[ -n "$API_URL" ] || die "Deploy finished but no workers.dev URL was found. Copy it from the output above into API_BASE_URL in app.js."
echo "    Live at: $API_URL"

say "5/6  Checking both provider chains"
HEALTH="$(curl -fsS "$API_URL/api/health" || true)"
echo "    $HEALTH"
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
  git push origin main

GitHub Pages redeploys in about a minute, then open the site on a phone
and the deed upload card will be live.

To change a key later:  npx wrangler secret put GEMINI_API_KEY && npx wrangler deploy
------------------------------------------------------------------
EOF
