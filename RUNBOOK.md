# Go-live runbook

How this runs 24×7, how to deploy it, and how to insert or rotate keys.

---

## 1. Where it runs

Two pieces, neither of which is a server you keep alive.

| Piece | Runs on | Availability |
|---|---|---|
| Static app (`index.html`, `app.js`, `vendor/`) | GitHub Pages today; Cloudflare Pages optional later | Files on a CDN. Nothing to start or restart. |
| API (`worker/`) | Cloudflare Workers | Runs per request at the edge. Idle costs nothing and there is no instance to crash. |

**There is no VM, no container, no laptop that must stay on, and no cron job.**
A Cloudflare Worker is not a process you boot — Cloudflare runs your code when a
request arrives, in whichever data centre is nearest the agent, and it is
already 24×7 the moment `wrangler deploy` finishes. This is why the free tier
is workable: you are billed per request, and there is no idle charge to avoid.

Current live URL: `https://ad-pp.github.io/kyc_doc_generator_ocr_prefill/`
(deployed by `.github/workflows/static.yml` on every push to `main`).

### Do I need to move off GitHub Pages?

No. Keep it. It already works, agents already have the URL, and the Worker is
callable from any origin you allowlist. Migrating to Cloudflare Pages is a
"one vendor, one dashboard" convenience, not a requirement — do it later if you
want preview deployments per branch or a custom domain.

---

## 2. Accounts and keys you need

Create these first; all are free and none need a card.

| # | Account | Gives you | Where |
|---|---|---|---|
| 1 | Cloudflare | Workers + KV | dash.cloudflare.com/sign-up |
| 2 | Google AI Studio | `GEMINI_API_KEY` — primary chain | aistudio.google.com/apikey |
| 3 | OCR.space | `OCRSPACE_API_KEY` — secondary OCR (free plan is enough) | ocr.space/ocrapi/freekey |
| 4 | Groq | `GROQ_API_KEY` — secondary LLM | console.groq.com/keys |

The old Gemini key was exposed in git history and has been revoked. Generate a
**new** one at step 2 — do not reuse it.

> **Note for later, not for the POC:** free-tier LLM APIs generally reserve the
> right to use submitted content to improve their services. Deeds carry partner
> names, addresses and ownership shares. Fine while proving the concept; revisit
> before a real rollout by moving to a paid tier or a Google Cloud project,
> where that reuse does not apply. No code change is needed for that — only a
> different key.

---

## 3. Deploy — the one-command path

From the repo root, with Node 18+ installed:

```bash
git clone -b claude/ocr-llm-key-security-aaa59u \
  https://github.com/ad-pp/kyc_doc_generator_ocr_prefill.git
cd kyc_doc_generator_ocr_prefill
bash worker/setup.sh
```

> **The `-b` matters.** This work lives on the branch
> `claude/ocr-llm-key-security-aaa59u` until PR #1 is merged. A plain
> `git clone` gives you `main`, which has no `worker/` directory yet, and
> `bash worker/setup.sh` fails with *No such file or directory*. Already
> cloned without it? Just switch:
>
> ```bash
> git checkout claude/ocr-llm-key-security-aaa59u
> ```
>
> Once PR #1 is merged, drop the `-b` and use `main`.

You will be asked for exactly three things — the Gemini, OCR.space and Groq
keys — plus one browser approval for the Cloudflare login. The script does the
rest: creates the KV namespace and writes its id into `wrangler.toml`, stores
each key as an encrypted secret, deploys the Worker, checks both provider
chains, then points `app.js` at the deployed URL and bumps the cache version.
It is safe to re-run.

It finishes by printing the one `git push` needed to publish the app.

The manual equivalent is below, for when something needs doing by hand.

### Manual steps

Run everything from a machine with Node 18+ installed. You do this **once**;
agents never do any of it.

### Step 1 — Get the code

```bash
git clone -b claude/ocr-llm-key-security-aaa59u \
  https://github.com/ad-pp/kyc_doc_generator_ocr_prefill.git
cd kyc_doc_generator_ocr_prefill/worker
```

### Step 2 — Log in to Cloudflare

```bash
npx wrangler login
```

Opens a browser once and stores the session locally.

### Step 3 — Create the KV namespace

```bash
npx wrangler kv namespace create USAGE
```

It prints an `id`. Open `worker/wrangler.toml`, **uncomment** the
`[[kv_namespaces]]` block and paste the id in.

> On older wrangler the command is `wrangler kv:namespace create USAGE`
> (with a colon). Both do the same thing.

This namespace holds the daily quota counters and the central usage log. Skip
it and extraction still works — you just lose caps and central logging.

### Step 4 — Insert the keys

```bash
npx wrangler secret put GEMINI_API_KEY
npx wrangler secret put OCRSPACE_API_KEY
npx wrangler secret put GROQ_API_KEY
```

Each prompts for the value and hides it as you paste. Secrets are encrypted at
rest, are not readable back from the dashboard, never enter git, and never
reach the browser. This is the one-time integration — agents are never asked
for a key.

### Step 5 — Set the allowed origin

In `worker/wrangler.toml`, confirm:

```toml
ALLOWED_ORIGINS = "https://ad-pp.github.io"
```

Scheme and host only — no path, no trailing slash. Anything else calling the
Worker gets no CORS grant. Add more origins comma-separated if you also deploy
to Cloudflare Pages.

### Step 6 — Deploy the Worker

```bash
npx wrangler deploy
```

It prints the live URL, e.g. `https://docgen-api.<your-subdomain>.workers.dev`.
That URL is live worldwide from this moment.

### Step 7 — Check it

```bash
curl https://docgen-api.<your-subdomain>.workers.dev/api/health
```

Expect:

```json
{"ok":true,"primary":true,"secondary":true}
```

`false` anywhere means that chain's key is missing — redo step 4 for it.

### Step 8 — Point the app at the Worker

In `app.js`, near the top:

```js
const API_BASE_URL = "https://docgen-api.<your-subdomain>.workers.dev";
```

Bump the cache-buster in `index.html` (`app.js?v=9` → `v=10`) so phones with
the old file cached pick up the change.

### Step 9 — Ship it

```bash
git add app.js index.html worker/wrangler.toml
git commit -m "Point app at deployed API Worker"
git push origin claude/ocr-llm-key-security-aaa59u
```

Then merge PR #1. GitHub Pages deploys from `main`, so the app goes live about
a minute after the merge, not after this push.

### Step 10 — Verify on an actual phone

1. Open the site on an Android handset.
2. Enter an agent mobile, pick **ACE → Partnership Firm**, choose a document set.
3. On Step 4 the **Partnership Deed Upload Prefill** card should now appear —
   if it is missing, `API_BASE_URL` is still blank or the deploy has not landed.
4. Photograph a real deed page and upload it.
5. Confirm suggestions appear, then **Apply Suggested Prefill**.
6. Generate a document and confirm the download works.

Then confirm the log recorded it:

```bash
cd worker && npx wrangler kv key list --binding USAGE | grep '"log:'
```

---

## 4. Rotating or updating a key

Same command as insertion — it overwrites:

```bash
cd worker
npx wrangler secret put GEMINI_API_KEY
npx wrangler deploy
```

No code change, no app redeploy, no agent action. Rotation is invisible to the
field: the next request simply uses the new key.

To retire a provider entirely:

```bash
cd worker
npx wrangler secret delete GROQ_API_KEY
npx wrangler deploy
```

`/api/health` will then report `secondary:false`, and the Worker will only
attempt the primary chain.

**If a key ever leaks:** revoke it at the provider first, then `secret put` the
replacement. Revoking first is what actually stops the abuse; rotating without
revoking leaves the old key valid.

---

## 5. Day-to-day operation

**Watching it**

```bash
cd worker && npx wrangler tail     # live request log
```

Every `wrangler` command reads `worker/wrangler.toml` for the Worker name and
its bindings, so run them from `worker/` — or name the Worker explicitly, e.g.
`npx wrangler tail docgen-api`, which works from anywhere.

The dashboard (Workers → docgen-api) shows requests, errors and CPU time. Each
extraction is logged with the chain that served it, so a rising share of
`ocrspace-groq` means the primary is degraded or out of quota.

**Reading the usage log**

```bash
cd worker
npx wrangler kv key list --binding USAGE
npx wrangler kv key get --binding USAGE "log:2026-08-15T..."
```

Agents can also export their own device log as CSV from Step 1.

**Limits to keep in mind**

| Limit | Value | What happens at the edge |
|---|---|---|
| Worker requests | 100,000/day (free) | Well beyond a field team's volume |
| KV writes | 1,000/day (free) | Each extraction writes ~3 keys; roughly 300 extractions/day |
| Per agent | 40 extractions/day | Configurable in `wrangler.toml` |
| Per IP | 120 extractions/day | Stops one connection burning the pool |
| Upload | 8 MB | Photos are compressed on the phone first |
| OCR.space free plan | 1 MB/file, 3 PDF pages, 25,000/month, 500/day per IP | Engine 3 has its own 2,500/month quota; exhausting it retries on Engine 2 |
| Gemini / Groq free tiers | Provider-set, and they change | Exhausting the primary triggers automatic failover to the secondary |

Raising a cap is a `wrangler.toml` edit plus `npx wrangler deploy`.

**If extraction breaks entirely** — both chains down, or a bad deploy — set
`API_BASE_URL = ""` in `app.js` and push. The upload card disappears and agents
carry on filling forms manually. Document generation has never depended on the
API.

---

## 6. Local testing before deploying

```bash
cd worker
cp .dev.vars.example .dev.vars     # fill in test keys; gitignored
npx wrangler dev                   # serves on http://localhost:8787
```

Point `API_BASE_URL` at `http://localhost:8787`, add
`ALLOWED_ORIGINS = "http://localhost:8000"`, and serve the site with
`python3 -m http.server 8000`. Revert both before committing.
