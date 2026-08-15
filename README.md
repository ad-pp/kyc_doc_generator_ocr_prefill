# Merchant Onboarding Document Generator

A static web application that lets an on-ground sales agent fill in
merchant/entity details once and generate every required onboarding
document (Partner/Board Resolution, BO Declaration, MDF) as real Word
(`.docx`) files or print-to-PDF — no backend server required.

## Partnership deed upload prefill prototype

This branch adds a **Partnership Firm-only** upload flow that can:

1. Upload a partnership deed/agreement
2. Run OCR through a configurable provider
3. Send OCR text to an LLM for structured extraction
4. Show **suggested** prefills that the agent must review before generation

### No login, no keys on the device

Agents open a URL and start working — there is no sign-in, no app install, and
nothing for them to configure. That is only possible because the OCR and LLM
keys are not on the device at all.

A static site ships its own source, so anything committed to `app.js` is
publicly readable and any key typed into the browser belongs to whoever is
holding the phone. The keys therefore live in a small Cloudflare Worker
(`worker/`), set **once** by the maintainer with `wrangler secret put`. The
browser only knows the Worker's public URL:

```js
const API_BASE_URL = "https://docgen-api.<your-subdomain>.workers.dev";
```

Leave it blank and the deed-upload card is hidden entirely — the app still
generates every document manually, exactly as before.

### Providers — chosen by the backend, not the agent

The Worker decides the order. The client sends one file and gets structured
JSON back; it has no provider setting and no way to influence the choice.

| | Chain | What it does |
|---|---|---|
| **Primary** | Gemini 2.5 Flash (multimodal) | Reads the deed image/PDF directly — OCR and extraction in a single call |
| **Secondary** | OCR.space (Engine 3 → 2) → Groq (Llama 3.3 70B) | A real OCR engine produces page text, then a different LLM extracts from it |

Both are free tiers. The secondary deliberately uses different vendors at both
stages, so one provider's outage or exhausted quota cannot take down both
attempts. If the primary fails, the response carries a warning saying a backup
was used; if both fail, the agent is told to fill the form manually and the
upload is not retried automatically.

**OCR.space engine choice** (per the [API docs](https://ocr.space/ocrapi#ocrengine)):
Engine 3 is the primary — highest accuracy, and it returns tables as Markdown,
which matters because partner name/share tables are exactly what we extract.
Its free quota is 2,500 conversions/month, held separately from the 25,000
shared by Engines 1 and 2. Engine 2 is the retry: the documented all-rounder,
strong on noisy photo backgrounds and rotated text, drawing on the larger pool.
Exhausting the Engine 3 quota therefore degrades automatically rather than
failing. `scale` and `detectOrientation` are on (the API defaults `scale` to
false; the docs note a significant gain on low-resolution scans), and language
is `auto`, which Engines 2/3 support — the Worker substitutes `eng` if a call
ever lands on Engine 1, which has no auto mode.

Free-plan limits that shaped the design: **1 MB per file**, **3 PDF pages**,
25,000 requests/month, and 500 requests/day per IP. The 1 MB cap is why the
browser compresses to ~900 KB rather than the 1.4 MB the primary would happily
accept — otherwise the fallback would fail on exactly the scans that needed it.
The 3-page PDF limit only constrains the fallback; the primary reads full
multi-page PDFs, and a PDF failure surfaces that hint in the error.

The one-call primary is not just simpler — a multimodal model reading the
original scan generally beats OCR-then-LLM on phone-camera photos, where
classical OCR loses text to skew, shadow and low contrast.

### Quota without a login

No sign-in means no user identity, so limits are enforced on what is available:
a per-day cap per agent mobile number (already collected in Step 1) and a
per-day cap per client IP so a missing or mistyped number cannot bypass it.
Both are set in `worker/wrangler.toml` and counted in Workers KV. Over the cap,
the upload returns a clear message; document generation is never blocked.

### Deploying the API Worker (one time)

Full step-by-step instructions, including key insertion, rotation and
monitoring, are in **[RUNBOOK.md](RUNBOOK.md)**. The short version:

```bash
cd worker
npx wrangler login
npx wrangler kv namespace create USAGE     # uncomment the block in wrangler.toml, paste the id
npx wrangler secret put GEMINI_API_KEY     # primary
npx wrangler secret put OCRSPACE_API_KEY   # secondary OCR
npx wrangler secret put GROQ_API_KEY       # secondary LLM
npx wrangler deploy
```

Then set `ALLOWED_ORIGINS` in `wrangler.toml` to the exact origin the app is
served from, and paste the deployed Worker URL into `API_BASE_URL` in `app.js`.
`GET /api/health` reports which chains are configured. Secrets are encrypted at
rest and never appear in the repo, in `wrangler.toml`, or in the browser.
Rotating a key is the same `secret put` command plus a redeploy — no code
change and nothing for agents to do.

A Worker is not a process you keep alive: Cloudflare runs it per request at the
edge, so it is 24×7 from the moment it deploys, with no VM, container or cron
to maintain. The free plan covers 100,000 requests/day and 1,000 KV writes/day.

### Important limits

- Upload prefill is enabled only for **Partnership Firm** flows.
- Only explicit deed facts should be accepted into the form.
- Ambiguous or missing fields must stay blank.
- Document generation still relies on the app's existing validation rules.

## What's in this folder

- `index.html` — the page shell (styles + layout container). Open this file.
- `app.js` — the entire application (state, validation, docx generation,
  print/PDF fallback). Loaded by `index.html` as an ES module.
- `RUNBOOK.md` — go-live steps, key insertion/rotation, monitoring.
- `worker/` — Cloudflare Worker holding the OCR/LLM keys and the
  primary/secondary failover logic. Deployed separately; see above.
- `google-apps-script.gs` — companion script for logging to a Google Sheet.

## How it works

- Everything runs in the browser. Version-pinned document libraries are
  stored under `vendor/` and deployed with the application, so generation
  and preview do not depend on a third-party CDN at runtime. There is no
  Node/Express server to host or maintain.
- Data you enter is auto-saved to `localStorage` as you type, so agents can
  close the tab and resume later without losing progress.
- Two storage buckets are kept separate:
  - **Agent mobile number** — saved permanently until Global Reset is used.
  - **Merchant/entity data** (firm name, partners, resolution/BO details,
    etc.) — cleared only when the agent taps "Start New Merchant".
- Information is captured exactly once and reused across every document in
  the selected set (e.g. firm name, address, partner list, letterhead are
  shared between the Resolution and the BO Declaration).
- Partnership Resolution captures ownership even when generated alone and
  blocks generation unless selected signing partners hold at least 50%
  collectively. The UI shows the selected total and deed/upload guidance.
- Partner Resolution supports Version 5A (all partners are natural persons)
  and Version 5B (one or more non-natural partners). V5A is preselected but
  remains editable. Partnership Resolution captures meeting date only.
- Partnership registration has no default and includes Registrar/deed hints.
- Salesforce Partnership is restricted to MDF-only or the full mandatory
  combination in v5; v6 restores all four individual/combined choices while
  retaining the required mandatory badges. Salesforce Company also exposes
  all four requested choices.
- ACE Company marks Board Resolution, BO Declaration and their combination
  mandatory; the ACE OSV combination remains optional.
- Company Board Resolution allows one selected signer and enforces 100%
  ownership across listed directors/officials; OPC can list one director.
- Standardized optional letterhead uses one bold black line for firm name
  and registered/principal addresses with a purple rule.
- Letterhead uses the legal firm name and main-office address directly;
  duplicate tagline, phone, email and website fields are not required.
- Word documents can be reviewed in the application before download. The
  preview renders the exact generated DOCX blob using lazily loaded,
  version-pinned `docx-preview` and `JSZip` libraries. Microsoft Word may
  paginate a document slightly differently, so the UI labels it as an
  approximate Word preview.
- Section Reset clears one card, Page Reset clears the active page, and
  Global Reset clears all merchant data plus the permanently saved agent
  mobile after confirmation.

## Hosting

Because this is a static site, you can host it anywhere that serves plain
files:

- **GitHub Pages** — push this folder to a repo, enable Pages, done.
- **Any static file host** (Netlify, Cloudflare Pages, S3 + CloudFront,
  internal web server, etc.)
- **Locally** — serve the folder over HTTP using the command below. Browsers
  usually block ES module imports if `index.html` is opened directly as a
  `file://` URL.

No build step, no `npm install`, no server process to keep alive.

## Usage logging

The Sheet was used purely as an append-only log of:

1. Agent's mobile number
2. Merchant Mobile Number (new merchant) or Merchant ID (existing merchant)

<details>
<summary>How the (now suspended) Apps Script logging worked</summary>

**Note on Workspace-restricted deployments:** most corporate Google
Workspace accounts don't offer a fully public "Anyone" access level for
Apps Script Web Apps — only "Anyone within your domain". To work within
that restriction, the app does **not** call the script with `fetch()`.
Instead it submits a hidden HTML `<form>` (a real browser navigation) to a
hidden iframe. A form POST carries the browser's existing Google session
cookie, so on any device/browser already signed into an `@phonepe.com`
Google account (typical on company-managed phones) it logs successfully
with no visible prompt. If the browser has no active Google session, the
hidden iframe silently shows a sign-in page instead — the log entry is
simply skipped that time. Either way, document generation itself is never
blocked or delayed by this.

</details>

**Google Sheets logging is currently suspended.** `SHEETS_URL` in `app.js` is
blank, so nothing is posted to Apps Script. `google-apps-script.gs` is retained
only so logging can be switched back on later — redeploy the web app and paste
the `/exec` URL back into `SHEETS_URL`; no other code change is needed.

### Usage log (current behaviour)

Every generation records one row — timestamp, agent mobile, merchant mobile/ID,
new/existing:

- **Centrally**, posted to the Worker's `/api/log` and kept in Workers KV for
  180 days, when `API_BASE_URL` is configured. No login needed, and the post is
  fire-and-forget so it can never delay a download.
- **On the device**, in `localStorage`, capped at the most recent 1000 rows.
  Step 1 shows the count and offers **Export usage log (CSV)** and **Clear log**,
  which stays useful as an offline fallback.

Extraction calls are logged separately by the Worker, including which provider
chain served them, so primary-versus-fallback rates are visible.

## Document coverage

| Platform   | Entity      | Documents available                              |
|------------|-------------|---------------------------------------------------|
| ACE        | Partnership | Partner Resolution, BO Declaration, or both       |
| ACE        | Company     | Board Resolution, BO Declaration, or both         |
| ACE        | Partnership | Partner Resolution + BO Declaration + ACE OSV/MDF  |
| ACE        | Company     | Board Resolution + BO Declaration + ACE OSV/MDF    |
| Salesforce | Partnership | Partner Resolution, BO Declaration, MDF, or all three |
| Salesforce | Company     | Board Resolution, BO Declaration, MDF, or all three |

For every document you can either:

- **Download .docx** — a real, properly formatted Word file (tables,
  checkboxes, underlined filled-in values, optional letterhead header),
  generated entirely client-side.
- **Download PDF** — generates a clean A4 PDF without browser URL/date
  headers. Traditional Print remains available from Preview.
- **Preview Word** — renders the generated Word file inside the app before
  download. Preview libraries load only when requested to keep initial load
  light on low-end devices.

## Hosting recommendation

GitHub Pages works technically because all document generation happens in
the browser. It has no server runtime to overload as daily usage grows.
However, ordinary GitHub Pages is publicly reachable, offers no application
authentication, and GitHub describes Pages as unsuitable for sensitive
transactions or commercial SaaS-style hosting. A private repository does
not automatically make its Pages URL private. This matters because the app
handles PAN, DOB, residential address and ownership information.

Since the app is used without a login, put it on **Cloudflare Pages** alongside
the API Worker — one vendor, one dashboard, both free:

- GitHub integration and automatic deployment on every push.
- Preview deployments for every branch/pull request.
- Static asset requests are not metered on the free Pages plan.
- Security headers and redirects can be defined in repository files.
- Cloudflare Access can add corporate SSO later if a login is ever wanted
  (free for up to 50 users). It is deliberately **not** enabled here: agents
  work without any sign-in, and usage is instead capped per mobile number and
  per IP in the Worker.

Deployment is change-friendly: connect the same GitHub repository in
Cloudflare Pages, select no framework/build command, and publish the folder
containing `index.html`. Future regulatory changes deploy through the same
normal Git commit/push workflow.

Other alternatives:

- **Firebase Hosting** — GitHub Actions integration and a generous static
  CDN, but free transfer has a monthly cap and requires billing to scale.
- **Azure Static Web Apps** — good for organizations already using Azure and
  Entra ID; the free tier has quotas and no SLA.
- **Netlify** — excellent previews, but its current credit-based free plan is
  less predictable for frequent deployments and high daily traffic.
- **Vercel Hobby** — not recommended because its free tier is intended for
  personal/non-commercial use.

Regardless of host, nothing shipped to the browser is a secret — not URLs and
not API keys. That is why provider keys are entered at runtime and held in
`sessionStorage` rather than committed (see "API keys are never stored in this
repository" above).

If you need to cap who can use the app and how much, a static host alone cannot
do it. On the free tier the practical options are **Cloudflare Pages +
Cloudflare Access** (SSO/one-time-PIN gate in front of the whole site, free for
up to 50 users) or, for counting rather than gating, a **Cloudflare Worker**
(100k requests/day free) that both proxies the AI calls and meters per-user
usage. Both keep the "no server to run" property.

## Data protection

Merchant data, and the local usage log, are stored in browser `localStorage`
so interrupted sessions can resume. This storage is not encrypted. (API keys
are the exception — they go to `sessionStorage` and are dropped when the tab
closes.) Use managed devices, clear each
merchant after document generation, and avoid sharing devices between
agents. For stricter policy, add automatic expiry or disable merchant
persistence before production rollout.

## Editing legal text

All document wording lives in `app.js`:

- `buildPartnershipResolution` / `buildPartnershipBO` — Partnership .docx
- `buildCompanyResolution` / `buildCompanyBO` — Company .docx
- `buildMDF` — MDF .docx

Preview and Print/PDF render the DOCX generated by these same builders, so
there is only one legal-text implementation to update and test.
PDF export isolates each rendered DOCX page into one A4 page, application
styles are scoped away from document headers, and every generated table row
is marked non-splittable to prevent mid-row page breaks.

If your legal/compliance team asks for wording changes, these are the
functions to edit. Keep the `u(...)` wrapper around any value that should
appear bold + underlined (i.e. "this was filled in by the agent").

## Mobile / low-end device notes

- Pure vanilla JS, no framework — small bundle, fast even on low-end
  Android browsers.
- Inputs use large tap targets, `inputmode` hints for numeric fields, and a
  mobile-first responsive layout (single column below 600px).
- Auto-save is debounced (250ms) so typing doesn't feel laggy on slower
  devices.
- No login, no install, no per-agent setup: open the URL and start.
- Deed photos are downscaled to 2000px and re-encoded as JPEG **on the phone**
  before upload — a 9.6 MB camera shot goes out at about 1.4 MB, which matters
  on a 3G connection and keeps a large photo inside the Worker's 8 MB cap.
  Phones too old for `createImageBitmap` fall back to uploading the original.
- Heavy libraries (docx preview, PDF export) are still loaded only on demand.
