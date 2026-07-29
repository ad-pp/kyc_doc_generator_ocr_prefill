# Merchant Onboarding Document Generator

A static web application that lets an on-ground sales agent fill in
merchant/entity details once and generate every required onboarding
document (Partner/Board Resolution, BO Declaration, MDF) as real Word
(`.docx`) files or print-to-PDF — no backend server required.

## What's in this folder

- `index.html` — the page shell (styles + layout container). Open this file.
- `app.js` — the entire application (state, validation, docx generation,
  print/PDF fallback). Loaded by `index.html` as an ES module.
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

## Google Sheet logging (2 data points only)

Per your requirement, the Sheet is used purely as an append-only log of:

1. Agent's mobile number
2. Merchant Mobile Number (new merchant) or Merchant ID (existing merchant)

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

Setup:

1. Create a Google Sheet inside your corporate Google Workspace account.
2. Open **Extensions → Apps Script** and paste in `google-apps-script.gs`.
3. Deploy as a **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone within your organization** (this is fine —
     see the note above; you do not need the fully-public "Anyone" option).
4. Copy the `/exec` URL it gives you.
5. Open `app.js` and paste the URL into the `SHEETS_URL` constant near the
   top of the file:

   ```js
   const SHEETS_URL = "https://script.google.com/macros/s/XXXXXXXXXXXX/exec";
   ```

If you leave `SHEETS_URL` blank, the app works exactly the same — it just
skips logging.

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

For production, **Cloudflare Pages** is the recommended free static host:

- GitHub integration and automatic deployment on every push.
- Preview deployments for every branch/pull request.
- Static asset requests are not metered on the free Pages plan.
- Security headers and redirects can be defined in repository files.
- Cloudflare Access can add corporate SSO if required (its free tier is
  suitable only for a limited user count; verify current enterprise needs).

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

Regardless of host, browser-visible URLs such as `SHEETS_URL` are not
secrets. The domain-restricted Apps Script and corporate Google session are
the current access controls for logging.

## Data protection

Merchant data is stored in browser `localStorage` so interrupted sessions
can resume. This storage is not encrypted. Use managed devices, clear each
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
