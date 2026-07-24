# Merchant Onboarding Document Generator

A single static HTML page that lets an on-ground sales agent fill in
merchant/entity details once and generate every required onboarding
document (Partner/Board Resolution, BO Declaration, MDF) as real Word
(`.docx`) files or print-to-PDF — no backend server required.

## What's in this folder

- `index.html` — the page shell (styles + layout container). Open this file.
- `app.js` — the entire application (state, validation, docx generation,
  print/PDF fallback). Loaded by `index.html` as an ES module.
- `google-apps-script.gs` — companion script for logging to a Google Sheet.

## How it works

- Everything runs in the browser. `docx.js` (the same library used to
  generate Word documents) is loaded from a CDN as an ES module — there is
  no Node/Express server to host or maintain.
- Data you enter is auto-saved to `localStorage` as you type, so agents can
  close the tab and resume later without losing progress.
- Two storage buckets are kept separate:
  - **Agent mobile number** — saved *permanently* on the device. Never
    cleared automatically.
  - **Merchant/entity data** (firm name, partners, resolution/BO details,
    etc.) — cleared only when the agent taps "Start New Merchant".
- Information is captured exactly once and reused across every document in
  the selected set (e.g. firm name, address, partner list, letterhead are
  shared between the Resolution and the BO Declaration).

## Hosting

Because this is a static site, you can host it anywhere that serves plain
files:

- **GitHub Pages** — push this folder to a repo, enable Pages, done.
- **Any static file host** (Netlify, Cloudflare Pages, S3 + CloudFront,
  internal web server, etc.)
- **Locally** — you can even just double-click `index.html` on a phone or
  laptop and it will work (docx.js is fetched from the CDN, so an internet
  connection is required at least once per session to load it).

No build step, no `npm install`, no server process to keep alive.

## Google Sheet logging (2 data points only)

Per your requirement, the Sheet is used purely as an append-only log of:

1. Agent's mobile number
2. Merchant Display Name (new merchant) or Merchant ID (existing merchant)

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
| Salesforce | Partnership | Resolution, BO Declaration, MDF, or all three     |
| Salesforce | Company     | Board Resolution, BO Declaration, MDF, or all three |

For every document you can either:

- **Download .docx** — a real, properly formatted Word file (tables,
  checkboxes, underlined filled-in values, optional letterhead header),
  generated entirely client-side.
- **Print / PDF** — opens the browser print dialog so the agent can save
  directly as a PDF if that's more convenient on a particular device.

## Editing legal text

All document wording lives in `app.js`:

- `buildPartnershipResolution` / `buildPartnershipBO` — Partnership .docx
- `buildCompanyResolution` / `buildCompanyBO` — Company .docx
- `buildMDF` — MDF .docx
- `printable*` functions — the plain-HTML equivalents used by Print/PDF

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
