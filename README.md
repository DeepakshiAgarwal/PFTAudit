# PFT Quality Audit

Call-quality audit tool: weighted QA scorecard, per-agent scoring, reason dropdowns, category-wise reporting, CSV/JSON export, and an instant feedback-email draft on submit.

This is now a normal client-server web app — a static page (`index.html`) plus a small backend (`Code.gs`, a Google Apps Script Web App backed by a Google Sheet). It works from any browser, any Google account, no Claude login involved, so data sync is not tied to any particular org/account the way the old Claude-artifact version was.

## One-time setup (owner only, ~5 minutes)

1. **Create the backend:**
   - Go to [sheets.google.com](https://sheets.google.com), create a new blank spreadsheet (name it something like "PFT Audit Data").
   - Extensions → Apps Script. Delete the placeholder code and paste in the contents of `Code.gs` from this repo.
   - Click **Deploy → New deployment**. For "Select type," choose **Web app**.
   - Set **Execute as: Me**, **Who has access: Anyone**.
   - Click Deploy, authorize the permissions it asks for (it only touches this one spreadsheet), and copy the resulting URL (ends in `/exec`).
2. **Wire the URL in:** send that URL to whoever maintains this repo (or edit `index.html` yourself) — replace the `API_URL` placeholder near the top of the `<script>` block with it, then commit and push.
3. **Host the page:** in this repo's GitHub Settings → Pages, set Source to "Deploy from branch," branch `main`, folder `/ (root)`. GitHub gives you a public URL like `https://<username>.github.io/PFTAudit/` — that's the one link everyone uses, auditors and management alike, no invites or org restrictions needed.

## Day to day

- Share the GitHub Pages URL with your team. Anyone who opens it can submit audits and browse the Dashboard tab — no login required.
- The top-left badge shows **Synced** (talking to the Sheet fine), **Offline** (couldn't reach it — data stays on that device until you click Refresh), or **Not set up** (the API_URL step above hasn't been done yet).
- Every audit is stored as one row in the "Audits" sheet in your Google Sheet — you can open that sheet directly any time to see the raw data or build your own charts on top of it.
- The **Bulk import** box on the Dashboard tab accepts a pasted JSON array of audits (matching the "Copy as JSON" format) and adds any that aren't already present — handy for one-off migrations.

## Migrating the old data

The previous Claude-artifact version had 4 audits in it. Once this is deployed, open that old artifact, go to Dashboard → Copy as JSON, and paste the result into the **Bulk import** box here to bring them across.
