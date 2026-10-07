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

## Auto audit setup (transcribe + AI-score calls)

The **Auto Audit** tab takes call recording links, gets a transcript from the internal QMS transcription tool (`qms.wiom.in/transcription`), and has an AI model score the call against the 13 scorecard parameters. The result opens in the normal audit form pre-filled; the auditor checks it and submits under their own name, so scoring, Slack alerts, ACPT and every report work exactly as before. Parameters that depend on system data rather than the call audio (Kapture notes, ticket disposition) are left blank for the auditor to rate.

One-time setup:

1. **Update the backend.** Paste the current `Code.gs` into the Apps Script project (the one bound to the audit Sheet), then **Deploy → Manage deployments → pencil icon on the deployment whose URL is in `index.html` → Version: New version → Deploy**.
2. **Add the AI key.** Project Settings (gear icon) → Script Properties → add `ANTHROPIC_API_KEY` with the key value. Optionally add `CLAUDE_MODEL` to pick a different model (default is `claude-sonnet-5-5`). The key stays in Google's Script Properties and never goes into this repo.
3. **Check it.** Open the `/exec` URL with `?test=qms` on the end (should say `"reachable":true`), then with `?test=ai` (should say `"ok":true`). `?test=ai` makes one tiny AI call.

Using Google Gemini instead of Claude:

1. Create an API key at [aistudio.google.com/apikey](https://aistudio.google.com/apikey). A Gemini app subscription (Google AI Pro etc.) does **not** include API usage; the API has its own free tier and its own billing.
2. In Script Properties add `GEMINI_API_KEY` (the key) and `AI_PROVIDER` with the value `gemini`. Optionally add `GEMINI_MODEL` to pick a model (default `gemini-flash-latest`). Set `AI_PROVIDER` back to `anthropic` (or delete it) to return to Claude.
3. Check with `?test=ai` on the `/exec` URL — it reports which provider it used.

Google's terms treat free-tier prompts differently from paid-tier API data (free-tier content may be used to improve their products). Check the current terms before sending customer call transcripts through a free key.

Notes:

- Call transcripts are sent to the AI provider for scoring. Check that your company policy allows this before using real customer calls.
- The QMS tool must be reachable from Google's servers (not office-network-only) because the Apps Script backend is what calls it.
- The AI's scoring is a first draft. The form will not submit until an auditor is chosen and every parameter is rated.
