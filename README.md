# Quality Audit Dashboard

Source code for the call-quality audit dashboard (weighted QA scorecard, per-agent scoring, filterable audit log, CSV/JSON export).

**Live, shared version:** https://claude.ai/code/artifact/e2b6dc3b-7d2c-47db-842a-40bd276eb1b7

Use the link above for day-to-day auditing — it's a Claude Artifact, and its built-in sync is what lets every auditor see everyone else's submissions in real time.

This repo exists to version-control the source (`index.html`). It is a single self-contained HTML file with no build step. If you host `index.html` elsewhere (GitHub Pages, an internal server, etc.), note that the cross-user data sync will **not** work there — it depends on being served inside the Claude artifact viewer. Each visitor would only see their own local session's data.

To update the live artifact after editing `index.html` here, republish it through Claude Code in the conversation that owns the artifact (or ask Claude to do it for you), pointing at the artifact URL above.
