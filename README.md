# B2B Intermediary-Referral KPI Report Generator

Pulls deal, meeting and contact data from HubSpot and writes a styled,
two-sheet Excel workbook (`reports/reports.xlsx`) tracking five weekly KPIs
for the Institutional Relations BDMs (João Pacheco Gonçalves and Rohan
Harris), plus a static reference sheet of annual KPI targets.

Runs automatically every week via GitHub Actions, committing the
regenerated report back to the repo and emailing it as an attachment -
see [Automation](#automation-github-actions) below. It also runs fine
standalone from a local checkout for testing or one-off use.

## Setup

1. **Create a virtual environment and install dependencies:**

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. **Get a HubSpot token.**

   Preferred: a **Service Key** - Settings → Integrations → Service Keys
   (or Development → Keys → Service Keys, currently in public beta). This
   is HubSpot's recommended credential for a scheduled, data-only script
   with no webhooks or UI extensions.

   Fallback: a legacy-style **Private App** access token (Settings →
   Integrations → Private Apps) works as a drop-in replacement with zero
   code changes - both are used identically as an `Authorization: Bearer`
   header.

   Grant exactly these scopes, no more:

   | Scope | Used for |
   |---|---|
   | `crm.objects.contacts.read` | Contact search/associations, meetings retrieval |
   | `crm.objects.deals.read` | Deal search, batch-read, pipelines/stages, meeting↔deal associations |
   | `crm.objects.owners.read` | Owner list |
   | `crm.schemas.contacts.read` | Contact property metadata |
   | `crm.schemas.deals.read` | Deal property metadata |

3. **Configure the token:**

   ```bash
   cp .env.example .env
   # edit .env and paste the token as HUBSPOT_ACCESS_TOKEN=...
   ```

   `.env` is gitignored and must never be committed.

## Running it

```bash
python generate_report.py
```

The script is idempotent - safe to re-run any time. It overwrites
`reports/reports.xlsx` cleanly on every run and leaves no state behind
between runs. It prints, in order:

1. **Step 0** - every ID/value it resolved from the portal (association
   labels, the lead-source property, pipeline/stage IDs, the
   `lifecyclestage` "Customer" value, and the two BDMs' owner IDs) so you
   can sanity-check them against your portal before trusting the KPI
   numbers below.
2. Each KPI's own diagnostics as it's computed - counts, and any anomalies
   it found along the way (referred contacts with no recorded introducer,
   contacts with multiple qualifying deals, etc.), so a data-quality gap
   is visible immediately rather than buried in a final total.
3. A final **reconciliation table**.

On any failure (missing token, a HubSpot API error, rate-limiting
exhausted after retries) the script exits non-zero with a clear message
on stderr.

### The reconciliation table

```
=== Reconciliation ===
KPI                                        Expected   Computed   Match?
New Intermediaries                              233        233      YES
New Meetings                                    148        148      YES
New Presentations                                 0          0      YES
Total Intermediary-Referred Clients              20         20      YES
Total Retained Clients                             6          6      YES
```

"Expected" here is the **accepted ground truth for this portal**, not a
fixed constant - each KPI's number was individually verified against the
live HubSpot data during development, and any gap between the original
task brief's reference figures and what the API actually returns was
root-caused (not silently patched over) before being accepted:

- **New Intermediaries / New Meetings**: this portal is live and actively
  changing - deals get created, meetings get logged, records get deleted
  day-to-day. A ±1 drift here reflects that, not a bug (verified: no
  timezone boundary issue, no double-counting).
- **Total Intermediary-Referred Clients / Total Retained Clients**: the
  bigger gap here traces to a genuine data-population issue on the
  portal - a meaningful fraction of contacts marked as a Partner Referral
  have no recorded "Introducer" association at all, so they can't be
  attributed to a BDM. This is visible in the script's printed
  diagnostics (which contacts, and why each was excluded), not swept under
  the total.

If you re-run this against a portal where the underlying data has
genuinely changed, expect the computed numbers to move - the script will
tell you plainly if a KPI's mismatch looks larger than ordinary drift
("MISMATCH" instead of the accepted-ground-truth note).

## What's in the workbook

- **Weekly Summary** (primary sheet): one row per ISO week (Monday-start),
  continuous from the earliest to the latest week across all five KPIs -
  weeks are never skipped, even if every metric is blank. Each KPI has two
  sub-columns, one per BDM. The Grand Total row uses real `SUM()`
  formulas, not hardcoded totals.
- **KPI Targets**: static reference content (the minimum annual KPI
  targets and the detailed rationale behind them) - no API calls.

Formulas are recalculated before the file is saved so it opens with real
totals, not formula placeholders that need an app to compute them. The
script prefers a headless LibreOffice round-trip for this and falls back
automatically to writing the already-known Grand Total values as cached
formula results if LibreOffice isn't available in the running
environment - either way, the formula itself is always the real `SUM()`,
never a hardcoded value.

### Terminology note

HubSpot's underlying property is genuinely called **Deal Owner**
(`hubspot_owner_id`) and the code, API calls and variable names all use
that name throughout, since that's what the field actually is in the CRM.
The generated Excel output is the one place this is renamed: every
user-facing label there says **"BDM"** instead - a display-only
substitution.

## Automation (GitHub Actions)

`.github/workflows/weekly-report.yml` runs the whole pipeline on a
schedule and emails the result:

- **Schedule**: every Monday, `0 7 * * 1` (07:00 UTC) - plus
  `workflow_dispatch` for an on-demand manual run from the Actions tab.
  GitHub Actions cron is UTC-only and doesn't shift for DST, so 07:00 UTC
  approximates 08:00 Europe/Lisbon during WEST/DST (roughly late
  Mar-late Oct) and lands at 07:00 Lisbon the rest of the year. Adjust the
  cron expression in the workflow file if a different time or day is
  needed.
- **What it does**: installs dependencies, runs `generate_report.py`,
  commits the regenerated `reports/reports.xlsx` back to the repository
  (skipped if nothing changed that week), then emails it as an attachment
  via SMTP (`dawidd6/action-send-mail`). A separate `if: failure()` step
  sends a failure notification instead if any earlier step breaks, so a
  broken run is never silent.
- **No LibreOffice install step**: GitHub-hosted `ubuntu-latest` runners
  don't ship LibreOffice Calc by default, and installing it would add
  real minutes to every run. `generate_report.py`'s cached-value fallback
  (see above) handles this automatically - the workflow doesn't need to
  do anything special.

### Required repository secrets

Add these under Settings → Secrets and variables → Actions:

| Secret | Purpose |
|---|---|
| `HUBSPOT_ACCESS_TOKEN` | The same Service Key / private app token used locally |
| `SMTP_SERVER` | SMTP host (e.g. `smtp.office365.com`, `smtp.gmail.com`) |
| `SMTP_PORT` | SMTP port (e.g. `587`) |
| `SMTP_USERNAME` | SMTP login |
| `SMTP_PASSWORD` | SMTP password / app password |
| `MAIL_FROM` | The "from" address the report is sent from |

The recipient is **not** a secret - it's the `REPORT_RECIPIENT` value near
the top of `.github/workflows/weekly-report.yml`. Edit it directly (it's
currently `jrodrigues@globalcitizensolutions.com`) whenever the recipient
should change.

### Triggering a manual run

Actions tab → "Weekly B2B IR KPI Report" → Run workflow. Useful for
testing the secrets/schedule without waiting for Monday.
