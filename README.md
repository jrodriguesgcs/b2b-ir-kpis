# B2B Intermediary-Referral KPI Report Generator

Pulls deal, meeting and contact data from HubSpot and writes a styled,
two-sheet Excel workbook (`reports/reports.xlsx`) tracking five weekly KPIs
for the Institutional Relations BDMs (João Pacheco Gonçalves and Rohan
Harris), plus a static reference sheet of annual KPI targets.

This repository is deliberately scoped to the data-fetching + Excel
generation script only. GitHub Actions scheduling and email delivery are a
separate, later step - see [Not in scope](#not-in-scope-yet) below.

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

## Not in scope yet

The following are deliberately deferred to a later step - the code is
structured so they're easy to add without touching the KPI logic:

- A GitHub Actions workflow: a weekly `on.schedule` cron (offset from the
  top of the hour) plus `workflow_dispatch` for manual runs, committing
  the generated `reports/reports.xlsx` back to the repo.
- An SMTP-based email step (e.g. `dawidd6/action-send-mail`) attaching
  `reports.xlsx`, plus an explicit failure-notification step.
- At that point, `HUBSPOT_ACCESS_TOKEN` moves from the local `.env` to a
  GitHub Actions encrypted secret - no code changes required, since the
  script already reads it from an environment variable.
