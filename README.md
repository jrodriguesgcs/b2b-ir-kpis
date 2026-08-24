# B2B Intermediary-Referral KPI Report Generator

Pulls deal, meeting and contact data from HubSpot and writes a styled,
two-sheet Excel workbook (`reports/reports.xlsx`) tracking five
year-to-date stages (Retained Clients, Referred Clients, New
Intermediaries, Presentations, Calls/Meetings) for the Institutional
Relations BDMs (João Pacheco Gonçalves and Rohan Harris), with a
click-to-expand Month → Week → Day drill-down, plus a static reference
sheet of annual KPI targets.

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
   can sanity-check them against your portal before trusting the numbers
   below.
2. Each stage's own diagnostics as it's computed - counts, and any
   anomalies it found along the way (referred contacts with no recorded
   introducer, contacts with multiple qualifying deals, meetings already
   booked for a future date, etc.), so a data-quality gap is visible
   immediately rather than buried in a final total.
3. A final **Year-to-Date Summary** - one line per stage with its current
   grand total.

On any failure (missing token, a HubSpot API error, rate-limiting
exhausted after retries) the script exits non-zero with a clear message
on stderr.

### Reading the numbers

Every stage is scoped to **this calendar year, through today** - a
meeting already booked for next week, or a deal signed last year, is
excluded from the current total (the script tells you when it excludes
something like that, and why). Because of this, the totals move day to
day, and don't chase a single fixed "expected" figure the way an
all-time report would. Two things worth knowing if a number looks lower
than you expected:

- **Retained Clients** in particular can look small this early in the
  year - it only counts deals actually signed within the current year, so
  a deal signed last year for a client referred last year won't show up
  even though that client is genuinely retained.
- Any KPI-population gap the script finds (e.g. a Partner-Referral
  contact with no recorded introducer association) is logged with the
  specific contact/deal IDs responsible, not silently dropped.

## What's in the workbook

- **Year to Date** (primary sheet): **Stage** is the top-level column
  grouping, in this order: Retained Clients, Referred Clients, New
  Intermediaries, Presentations, Calls/Meetings. Each Stage's columns
  drill down Month → Week → Day using Excel's native **column outline
  grouping** - only the Month-total columns are visible by default;
  click the **+** control above a Month to reveal its Week-total columns,
  and click **+** again above a Week to reveal its individual Days. This
  works in Excel, LibreOffice Calc, and Google Sheets (via File → Import
  or opening the .xlsx directly) - look for small **+ / −** buttons in
  the grey bar just above the column headers. Rows are the two BDMs plus
  a Grand Total row. Every Week/Month total is a real `SUM()` formula
  over its own Day/Week columns, never a hardcoded number.
- **KPI Targets**: static reference content (the minimum annual KPI
  targets and the detailed rationale behind them, reordered to match the
  Year to Date sheet's Stage order) - no API calls.

An in-progress month or week is never padded with future placeholder
columns - it simply has fewer Day columns than a finished one, which is
what makes its total read as "month-to-date" / "week-to-date" without any
special-case logic.

Formulas are recalculated before the file is saved so it opens with real
totals, not formula placeholders that need an app to compute them. The
script prefers a headless LibreOffice round-trip for this and falls back
automatically to writing the already-known values as cached formula
results if LibreOffice isn't available in the running environment -
either way, the formula itself is always the real `SUM()`, never a
hardcoded value.

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
