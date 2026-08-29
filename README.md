# B2B Intermediary-Referral KPI Report Generator

Pulls deal, meeting and contact data from HubSpot and tracks five
year-to-date stages (Retained Clients, Referred Clients, Presentations,
Calls/Meetings, New Intermediaries) for the Institutional Relations BDMs
(João Pacheco Gonçalves and Rohan Harris).

The primary output today is the **[web dashboard](#web-dashboard)**
(`dashboard/data/kpi-data.json`, refreshed daily via GitHub Actions and
published on Vercel). An Excel-workbook output also exists in this file
(`build_workbook()` and its styling helpers) but is currently **dormant**
- not run automatically, not wired into `main()` - kept in case it's
needed again later; see [Excel workbook (dormant)](#excel-workbook-dormant)
below if you want to regenerate one manually.

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
`dashboard/data/kpi-data.json` cleanly on every run and leaves no state
behind between runs. It prints, in order:

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

### Terminology note

HubSpot's underlying property is genuinely called **Deal Owner**
(`hubspot_owner_id`) and the code, API calls and variable names all use
that name throughout, since that's what the field actually is in the CRM.
Both outputs rename this: every user-facing label says **"BDM"** instead
- a display-only substitution.

## Excel workbook (dormant)

`build_workbook()` and its GCS-styled sheet-builder helpers are still in
`generate_report.py`, fully working, just not called from `main()` any
more and not run by any workflow - the dashboard replaced the workbook as
the primary deliverable, and the two were never dependent on each other
(the dashboard reads `dashboard/data/kpi-data.json` only). To regenerate
`reports/reports.xlsx` manually, call `build_workbook(kpi_data_list,
retained_program_breakdown, ref, today)` yourself (see the git history
before this change, or `main()`'s old structure, for how the pieces fit
together) - one tab per Stage with a Month→Week→Day column drill-down via
Excel's native outline grouping, a Retained Clients row-level drill-down
by Country/Program of Interest, a Calls-Meetings "% of Target" column,
and a static KPI Targets reference tab.

## Web dashboard

`generate_report.py` writes `dashboard/data/kpi-data.json` - the computed
numbers, reshaped for a live web dashboard published on Vercel. This is
the primary output today (the Excel workbook above is dormant).

The dashboard follows the same pattern as the sibling
`gcs-hubspot-funnel-reporting` project: no framework, no server. A small
Python build step inlines the JSON dataset and the GCS logos into three
HTML templates (`dashboard/template_head.html`, `template_body.html`,
`template_js.html`), producing one self-contained `outputs/index.html`.
A second step wraps that file in a client-side, AES-256-GCM password lock
screen (`outputs/vercel/index.html`) - the actual access control, since
Vercel's own deployment protection doesn't cover a bare `*.vercel.app`
alias.

### Building it locally

```bash
python generate_report.py                                    # writes dashboard/data/kpi-data.json
python3 dashboard/build_dashboard.py                          # -> outputs/index.html
python3 dashboard/build_locked.py --password 'your-password'  # -> outputs/vercel/index.html
```

Open `outputs/vercel/index.html` directly in a browser to preview the
lock screen and, once unlocked, the dashboard itself - no server needed.

### Automation and deployment

`.github/workflows/dashboard-data.yml` runs the whole thing daily
(06:00 UTC, plus `workflow_dispatch` for a manual run), commits the
refreshed `dashboard/data/kpi-data.json` and the newly-locked
`outputs/vercel/index.html` back to the repo, and stops there - it does
**not** deploy directly. Instead, connect this repository to a Vercel
project once (Vercel dashboard → Add New Project → import this repo,
framework preset "Other", no build command, output directory
`outputs/vercel`): Vercel's native Git integration then deploys
automatically on every push to `main`, since the push itself is what
lands a new `outputs/vercel/index.html` in the repo.

Required repository secret in addition to the ones above:

| Secret | Purpose |
|---|---|
| `DASHBOARD_PASSWORD` | The password the lock screen decrypts with - share this with whoever needs dashboard access, rotate by changing the secret (next daily run re-locks with it) |

### How it works: HubSpot → GitHub → Vercel

The end-to-end pipeline, for reference or for replicating this same
pattern in another project:

```
HubSpot API
    │  generate_report.py (HUBSPOT_ACCESS_TOKEN secret)
    ▼
dashboard/data/kpi-data.json         ← the one handoff point
    │  dashboard/build_dashboard.py
    ▼
outputs/index.html                   (gitignored, plaintext, local-only)
    │  dashboard/build_locked.py (DASHBOARD_PASSWORD secret)
    ▼
outputs/vercel/index.html            ← force-added, the only file served
    │  git commit + push to main
    ▼
GitHub                                (webhook fires on push)
    │  Vercel's native Git integration, no build
    ▼
Live on Vercel's CDN
```

**HubSpot → GitHub** (`.github/workflows/dashboard-data.yml`, daily cron
at 06:00 UTC plus `workflow_dispatch` for a manual run):

1. Checkout, install dependencies.
2. Run `generate_report.py` with the `HUBSPOT_ACCESS_TOKEN` secret - it
   fetches from HubSpot and writes `dashboard/data/kpi-data.json`. This
   JSON file is the **only** handoff point between "fetch the data" and
   "build the page" - nothing downstream talks to HubSpot directly.
3. Run `dashboard/build_dashboard.py`, then `dashboard/build_locked.py
   --password "$DASHBOARD_PASSWORD"` - together these inline the JSON
   and the GCS logos into one self-contained, password-locked
   `outputs/vercel/index.html`.
4. Commit `dashboard/data/kpi-data.json` and `outputs/vercel/index.html`
   back to `main`. `outputs/` is gitignored (so local test builds don't
   pollute the repo), but this step force-adds
   (`git add -f outputs/vercel/index.html`) that one file specifically -
   it has to be tracked, since landing it in a commit is what triggers
   the next stage.

**GitHub → Vercel**: no deploy script, no `VERCEL_TOKEN`, no API call at
all - this relies entirely on **Vercel's native Git integration**,
configured once, by hand, in the Vercel dashboard (Project Settings):

| Setting | Value |
|---|---|
| Framework Preset | Other |
| Build Command | *(empty)* |
| Install Command | *(empty)* |
| Output Directory | `outputs/vercel`, **override toggle switched on** |

The override toggle is the part most likely to get missed - a typed
Output Directory value is silently ignored if that toggle isn't
explicitly enabled, and Vercel falls back to serving the repo root
instead (this produced a real 404 the first time this project was set
up). With it configured correctly, every push to `main` makes Vercel's
GitHub webhook fire, it deploys near-instantly (there's nothing to
build, just static files to serve), and the new content is live on
Vercel's CDN - typically within a few seconds of the workflow's commit
landing.

**Adapting this for another project**: the reusable idea is the
JSON-in-the-middle handoff - decoupling "fetch your data" from "build
the page" means either half can change independently, and the page-build
step never needs credentials for whatever you're fetching from. The
password-lock step (`build_locked.py`) is optional - skip it entirely if
the other project's data isn't sensitive, or rely on Vercel's own
deployment protection instead if the plan supports it.

### What's on the dashboard

Stage order throughout the dashboard: **Retained, Referred,
Presentations, Calls/Meetings, New Intermediaries** - independent of the
(dormant) Excel workbook's own tab order, via `DASHBOARD_STAGE_ORDER` in
`generate_report.py`.

- **KPI cards** for all 5 stages, split by BDM, for the selected period
  (Year to Date or any individual month); the Calls/Meetings card shows
  each BDM's percentage of their 440/year target alongside the raw count.
- **Funnel by BDM**: a stacked bar per stage showing each BDM's share.
- **Trend**: a line chart, one line per stage, with a **Both / João /
  Rohan** filter above it. Year to Date shows monthly totals across the
  year; selecting a specific month switches the x-axis to that month's
  weeks instead.
- **Stage performance by BDM**: BDM rows (João, Rohan, Grand Total) x
  stage columns, in this order - Retained, Retained Target (`n/12` per
  BDM, `n/24` combined - a pacing indicator, not a formula), Referred,
  Presentations, Calls/Meetings, Calls/Meetings Target (actual/440 with
  a percentage, `x/880` combined), New Intermediaries, Total (the sum of
  the 5 raw count columns for the selected period).
- **Click-through to HubSpot**: every populated Retained or Referred
  number, in the KPI cards, the Funnel bars, and the table, is clickable
  - exactly one matching record opens directly in HubSpot; more than one
  opens a small modal listing each by name, linked to its own record.
  Retained Clients links to the underlying **deal** (a retained client
  always has one); Referred Clients links to the **contact** (a referred
  contact doesn't need a deal to count, so there isn't always one to
  link to).

Numbers refresh once a day.

### Keeping this on-brand and accessible

Any future change to `dashboard/` should be checked against
**`.claude/skills/web-design-guidelines`** (interaction/accessibility
rules - focus states, semantic HTML, `prefers-reduced-motion`, etc.) and
the **`gcs-design-system`** skill's tokens (colors, fonts, spacing) before
being considered done - both were applied deliberately throughout this
dashboard and should stay that way as it evolves.
