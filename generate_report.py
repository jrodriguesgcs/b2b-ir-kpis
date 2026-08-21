#!/usr/bin/env python3
"""
B2B Intermediary-Referral KPI report generator.

Fetches deal, meeting and contact data from HubSpot and writes a styled
two-sheet Excel workbook to reports/reports.xlsx:

  - "Weekly Summary": five KPIs, grouped by ISO week and by deal owner
    (the internal HubSpot property `hubspot_owner_id`; rendered in the
    workbook as "BDM").
  - "KPI Targets": static reference content (no API calls).

Run with:  python generate_report.py

Deliberately NOT in scope here: GitHub Actions / cron scheduling, email
delivery, and secrets management beyond a local .env file. See README.md.
"""

from __future__ import annotations

import os
import random
import sys
import time
from dataclasses import dataclass, field

import requests
from dotenv import load_dotenv

HUBSPOT_API_BASE = "https://api.hubapi.com"

# Search API / associations endpoints have a tighter effective rate limit
# (~4-5 req/s) than the general API. A small fixed delay before each such
# call keeps us comfortably under it without needing per-call bookkeeping.
SEARCH_API_DELAY_SECONDS = 0.25

MAX_RETRIES = 5

# The two BDMs (HubSpot deal owners) this report is scoped to.
OWNER_NAMES = {
    "joao": "João Pacheco Gonçalves",
    "rohan": "Rohan Harris",
}


class HubSpotError(RuntimeError):
    """Raised when a HubSpot API call fails after exhausting retries."""


class HubSpotClient:
    """Thin wrapper around requests.Session with auth + backoff baked in."""

    def __init__(self, access_token: str):
        if not access_token:
            raise HubSpotError(
                "HUBSPOT_ACCESS_TOKEN is not set. Copy .env.example to .env "
                "and fill in a HubSpot Service Key or private app token."
            )
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
        )

    def request(self, method: str, path: str, is_search: bool = False, **kwargs):
        """Issue one HubSpot API call with retry/backoff on 429s.

        is_search=True adds a small fixed delay beforehand to respect the
        Search API's stricter rate limit.
        """
        if is_search:
            time.sleep(SEARCH_API_DELAY_SECONDS)

        url = f"{HUBSPOT_API_BASE}{path}"
        last_exc = None
        for attempt in range(MAX_RETRIES + 1):
            try:
                resp = self.session.request(method, url, timeout=30, **kwargs)
            except requests.RequestException as exc:
                last_exc = exc
                if attempt == MAX_RETRIES:
                    break
                delay = (2 ** attempt) + random.uniform(0, 1)
                time.sleep(delay)
                continue

            if resp.status_code == 429:
                if attempt == MAX_RETRIES:
                    raise HubSpotError(
                        f"HubSpot rate limit exceeded after {MAX_RETRIES} retries "
                        f"on {method} {path}"
                    )
                retry_after = resp.headers.get("Retry-After")
                if retry_after is not None:
                    delay = float(retry_after)
                else:
                    delay = (2 ** attempt) + random.uniform(0, 1)
                time.sleep(delay)
                continue

            if resp.status_code >= 400:
                raise HubSpotError(
                    f"HubSpot API error {resp.status_code} on {method} {path}: "
                    f"{resp.text[:500]}"
                )

            return resp.json() if resp.text else {}

        raise HubSpotError(f"HubSpot request failed on {method} {path}: {last_exc}")

    def get(self, path: str, is_search: bool = False, **kwargs):
        return self.request("GET", path, is_search=is_search, **kwargs)

    def post(self, path: str, is_search: bool = False, **kwargs):
        return self.request("POST", path, is_search=is_search, **kwargs)

    def paginate(self, path: str, params: dict | None = None, results_key: str = "results"):
        """GET-paginate a HubSpot list endpoint using the `paging.next.after` cursor."""
        params = dict(params or {})
        results = []
        while True:
            data = self.get(path, params=params)
            results.extend(data.get(results_key, []))
            next_after = data.get("paging", {}).get("next", {}).get("after")
            if not next_after:
                break
            params["after"] = next_after
        return results

    def search_all(self, object_type: str, body: dict):
        """POST-paginate /crm/v3/objects/{type}/search using the `after` cursor."""
        body = dict(body)
        body.setdefault("limit", 100)
        results = []
        while True:
            data = self.post(f"/crm/v3/objects/{object_type}/search", is_search=True, json=body)
            results.extend(data.get("results", []))
            next_after = data.get("paging", {}).get("next", {}).get("after")
            if not next_after:
                break
            body["after"] = next_after
        return results

    def batch_read(self, object_type: str, ids: list[str], properties: list[str]):
        """POST /crm/v3/objects/{type}/batch/read, chunked to 100 IDs/call."""
        results = []
        ids = list(dict.fromkeys(ids))  # de-dupe, preserve order
        for i in range(0, len(ids), 100):
            chunk = ids[i : i + 100]
            body = {
                "properties": properties,
                "inputs": [{"id": obj_id} for obj_id in chunk],
            }
            data = self.post(f"/crm/v3/objects/{object_type}/batch/read", json=body)
            results.extend(data.get("results", []))
        return results


@dataclass
class ReferenceData:
    """Everything Step 0 resolves, cached once per run."""

    typeid_introducer: int
    typeid_referred_person: int
    lead_source_property: str
    lead_source_candidates: list[str]  # all candidates with a "Partner Referral" option
    proposal_signed_property: str
    intermediary_pipeline_id: str
    sales_pipeline_id: str
    proposal_accepted_stage_ids: list[str]
    lifecyclestage_customer_value: str
    owner_ids: dict = field(default_factory=dict)  # {"joao": "123", "rohan": "456"}


def resolve_reference_data(client: HubSpotClient) -> ReferenceData:
    """Step 0: resolve every ID/value the KPI logic needs, print as we go."""
    print("=== Step 0: resolving reference data ===")

    # --- Association labels: contacts <-> contacts (Referred Person / Introducer)
    labels_resp = client.get("/crm/v4/associations/contacts/contacts/labels")
    typeid_introducer = None
    typeid_referred_person = None
    for entry in labels_resp.get("results", []):
        label = (entry.get("label") or "").strip().lower()
        type_id = entry.get("typeId")
        if label == "introducer":
            typeid_introducer = type_id
        elif label == "referred person":
            typeid_referred_person = type_id
    if typeid_introducer is None or typeid_referred_person is None:
        raise HubSpotError(
            "Could not find both 'Referred Person' / 'Introducer' association "
            f"labels on contacts<->contacts. Got: {labels_resp.get('results')}"
        )
    print(f"  TYPEID_INTRODUCER (read from a referred contact) = {typeid_introducer}")
    print(f"  TYPEID_REFERRED_PERSON (inverse)                 = {typeid_referred_person}")

    # --- Contact lead-source property holding "Partner Referral"
    # Multiple overlapping legacy properties can carry this option; KPI 4
    # reconciles all candidates against its expected total (37) rather than
    # committing to a guess here (see fetch_referred_clients).
    lead_source_property, lead_source_candidates = _resolve_lead_source_property(client)
    print(f"  Lead-source property (best guess, reconciled in KPI 4) = {lead_source_property}")
    print(f"  Lead-source candidates carrying 'Partner Referral'     = {lead_source_candidates}")

    # --- Deal property: "Proposal Signed Date Time"
    deal_props = client.get("/crm/v3/properties/deals")
    proposal_signed_property = None
    for prop in deal_props.get("results", []):
        label = (prop.get("label") or "").lower()
        name = (prop.get("name") or "").lower()
        if "proposal" in label and "sign" in label:
            proposal_signed_property = prop["name"]
            break
        if "proposal" in name and "sign" in name:
            proposal_signed_property = prop["name"]
            break
    if not proposal_signed_property:
        raise HubSpotError("Could not resolve 'Proposal Signed Date Time' deal property.")
    print(f"  Deal property 'Proposal Signed Date Time' = {proposal_signed_property}")

    # --- Pipelines / stages
    pipelines_resp = client.get("/crm/v3/pipelines/deals")
    intermediary_pipeline_id = None
    sales_pipeline_id = None
    proposal_accepted_stage_ids = []
    for pipeline in pipelines_resp.get("results", []):
        label = (pipeline.get("label") or "").strip()
        if label == "[GCS] Institutional Relations":
            intermediary_pipeline_id = pipeline["id"]
        if "sales pipeline" in label.lower():
            sales_pipeline_id = pipeline["id"]
            for stage in pipeline.get("stages", []):
                stage_label = (stage.get("label") or "").strip().lower()
                if "proposal accepted" in stage_label or "closed won" in stage_label:
                    proposal_accepted_stage_ids.append(stage["id"])
    if not intermediary_pipeline_id:
        raise HubSpotError("Could not resolve pipeline '[GCS] Institutional Relations'.")
    if not sales_pipeline_id or not proposal_accepted_stage_ids:
        raise HubSpotError("Could not resolve 'Sales Pipeline' / Proposal Accepted/Closed Won stage.")
    print(f"  Pipeline '[GCS] Institutional Relations' id = {intermediary_pipeline_id}")
    print(f"  Pipeline '[GCS] Sales Pipeline' id           = {sales_pipeline_id}")
    print(f"  Stage id(s) Proposal Accepted/Closed Won    = {proposal_accepted_stage_ids}")

    # --- lifecyclestage "Customer" value
    lifecyclestage_prop = client.get("/crm/v3/properties/contacts/lifecyclestage")
    customer_value = None
    for option in lifecyclestage_prop.get("options", []):
        if (option.get("label") or "").strip().lower() == "customer":
            customer_value = option["value"]
            break
    if not customer_value:
        raise HubSpotError("Could not resolve lifecyclestage 'Customer' value.")
    print(f"  lifecyclestage 'Customer' value = {customer_value}")

    # --- Owners
    owners_resp = client.paginate("/crm/v3/owners")
    owner_ids = {}
    for owner in owners_resp:
        full_name = f"{owner.get('firstName', '')} {owner.get('lastName', '')}".strip()
        if full_name == OWNER_NAMES["joao"]:
            owner_ids["joao"] = owner["id"]
        elif full_name == OWNER_NAMES["rohan"]:
            owner_ids["rohan"] = owner["id"]
    missing = [key for key, _ in OWNER_NAMES.items() if key not in owner_ids]
    if missing:
        raise HubSpotError(f"Could not resolve owner id(s) for: {missing}")
    print(f"  Owner id (João Pacheco Gonçalves) = {owner_ids['joao']}")
    print(f"  Owner id (Rohan Harris)           = {owner_ids['rohan']}")

    print("=== Step 0 complete ===\n")

    return ReferenceData(
        typeid_introducer=typeid_introducer,
        typeid_referred_person=typeid_referred_person,
        lead_source_property=lead_source_property,
        lead_source_candidates=lead_source_candidates,
        proposal_signed_property=proposal_signed_property,
        intermediary_pipeline_id=intermediary_pipeline_id,
        sales_pipeline_id=sales_pipeline_id,
        proposal_accepted_stage_ids=proposal_accepted_stage_ids,
        lifecyclestage_customer_value=customer_value,
        owner_ids=owner_ids,
    )


_LEAD_SOURCE_CANDIDATES = ["lead_source", "manual_lead_source", "gc_manual_lead_source"]


def _resolve_lead_source_property(client: HubSpotClient) -> tuple[str, list[str]]:
    """Check each candidate lead-source property for a 'Partner Referral' option
    and report counts so the caller can reconcile against the expected KPI 4/5
    totals rather than assuming a fixed property name. Returns (best_guess,
    all_candidates) - KPI 4 tries every candidate and locks in whichever one
    reconciles to the expected total.
    """
    props_by_name = {p["name"]: p for p in client.get("/crm/v3/properties/contacts").get("results", [])}
    candidates_found = {}
    for name in _LEAD_SOURCE_CANDIDATES:
        prop = props_by_name.get(name)
        if not prop:
            continue
        has_partner_referral = any(
            (opt.get("label") or "").strip().lower() == "partner referral"
            for opt in prop.get("options", [])
        )
        if has_partner_referral:
            count = _count_contacts_with_value(client, name, "Partner Referral", prop)
            candidates_found[name] = count
            print(f"    candidate property '{name}': {count} contacts = Partner Referral")

    if not candidates_found:
        raise HubSpotError(
            "None of the candidate lead-source properties "
            f"{_LEAD_SOURCE_CANDIDATES} have a 'Partner Referral' option."
        )

    # Prefer whichever candidate has the most matching contacts; ties broken by
    # candidate priority order. This is reconciled against KPI 4/5 expected
    # totals during isolated verification (see README / task notes).
    best_name = max(candidates_found, key=lambda n: (candidates_found[n], -_LEAD_SOURCE_CANDIDATES.index(n)))
    return best_name, list(candidates_found.keys())


def _count_contacts_with_value(client: HubSpotClient, property_name: str, label: str, prop: dict) -> int:
    # Search uses the option *value*, not its label.
    value = None
    for opt in prop.get("options", []):
        if (opt.get("label") or "").strip().lower() == label.strip().lower():
            value = opt["value"]
            break
    if value is None:
        return 0
    body = {
        "filterGroups": [
            {"filters": [{"propertyName": property_name, "operator": "EQ", "value": value}]}
        ],
        "properties": ["hs_object_id"],
        "limit": 1,
    }
    data = client.post("/crm/v3/objects/contacts/search", is_search=True, json=body)
    return data.get("total", 0)


def main() -> int:
    load_dotenv()
    access_token = os.environ.get("HUBSPOT_ACCESS_TOKEN", "")
    client = HubSpotClient(access_token)
    resolve_reference_data(client)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except HubSpotError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001 - top-level safety net per spec
        print(f"ERROR: unexpected failure: {exc}", file=sys.stderr)
        sys.exit(1)
