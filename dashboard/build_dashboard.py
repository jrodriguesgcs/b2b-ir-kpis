#!/usr/bin/env python3
"""
Assemble the single-file B2B IR KPI dashboard.

Usage:
    python3 build_dashboard.py [-d data/kpi-data.json] [-o ../outputs/index.html]

Inlines the dataset (written by generate_report.py) and the GCS logo SVGs
so the result is one self-contained HTML file, suitable for a static Vercel
deploy. Adapted from the same pattern used by the sibling
`gcs-hubspot-funnel-reporting` dashboard.
"""
import argparse
import base64
import datetime
import json
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).resolve().parent
LOGOS = HERE / "assets" / "logos"


def svg(name):
    p = LOGOS / name
    if not p.exists():
        sys.exit(f"missing logo: {p}")
    s = p.read_text()
    return re.sub(r"<\?xml.*?\?>", "", s, flags=re.S).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--data", default=str(HERE / "data" / "kpi-data.json"))
    ap.add_argument("-o", "--out", default=str(HERE / ".." / "outputs" / "index.html"))
    args = ap.parse_args()

    data_path = pathlib.Path(args.data)
    if not data_path.exists():
        sys.exit(f"missing {data_path} - run generate_report.py first")
    data = json.loads(data_path.read_text())

    parts = [(HERE / f"template_{p}.html").read_text() for p in ("head", "body", "js")]
    html = "\n".join(parts)

    repl = {
        "__DATA__": json.dumps(data, separators=(",", ":")),
        "__LOGO_SECONDARY_WHITE__": svg("GCS-Secondary-White.svg"),
        "__FAVICON__": base64.b64encode((LOGOS / "GCS-Symbol-Blue.svg").read_bytes()).decode(),
        "__YEAR__": str(data.get("meta", {}).get("year", datetime.date.today().year)),
    }
    for k, v in repl.items():
        if k not in html:
            sys.exit(f"placeholder {k} not found in template")
        html = html.replace(k, v)

    left = re.findall(r"__[A-Z_]+__", html)
    if left:
        sys.exit(f"unreplaced placeholders: {set(left)}")

    out = pathlib.Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html)
    print(f"wrote {out} ({len(html) // 1024} KB)")


if __name__ == "__main__":
    main()
