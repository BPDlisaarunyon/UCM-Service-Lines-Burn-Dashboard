#!/usr/bin/env python3
"""
Parse the "All Projects" export into the JSON the burn-down dashboard reads.

Usage:
    python parse_projects.py <input.xlsx> <output.json>

How classification works
-------------------------
Business judgment (is a project real "burn" against the client budget, an
added-value freebie, or excluded entirely as media/pipeline noise) is NOT
inferred from the "Project Type" text. It comes from a "Dashboard Bucket"
column that a human maintains directly in the source spreadsheet. That
column is the single source of truth — this script just reads it.

Add a column named exactly "Dashboard Bucket" to the export, with one of
these values on every row that belongs to the target campaign:

    Active        -> counts against the budget, shown under Active Projects
    Completed     -> counts against the budget, shown under Completed Projects
    Added Value   -> shown in the Added Value section, EXCLUDED from the burn
    Exclude       -> left off the dashboard entirely (media buys/labor,
                     opportunity/proposal pipeline rows, zero-budget admin
                     rows, etc.)

Rows in the target campaign with a blank/unrecognized bucket are treated
as excluded but are also collected into a "warnings" list in the output
JSON, so the dashboard can surface a "N rows need classification" banner
instead of silently dropping something that matters.
"""
import sys
import os
import glob
import json
import datetime
import openpyxl

TARGET_CAMPAIGN = "2026 UCM Service Lines Campaigns"

# Every run writes a dated snapshot here (in addition to overwriting the
# root data.json), so the dashboard's version dropdown always has something
# to point at even after this week's numbers get overwritten next week.
HISTORY_DIR = "data/history"

# These stay fixed here because they're contract terms, not project data.
# Move them into the spreadsheet too (e.g. a small "Config" sheet) if you'd
# rather not edit this file when they change.
TOTAL_BUDGET = 1865000
SCOPE_END = "2027-06-30"

BUCKET_MAP = {
    "active": "active",
    "completed": "completed",
    "added value": "added_value",
    "added-value": "added_value",
    "addedvalue": "added_value",
    "exclude": "exclude",
    "excluded": "exclude",
}


def forward_fill_campaign(rows, header):
    idx = header.index("Campaign Name")
    last = None
    out = []
    for r in rows:
        val = r[idx]
        if val is not None and str(val).strip() != "":
            last = val
        out.append((last,) + r[1:])
    return out


def month_day_year(date_str):
    """'2026-08-25' -> 'August 25, 2026' (no leading-zero day, cross-platform)."""
    try:
        d = datetime.date.fromisoformat(date_str)
        return f"{d.strftime('%B')} {d.day}, {d.year}"
    except (ValueError, TypeError):
        return date_str or "unknown date"


def write_history_and_manifest(out):
    """Snapshot this run's output into data/history/<date>.json, then rebuild
    data/history/manifest.json from every snapshot found on disk. Rebuilding
    from disk each time (rather than appending to a saved list) means the
    manifest can never drift out of sync with what's actually there."""
    os.makedirs(HISTORY_DIR, exist_ok=True)

    date_str = out["lastUpdated"]
    snapshot_path = os.path.join(HISTORY_DIR, f"{date_str}.json")
    with open(snapshot_path, "w") as f:
        json.dump(out, f, indent=2)

    snapshot_files = sorted(
        (p for p in glob.glob(os.path.join(HISTORY_DIR, "*.json"))
         if os.path.basename(p) != "manifest.json"),
        reverse=True,  # filenames are ISO dates, so this sorts newest first
    )

    versions = []
    for i, path in enumerate(snapshot_files):
        snap_id = os.path.basename(path)[: -len(".json")]
        is_current = i == 0
        versions.append(
            {
                "id": snap_id,
                "label": month_day_year(snap_id),
                # The newest snapshot points at the root data.json (always
                # freshest); older ones point at their own history file.
                "file": "data.json" if is_current else path.replace(os.sep, "/"),
                "current": is_current,
            }
        )

    with open(os.path.join(HISTORY_DIR, "manifest.json"), "w") as f:
        json.dump({"versions": versions}, f, indent=2)


def main():
    if len(sys.argv) != 3:
        print("Usage: parse_projects.py <input.xlsx> <output.json>", file=sys.stderr)
        sys.exit(1)

    in_path, out_path = sys.argv[1], sys.argv[2]

    wb = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    header = list(rows[0])
    data_rows = rows[1:]

    if "Dashboard Bucket" not in header:
        print(
            'ERROR: no "Dashboard Bucket" column found in the export. '
            "Add one with values Active / Completed / Added Value / Exclude "
            "before this script can run.",
            file=sys.stderr,
        )
        sys.exit(2)

    filled = forward_fill_campaign(data_rows, header)

    def get(rec, name):
        return rec[header.index(name)] if name in header else None

    projects = []
    warnings = []
    excluded_count = 0
    excluded_budget = 0.0

    for rec in filled:
        campaign = rec[header.index("Campaign Name")]
        if campaign != TARGET_CAMPAIGN:
            continue

        proj_num = get(rec, "Project Number")
        proj_name = get(rec, "Project Name")
        if proj_num is None and proj_name is None:
            continue  # blank trailing row

        raw_bucket = get(rec, "Dashboard Bucket")
        bucket_key = (str(raw_bucket).strip().lower() if raw_bucket else "")
        bucket = BUCKET_MAP.get(bucket_key)

        if bucket is None:
            warnings.append(
                f"{proj_num} ({proj_name}): missing/unrecognized Dashboard "
                f"Bucket value ({raw_bucket!r}) — excluded until classified."
            )
            continue

        if bucket == "exclude":
            excluded_count += 1
            excluded_budget += float(get(rec, "Current Total Budget") or 0)
            continue

        created = get(rec, "Created Date")
        projects.append(
            {
                "id": str(proj_num),
                "name": str(proj_name).strip() if proj_name else "",
                "type": str(get(rec, "Project Type") or "").strip(),
                "status": str(get(rec, "Project Status") or "").strip().title(),
                "am": str(get(rec, "Account Manager") or "").strip(),
                "budget": float(get(rec, "Current Total Budget") or 0),
                "billed": float(get(rec, "Amount Billed") or 0),
                "lifecycle": "completed" if bucket == "completed" else "active",
                "addedValue": bucket == "added_value",
            }
        )

    out = {
        "campaign": TARGET_CAMPAIGN,
        "totalBudget": TOTAL_BUDGET,
        "scopeEnd": SCOPE_END,
        "lastUpdated": datetime.date.today().isoformat(),
        "projects": projects,
        "excludedCount": excluded_count,
        "excludedBudget": excluded_budget,
        "warnings": warnings,
    }

    with open(out_path, "w") as f:
        json.dump(out, f, indent=2)

    write_history_and_manifest(out)

    print(f"Wrote {len(projects)} projects to {out_path}")
    print(f"Snapshotted this run to {HISTORY_DIR}/{out['lastUpdated']}.json and rebuilt manifest.json")
    if warnings:
        print(f"WARNING: {len(warnings)} row(s) need Dashboard Bucket classification:")
        for w in warnings:
            print("  -", w)


if __name__ == "__main__":
    main()
