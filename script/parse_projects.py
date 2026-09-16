#!/usr/bin/env python3
"""
Parse the daily "Projects Report" export into the JSON the burn-down
dashboard reads.

Usage:
    python parse_projects.py <input.xlsx> <output.json>

How classification works
-------------------------
Business judgment (is a project real "burn" against the client budget, an
added-value freebie, or excluded entirely as pipeline/admin/expense noise)
is NOT inferred from free-text alone. It's driven by a small, explicit set
of rules maintained right here in this script, checked in this order:

  1. FIXED_ADDED_VALUE_IDS — a fixed allowlist of project numbers that are
     always "Added Value" (shown separately, excluded from the burn).
     These don't change often; if BPD delivers a new added-value project,
     add its project number to this set.

  2. Admin/overhead rows — any project number containing "999" (the
     agency's convention for internal/admin project numbers) is excluded.

  3. Expense-only rows — any project whose "Project Type" contains the
     word "expense", or whose "Project Status" is "Expenses Only", is
     excluded (not just "Media: Expense" — any expense type).

  4. Opportunity/pipeline rows — any project whose "Project Status" or
     "Project Type" contains "opportunity" or "oppty" is excluded (not
     yet real, contracted work).

  5. EXCLUDE_TYPES — any remaining project whose "Project Type" exactly
     matches one of these (case-insensitive) is excluded, e.g. media
     labor/management fees that don't count as agency burn.

Everything else in the target campaign counts as real burn: "Active" by
default, or "Completed" if its status text contains "complete". Because
this script only ever reads today's spreadsheet (it never merges in
projects from a previous day), "Completed Projects" on the dashboard is
always exactly what's marked complete in the current Projects Report —
nothing lingers from an earlier export once it drops off the list.

This replaces the older "Dashboard Bucket" spreadsheet-column approach —
the new report format doesn't include that column. If the classification
rule ever needs to change (a new added-value project, a new type/status
to exclude, etc.), update the rules below.
"""
import sys
import os
import glob
import json
import datetime
import openpyxl

TARGET_CAMPAIGN = "2026 UCM Service Lines Campaigns"

# Every run writes a dated snapshot here (in addition to overwriting the
# root data.json), so the dashboard's calendar picker always has something
# to point at even after today's numbers get overwritten tomorrow.
HISTORY_DIR = "data/history"

# These stay fixed here because they're contract terms, not project data.
TOTAL_BUDGET = 2131000
SCOPE_END = "2027-06-30"

# Project numbers that are always "Added Value" — delivered outside the
# client budget, shown separately, excluded from the burn calculation.
FIXED_ADDED_VALUE_IDS = {
    "26-UCMC-034",  # Media Transition & Campaign Builds
    "26-UCMC-061",  # Market Assessment Tool Trial
    "26-UCMC-066",  # Digestive Diseases and Transplant - Media Plan
}

# Project Type values (case-insensitive, exact match) that are excluded
# from the dashboard entirely — media buys/labor, not agency burn.
# (Expense types are also caught more broadly below, regardless of exact
# label, so "media: expense" here is a belt-and-suspenders duplicate.)
EXCLUDE_TYPES = {
    "media: expense",
    "media: labor",
    "media: management fee",
}

# Project numbers containing this are internal/admin/overhead rows, not
# client-billable projects — e.g. "26-UCMC-999" style numbering.
ADMIN_ID_MARKER = "999"


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


def classify(proj_num, proj_type, status):
    """Returns one of: 'active', 'completed', 'added_value', 'exclude'."""
    if proj_num in FIXED_ADDED_VALUE_IDS:
        return "added_value"

    proj_num_key = proj_num or ""
    type_key = (proj_type or "").strip().lower()
    status_key = (status or "").strip().lower()

    # Admin/overhead project numbers.
    if ADMIN_ID_MARKER in proj_num_key:
        return "exclude"

    # Any expense-only row, regardless of exact type label.
    if "expense" in type_key or status_key == "expenses only":
        return "exclude"

    # Opportunity/pipeline rows that aren't real, contracted work yet.
    if "opportunity" in status_key or "oppty" in status_key \
            or "opportunity" in type_key or "oppty" in type_key:
        return "exclude"

    if type_key in EXCLUDE_TYPES:
        return "exclude"

    if "complete" in status_key:
        return "completed"

    return "active"


def main():
    if len(sys.argv) != 3:
        print("Usage: parse_projects.py <input.xlsx> <output.json>", file=sys.stderr)
        sys.exit(1)

    in_path, out_path = sys.argv[1], sys.argv[2]

    wb = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb["Sheet1"]
    # Row 1 is a report title ("Projects Report"), row 2 is the real header.
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    header = list(rows[0])
    data_rows = rows[1:]

    filled = forward_fill_campaign(data_rows, header)

    def get(rec, name):
        return rec[header.index(name)] if name in header else None

    projects = []
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

        proj_num = str(proj_num).strip()
        proj_type = str(get(rec, "Project Type") or "").strip()
        status = str(get(rec, "Project Status") or "").strip()
        budget = float(get(rec, "Current Total Budget") or 0)

        bucket = classify(proj_num, proj_type, status)

        if bucket == "exclude":
            excluded_count += 1
            excluded_budget += budget
            continue

        projects.append(
            {
                "id": proj_num,
                "name": str(proj_name).strip() if proj_name else "",
                "type": proj_type,
                "status": status.title(),
                "budget": budget,
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
        "warnings": [],
    }

    with open(out_path, "w") as f:
        json.dump(out, f, indent=2)

    write_history_and_manifest(out)

    print(f"Wrote {len(projects)} projects to {out_path}")
    print(f"Snapshotted this run to {HISTORY_DIR}/{out['lastUpdated']}.json and rebuilt manifest.json")


if __name__ == "__main__":
    main()
