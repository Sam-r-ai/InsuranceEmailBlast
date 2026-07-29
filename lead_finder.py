# lead_finder.py
# Search the whole spreadsheet for a lead by FIRST + LAST name.
# Prints which tab the lead is in and the full row (as header: value pairs).
# Works even if columns differ per tab — each tab's own header row is used,
# including tabs that only have a single "Name" / "Full Name" column.

import re

from common import (
    build_header_map, get_cell, get_values, list_sheet_titles, sheets_service,
)

# =========================
# ✅ EDIT THESE SETTINGS
# =========================
SHEETS_TO_IGNORE = [
    "Clients",
    "Potential Clients",
]

# You can also hardcode the name here if you want:
SEARCH_FIRST_NAME = ""   # e.g. "John"
SEARCH_LAST_NAME = ""    # e.g. "Smith"
# =========================


def normalize_name(s):
    s = (s or "").strip().lower()
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return re.sub(r"\s+", " ", s)


def split_full_name(full):
    full = normalize_name(full)
    if not full:
        return "", ""
    parts = full.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def matches_name(row, header_map, first_query, last_query):
    f_idx = header_map.get("first_name")
    l_idx = header_map.get("last_name")
    n_idx = header_map.get("full_name")

    first_val = normalize_name(get_cell(row, f_idx)) if f_idx is not None else ""
    last_val = normalize_name(get_cell(row, l_idx)) if l_idx is not None else ""

    if (not first_val or not last_val) and n_idx is not None:
        f2, l2 = split_full_name(get_cell(row, n_idx))
        first_val = first_val or f2
        last_val = last_val or l2

    return first_val == first_query and last_val == last_query


def row_as_dict(header, row):
    out = {}
    for i, h in enumerate(header):
        key = (h or f"col_{i + 1}").strip()
        out[key] = row[i] if i < len(row) else ""
    return out


def find_lead(first_name, last_name):
    first_q = normalize_name(first_name)
    last_q = normalize_name(last_name)
    if not first_q or not last_q:
        raise SystemExit("Provide BOTH first and last name.")

    svc = sheets_service()
    ignore = {s.strip().lower() for s in SHEETS_TO_IGNORE}

    hits = []
    for sh in list_sheet_titles(svc):
        if sh.strip().lower() in ignore:
            continue

        rows = get_values(svc, sh)
        if not rows or len(rows) < 2:
            continue

        header = rows[0]
        header_map = build_header_map(header)
        if not any(k in header_map for k in ("first_name", "last_name", "full_name")):
            continue

        for row_number, row in enumerate(rows[1:], start=2):
            if matches_name(row, header_map, first_q, last_q):
                hits.append((sh, row_number, header, row))

    return hits


def print_hits(hits, first_name, last_name):
    if not hits:
        print(f"❌ No matches found for: {first_name} {last_name}")
        return

    print(f"✅ Found {len(hits)} match(es) for: {first_name} {last_name}\n")
    for i, (sheet, rownum, header, row) in enumerate(hits, start=1):
        print(f"--- Match #{i} ---")
        print(f"📄 Sheet: {sheet}")
        print(f"📍 Row: {rownum}")
        for k, v in row_as_dict(header, row).items():
            if str(v).strip() != "":
                print(f"  - {k}: {v}")
        print()


if __name__ == "__main__":
    if SEARCH_FIRST_NAME and SEARCH_LAST_NAME:
        f, l = SEARCH_FIRST_NAME, SEARCH_LAST_NAME
    else:
        f = input("First name: ").strip()
        l = input("Last name: ").strip()

    print_hits(find_lead(f, l), f, l)
