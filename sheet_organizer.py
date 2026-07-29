# sheet_organizer.py
# Rewrites one messy lead tab into the standard column layout
# (TARGET_HEADERS), locating source columns by header aliases.
#
# Safe to re-run on a tab that has already been emailed: existing
# emailed/emailed_date/status/notes values are carried through instead of
# being wiped (wiping them used to make every lead look unsent again).
# The rewrite writes first and clears leftovers after, so a crash can no
# longer empty the tab.

import json

from common import (
    build_header_map, get_cell, get_values, normalize_email, normalize_phone,
    overwrite_tab, sheets_service, split_name, titlecase_name, ZIP_RE,
)

# === CONFIGURE TARGET SHEET HERE ===
TARGET_SHEET_NAME = "Old Vets"   # change this
# ==================================

TARGET_HEADERS = [
    "first_name", "last_name", "phone", "email",
    "age", "address", "city", "state", "zip",
    "emailed", "emailed_date", "followup_sent", "replied",
    "status", "notes", "extras_json"
]


def extract_zip_anywhere(row):
    # Only accept a zip at the END of a cell — '12345 Olive Blvd' must not
    # donate its house number as a ZIP code.
    for c in row:
        s = str(c or "").strip()
        m = ZIP_RE.search(s)
        if m and m.end() == len(s):
            return m.group(0)
    return ""


def organize_one_sheet_by_headers(sheet_name):
    svc = sheets_service()
    rows = get_values(svc, sheet_name)
    if not rows:
        print(f"Nothing found in {sheet_name}.")
        return

    header = rows[0]
    header_map = build_header_map(header)

    if "email" not in header_map and "phone" not in header_map:
        raise SystemExit(
            f"Couldn't find an Email or Phone column from header row in '{sheet_name}'.\n"
            f"Header row was: {header}\n"
            "Add a header like 'Email' or 'Phone', or add its name to ALIASES in common.py."
        )

    def mapped(row, field):
        return str(get_cell(row, header_map.get(field))).strip() if field in header_map else ""

    organized = []
    for r_i, row in enumerate(rows[1:], start=2):
        extras = {
            "source_sheet": sheet_name,
            "source_row": r_i,
            "original_headers": header,
            "raw_row": row,
        }

        first = titlecase_name(mapped(row, "first_name"))
        last = titlecase_name(mapped(row, "last_name"))
        if (not first and not last) and "full_name" in header_map:
            first, last = split_name(mapped(row, "full_name"))

        email = normalize_email(mapped(row, "email"))
        # keep_original: a phone that doesn't normalize to 10 digits is kept
        # as-is (visible for manual fixing) instead of silently blanked.
        phone = normalize_phone(mapped(row, "phone"))

        zipc = mapped(row, "zip") or extract_zip_anywhere(row)

        organized.append([
            first, last, phone, email,
            mapped(row, "age"), mapped(row, "address"),
            mapped(row, "city"), mapped(row, "state"), zipc,
            # Carry send-tracking through re-runs instead of wiping it.
            mapped(row, "email_sent"), mapped(row, "emailed_date"),
            mapped(row, "followup_sent"), mapped(row, "replied"),
            mapped(row, "status"), mapped(row, "notes"),
            json.dumps(extras, ensure_ascii=False),
        ])

    overwrite_tab(svc, sheet_name, TARGET_HEADERS, organized)

    print(f"✅ Organized '{sheet_name}' using header mapping. Rows written: {len(organized)}")
    print(f"Detected columns: {header_map}")


if __name__ == "__main__":
    organize_one_sheet_by_headers(TARGET_SHEET_NAME)
