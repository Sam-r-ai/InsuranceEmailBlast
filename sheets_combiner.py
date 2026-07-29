# sheets_combiner.py
# Combines specific tabs from ONE Google Spreadsheet into a single output tab.
# - You hardcode which tabs to combine + output column order at the top
# - Source columns are located by header ALIASES (so "Phone", "Cell" and
#   "number" all map to the number column), not by exact name only
# - Duplicate leads (same email or phone appearing in multiple tabs) are
#   written once
# - email_sent history already on the output tab is carried forward, so
#   re-running the combiner can never cause already-emailed leads to be
#   blasted again
# - If every source tab is empty the run aborts WITHOUT touching the output

from common import (
    ALIASES, build_header_map, ensure_tab, get_cell, get_values,
    normalize_email, normalize_header, normalize_phone, overwrite_tab,
    sheets_service,
)

# =========================
# ✅ EDIT THESE SETTINGS
# =========================
SOURCE_SHEETS = [
    "TTC 12/30",
    "NEW TTC",
    "TTC HIGH INTENT",
]

OUTPUT_SHEET_NAME = "Combined TTC 12/30"

# ✅ Output column order (edit this anytime). The tracking columns stay at
# the end so blast/follow-up history survives a re-combine.
OUTPUT_HEADERS = [
    "first_name",
    "last_name",
    "number",
    "email",
    "notes",
    "beneficiary",
    "state",
    "hobby",
    "coverage",
    "fe",
    "email_sent",
    "followup_sent",
    "replied",
]
# =========================

TRACKING_HEADERS = ["email_sent", "followup_sent", "replied"]

# Map each output header to a canonical ALIASES field so source columns can
# be found under any recognizable name.
OUTPUT_FIELD_FOR_HEADER = {
    "first_name": "first_name",
    "last_name": "last_name",
    "number": "phone",
    "email": "email",
    "state": "state",
    "email_sent": "email_sent",
    "followup_sent": "followup_sent",
    "replied": "replied",
}


def dedupe_key(email, phone):
    if email:
        return f"e:{email}"
    if phone:
        return f"p:{phone}"
    return None


def load_existing_sent_map(svc, tab_title):
    """email/phone -> {tracking header: value} from the CURRENT output tab."""
    sent_map = {}
    rows = get_values(svc, tab_title)
    if not rows or len(rows) < 2:
        return sent_map
    hmap = build_header_map(rows[0])
    email_idx, phone_idx = hmap.get("email"), hmap.get("phone")
    tracking_idx = {h: hmap.get(OUTPUT_FIELD_FOR_HEADER[h]) for h in TRACKING_HEADERS}
    if all(i is None for i in tracking_idx.values()):
        return sent_map
    for r in rows[1:]:
        values = {h: str(get_cell(r, i)).strip()
                  for h, i in tracking_idx.items() if i is not None}
        if not any(values.values()):
            continue
        e = normalize_email(get_cell(r, email_idx))
        p = normalize_phone(get_cell(r, phone_idx), keep_original=False)
        key = dedupe_key(e, p)
        if key:
            sent_map[key] = values
    return sent_map


def combine():
    svc = sheets_service()

    existing_title = ensure_tab(svc, OUTPUT_SHEET_NAME)
    sent_map = load_existing_sent_map(svc, existing_title)
    if sent_map:
        print(f"🕘 Carrying forward email_sent history for {len(sent_map)} lead(s).")

    all_rows_out = []
    seen = set()
    duplicates = 0

    for tab in SOURCE_SHEETS:
        rows = get_values(svc, tab)
        if not rows or len(rows) < 2:
            print(f"⚠️ Skipping '{tab}' (empty or no data).")
            continue

        header = rows[0]
        header_map = build_header_map(header)
        # Columns with no alias mapping (notes, hobby, fe, ...) fall back to
        # exact normalized-name matching.
        exact_map = {normalize_header(h): i for i, h in enumerate(header)}

        col_for_output = {}
        for out_h in OUTPUT_HEADERS:
            field = OUTPUT_FIELD_FOR_HEADER.get(out_h)
            idx = header_map.get(field) if field else None
            if idx is None:
                idx = exact_map.get(normalize_header(out_h))
            col_for_output[out_h] = idx

        missing_critical = [h for h in ("email", "number")
                            if col_for_output.get(h) is None]
        if missing_critical:
            print(f"⚠️ '{tab}': no column found for {missing_critical} — "
                  "those cells will be blank.")

        added = 0
        for r in rows[1:]:
            e = normalize_email(get_cell(r, col_for_output.get("email")))
            p = normalize_phone(get_cell(r, col_for_output.get("number")),
                                keep_original=False)
            key = dedupe_key(e, p)
            if key and key in seen:
                duplicates += 1
                continue
            if key:
                seen.add(key)

            out_row = [get_cell(r, col_for_output[h]) for h in OUTPUT_HEADERS]
            if key and key in sent_map:
                for h, val in sent_map[key].items():
                    if val:
                        out_row[OUTPUT_HEADERS.index(h)] = val
            all_rows_out.append(out_row)
            added += 1

        print(f"✅ Added {added} row(s) from '{tab}'")

    if not all_rows_out:
        raise SystemExit(
            "❌ All source tabs were empty — refusing to overwrite "
            f"'{existing_title}' with nothing. (Tab left untouched.)"
        )

    overwrite_tab(svc, existing_title, OUTPUT_HEADERS, all_rows_out)
    print(f"\n✅ Done. Wrote {len(all_rows_out)} combined row(s) into "
          f"'{existing_title}' ({duplicates} duplicate(s) skipped).")


if __name__ == "__main__":
    combine()
