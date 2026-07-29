# sheets_scrubber.py
# Removes existing clients from lead tabs so they never get cold outreach.
#
#   1) Collects phone numbers AND emails from SOURCE_OF_TRUTH_SHEETS
#   2) Iterates every other tab (except ignored + source + bookkeeping tabs)
#   3) Deletes any row whose phone or email matches a protected contact
#   4) Prints exactly who got removed, per tab
#
# Rows are deleted with a single atomic batchUpdate per tab (bottom-up),
# instead of the old clear-the-tab-then-rewrite approach that could wipe a
# whole tab if the run died mid-way. Formatting and formulas survive.

from googleapiclient.errors import HttpError

from common import (
    BOOKKEEPING_TABS, build_header_map, delete_rows, format_phone_us,
    get_cell, get_values, list_sheet_titles, normalize_email, normalize_phone,
    sheets_service,
)

# =========================
# ✅ EDIT THESE SETTINGS
# =========================
SOURCE_OF_TRUTH_SHEETS = [
    "clients",
    "potential clients",
]

IGNORE_SHEETS = [
    "Archive",
    "Do Not Touch",
]
# =========================


def collect_protected_contacts(svc):
    phones, emails = set(), set()
    for sheet in SOURCE_OF_TRUTH_SHEETS:
        rows = get_values(svc, sheet)
        if not rows or len(rows) < 2:
            print(f"⚠️ Source sheet '{sheet}' empty or missing rows. Skipping.")
            continue

        header_map = build_header_map(rows[0])
        phone_col = header_map.get("phone")
        email_col = header_map.get("email")
        if phone_col is None and email_col is None:
            print(f"⚠️ Source sheet '{sheet}' has no phone or email column. Skipping.")
            continue

        added_p = added_e = 0
        bad_phones = []
        for r in rows[1:]:
            raw_phone = get_cell(r, phone_col)
            p = normalize_phone(raw_phone, keep_original=False)
            if p:
                phones.add(p)
                added_p += 1
            elif str(raw_phone).strip():
                bad_phones.append(str(raw_phone).strip())
            e = normalize_email(get_cell(r, email_col))
            if e:
                emails.add(e)
                added_e += 1

        print(f"✅ Collected {added_p} phone(s), {added_e} email(s) from '{sheet}'")
        if bad_phones:
            # These clients can't be matched by phone — surface them so the
            # sheet can be fixed rather than silently leaving them unprotected.
            print(f"   ⚠️ {len(bad_phones)} phone value(s) in '{sheet}' didn't "
                  f"normalize to 10 digits and give NO protection: "
                  f"{', '.join(bad_phones[:10])}"
                  + (" ..." if len(bad_phones) > 10 else ""))

    return phones, emails


def scrub_sheet(svc, sheet_name, protected_phones, protected_emails):
    rows = get_values(svc, sheet_name)
    if not rows or len(rows) < 2:
        return 0, 0, []

    header_map = build_header_map(rows[0])
    phone_col = header_map.get("phone")
    email_col = header_map.get("email")
    if phone_col is None and email_col is None:
        print(f"↪ Skipping '{sheet_name}' (no phone or email column found).")
        return 0, len(rows) - 1, []

    first_idx = header_map.get("first_name")
    last_idx = header_map.get("last_name")

    rows_to_delete = []
    removed_people = []

    for row_number, r in enumerate(rows[1:], start=2):
        p = normalize_phone(get_cell(r, phone_col), keep_original=False)
        e = normalize_email(get_cell(r, email_col))

        matched = (p and p in protected_phones) or (e and e in protected_emails)
        if not matched:
            continue

        first = str(get_cell(r, first_idx)).strip() if first_idx is not None else ""
        last = str(get_cell(r, last_idx)).strip() if last_idx is not None else ""
        name = f"{first} {last}".strip() or "(no name)"
        removed_people.append((name, format_phone_us(p) or e))
        rows_to_delete.append(row_number)

    if rows_to_delete:
        delete_rows(svc, sheet_name, rows_to_delete)

    kept = (len(rows) - 1) - len(rows_to_delete)
    return len(rows_to_delete), kept, removed_people


def main():
    svc = sheets_service()

    source_lower = {s.strip().lower() for s in SOURCE_OF_TRUTH_SHEETS}
    ignore_lower = {s.strip().lower() for s in IGNORE_SHEETS}
    ignore_lower |= {s.strip().lower() for s in BOOKKEEPING_TABS}

    print("🔎 Building protected list from source sheets...")
    protected_phones, protected_emails = collect_protected_contacts(svc)
    print(f"\n✅ Protected: {len(protected_phones)} phone(s), "
          f"{len(protected_emails)} email(s)\n")

    if not protected_phones and not protected_emails:
        print("No protected contacts found. Nothing to scrub.")
        return

    removed_total = 0
    for sh in list_sheet_titles(svc):
        sh_lower = sh.strip().lower()
        if sh_lower in source_lower:
            print(f"🔒 Not touching source sheet '{sh}'")
            continue
        if sh_lower in ignore_lower:
            print(f"⏭️ Ignoring '{sh}'")
            continue

        try:
            removed, kept, removed_people = scrub_sheet(
                svc, sh, protected_phones, protected_emails
            )
        except HttpError as e:
            print(f"❌ Failed on '{sh}' ({e}); continuing with the next tab.")
            continue

        removed_total += removed
        if removed > 0:
            print(f"🧹 Scrubbed '{sh}':")
            for name, contact in removed_people:
                print(f"   ❌ {name} | {contact}")
            print(f"   → removed {removed}, kept {kept}")
        else:
            print(f"✅ No matches in '{sh}' (kept {kept})")

    print(f"\n✅ Done. Total removed across sheets: {removed_total}")


if __name__ == "__main__":
    main()
