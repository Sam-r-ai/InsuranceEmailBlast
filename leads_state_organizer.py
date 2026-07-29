# leads_state_organizer.py
# Sorts one tab alphabetically by its State column, in place.
#
# Uses a single sortRange batchUpdate request — atomic on Google's side, so
# a crash or network error can no longer leave the tab cleared, and cell
# formatting/formulas survive (the old version rewrote the whole tab as raw
# text). Blank states end up at the bottom.

from common import (
    build_header_map, get_sheet_id, get_values, sheets_service, with_backoff,
    SPREADSHEET_ID,
)

# 👇 CHANGE THIS AND PRESS ▶ RUN
TARGET_SHEET_NAME = "Combined TTC 12/30"


def sort_sheet_by_state():
    svc = sheets_service()

    rows = get_values(svc, TARGET_SHEET_NAME, "A1:ZZ")
    if not rows or len(rows) < 2:
        print("Nothing to sort.")
        return

    header = rows[0]
    header_map = build_header_map(header)
    state_col = header_map.get("state")
    if state_col is None:
        raise SystemExit(
            f"❌ Could not find a State column.\n"
            f"Header row was: {header}"
        )

    sheet_id = get_sheet_id(svc, TARGET_SHEET_NAME)
    if sheet_id is None:
        raise SystemExit(f"❌ Tab not found: {TARGET_SHEET_NAME}")

    request = {
        "sortRange": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,  # keep the header in place
                "startColumnIndex": 0,
            },
            "sortSpecs": [
                {"dimensionIndex": state_col, "sortOrder": "ASCENDING"}
            ],
        }
    }
    with_backoff(
        lambda: svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": [request]}
        ).execute(),
        what=f"sort {TARGET_SHEET_NAME}",
    )

    print(
        f"✅ Sheet '{TARGET_SHEET_NAME}' sorted alphabetically by STATE "
        f"(column {state_col + 1})."
    )


if __name__ == "__main__":
    sort_sheet_by_state()
