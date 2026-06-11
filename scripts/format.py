from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border
def main():
    EXCEL_FILE = "../books/books_test.xlsx"

    COLORS = {
        "good":       "C6EFCE",  # green
        "no_results": "FCE4D6",  # orange
        "no_author":  "FFF2CC",  # yellow
        "bad_search": "F4CCCC",  # pink
        "error":      "EA9999",  # red
        "na":         "D9D9D9",  # grey
        "other":      "FFE599",  # light orange
    }

    NO_FILL   = PatternFill(fill_type=None)
    NO_BORDER = Border()

    def get_color(value):
        if value is None:
            return None
        v = str(value).strip()
        if v.isdigit() or v == "eb" or v.endswith("eb"):
            return COLORS["good"]
        elif v == "NO RESULTS":
            return COLORS["no_results"]
        elif v == "NO AUTHOR":
            return COLORS["no_author"]
        elif v == "BAD SEARCH":
            return COLORS["bad_search"]
        elif v == "ERROR":
            return COLORS["error"]
        elif v == "N/A":
            return COLORS["na"]
        else:
            return COLORS["other"]


    wb = load_workbook(EXCEL_FILE)

    # Remove all table definitions so openpyxl's fills aren't overridden by table styling
    for ws in wb.worksheets:
        for table in list(ws.tables.values()):
            del ws.tables[table.displayName]

    ws = wb.active

    # Find OL column
    ol_col = None
    for cell in ws[1]:
        if cell.value == "OL":
            ol_col = cell.column
            break

    if ol_col is None:
        print("Could not find OL column.")
        exit()

    # Wipe ALL cell fills on every row, then apply our own to OL
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.fill    = NO_FILL
            cell.border  = NO_BORDER

        ol_cell   = row[ol_col - 1]
        hex_color = get_color(ol_cell.value)
        if hex_color:
            ol_cell.fill = PatternFill(patternType="solid", fgColor=hex_color)

    wb.save(EXCEL_FILE)
    print("Done.")

if __name__ == "__main__":
    main()