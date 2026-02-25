"""
readexcelmain.py
----------------
Reads 'supermarket_sales.xlsx' and prints its contents as a clean table.

Run AFTER create_sales_data.py has been run to generate the Excel file.

Requires: pip install openpyxl
"""

import openpyxl


def print_table(filename="supermarket_sales.xlsx"):
    """
    Open the Excel file, read every row, and display
    everything as a formatted text table in the terminal.
    """

    # ── Load the workbook and grab the active sheet ──────────────────────────
    wb = openpyxl.load_workbook(filename)
    ws = wb.active   # 'Supermarket Sales' sheet

    # ── Read all rows into a plain list of tuples ────────────────────────────
    all_rows = []
    for row in ws.iter_rows(values_only=True):   # values_only skips cell objects
        all_rows.append([str(cell) for cell in row])  # convert every value to string

    if not all_rows:
        print("The Excel file is empty.")
        return

    # First row is the header, the rest are data rows
    header = all_rows[0]
    data   = all_rows[1:]

    # ── Calculate the display width for each column ──────────────────────────
    # Start with the header widths, then expand if any data cell is wider
    col_widths = [len(h) for h in header]
    for row in data:
        for i, cell in enumerate(row):
            if len(cell) > col_widths[i]:
                col_widths[i] = len(cell)

    # ── Build reusable separator and row-formatting helpers ──────────────────
    # e.g.  +----------+----------+
    separator = "+-" + "-+-".join("-" * w for w in col_widths) + "-+"

    def format_row(cells):
        """Left-align each cell value within its column width, wrap in pipes."""
        padded = [cells[i].ljust(col_widths[i]) for i in range(len(cells))]
        return "| " + " | ".join(padded) + " |"

    # ── Print the table ──────────────────────────────────────────────────────
    title = "SUPERMARKET SALES DATA — WORLDWIDE"
    print()
    print(title.center(len(separator)))   # centred title above the table
    print(separator)
    print(format_row(header))             # header row
    print(separator)
    for row in data:
        print(format_row(row))            # one line per sale
    print(separator)

    # Summary line at the bottom
    print(f"\nTotal sales records: {len(data)}")


# ── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print_table()
