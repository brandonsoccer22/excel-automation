"""Run the native Power Query acceptance test in desktop Excel on Windows."""
import argparse
import sys
import tempfile
from pathlib import Path


def expect(actual, expected, label):
    if actual != expected:
        raise RuntimeError(f"{label}: expected {expected!r}, got {actual!r}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("parent", type=Path, help="Existing absolute output directory")
    args = parser.parse_args()
    if sys.platform != "win32":
        parser.exit(2, "Desktop Microsoft Excel COM requires Windows.\n")
    if not args.parent.is_absolute() or not args.parent.is_dir():
        parser.error("parent must be an existing absolute directory")

    import xlwings as xw
    from power_query import add_query, get_query, load_worksheet, refresh_worksheet, set_formula

    directory = Path(tempfile.mkdtemp(prefix="excel-automation-", dir=args.parent))
    original = directory / "original.xlsx"
    revised = directory / "revised.xlsx"
    name = "SmokeQuery"
    first = '#table(type table [Amount = number], {{10}, {20}})'
    second = '#table(type table [Amount = number], {{11}, {22}, {33}})'

    app = xw.App(visible=False, add_book=False)
    book = None
    try:
        print(f"Excel {app.api.Version} build {app.api.Build}; artifacts: {directory}", flush=True)
        book = app.books.add()
        sheet = book.sheets[0]
        add_query(book.api, name, first)
        table = load_worksheet(book.api, name, sheet.api, "A1")
        table_name = table.Name
        expect(table.DataBodyRange.Value2, ((10.0,), (20.0,)), "Initial data")
        refresh_worksheet(table.QueryTable)
        expect(table.DataBodyRange.Value2, ((10.0,), (20.0,)), "Explicit refresh")
        book.save(str(original))
        set_formula(book.api, name, second)
        refresh_worksheet(table.QueryTable)
        expect(table.DataBodyRange.Value2, ((11.0,), (22.0,), (33.0,)), "Modified M results")
        sheet.range("D1").formula = f"=SUM({table_name}[Amount])"
        app.calculate()
        expect(sheet.range("D1").value, 66, "Calculated total")
        book.save(str(revised))
        book.close()
        book = None

        book = app.books.open(str(original), update_links=False, read_only=True)
        expect(Path(book.api.FullName).resolve(), original.resolve(), "Original path")
        expect(get_query(book.api, name).Formula.strip(), first, "Original persisted M")
        book.close()
        book = None

        book = app.books.open(str(revised), update_links=False, read_only=False)
        expect(Path(book.api.FullName).resolve(), revised.resolve(), "Revised path")
        expect(get_query(book.api, name).Formula.strip(), second, "Revised persisted M")
        table = book.sheets[0].api.ListObjects.Item(table_name)
        if "Microsoft.Mashup.OleDb.1" not in str(table.QueryTable.WorkbookConnection.OLEDBConnection.Connection):
            raise RuntimeError("Reopened table is missing its native Mashup connection.")
        expect(table.DataBodyRange.Value2, ((11.0,), (22.0,), (33.0,)), "Reopened results")
        expect(book.sheets[0].range("D1").value, 66, "Reopened total")
        # Prove persisted M can still be edited and executed on the reopened disposable copy.
        set_formula(book.api, name, first)
        refresh_worksheet(table.QueryTable)
        expect(table.DataBodyRange.Value2, ((10.0,), (20.0,)), "Reopened query re-edit")
    finally:
        try:
            if book is not None:
                book.close()  # Only the script-owned disposable workbook.
        finally:
            app.quit()

    print(f"PASS: native editable query and refreshed results verified in {revised}")


if __name__ == "__main__":
    main()
