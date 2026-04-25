# /// script
# requires-python = ">=3.11"
# dependencies = ["odfpy>=1.4.1"]
# ///

"""
Read and write ODS (OpenDocument Spreadsheet) files.

Exit codes:
  0  success
  1  invalid arguments
  2  file not found
  3  sheet not found
  5  ODF parse / write error

odfpy API note: getAttribute/setAttribute accept lowercase, hyphen-stripped
attribute names, e.g. "valuetype" not "office:value-type".
"""

import argparse
import json
import re
import sys
from pathlib import Path


# ── output helpers ────────────────────────────────────────────────────────────

def die(msg: str, code: int = 1) -> None:
    print(f"Error: {msg}", file=sys.stderr)
    sys.exit(code)


def emit(data: object) -> None:
    print(json.dumps(data, ensure_ascii=False))
    sys.exit(0)


# ── document helpers ──────────────────────────────────────────────────────────

def load_doc(path: str):
    from odf.opendocument import load  # type: ignore[import]
    p = Path(path)
    if not p.exists():
        die(f"File not found: {path}", code=2)
    try:
        return load(str(p))
    except Exception as exc:
        die(f"Failed to open ODS file: {exc}", code=5)


def save_doc(doc, path: str) -> None:
    try:
        doc.save(path)
    except Exception as exc:
        die(f"Failed to save ODS file: {exc}", code=5)


def get_sheets(doc) -> list:
    from odf.table import Table  # type: ignore[import]
    return doc.spreadsheet.getElementsByType(Table)


def find_sheet(doc, name: str):
    for sheet in get_sheets(doc):
        if sheet.getAttribute("name") == name:
            return sheet
    return None


def require_sheet(doc, name: str):
    sheet = find_sheet(doc, name)
    if sheet is None:
        names = [s.getAttribute("name") for s in get_sheets(doc)]
        die(f"Sheet {name!r} not found. Available sheets: {names}", code=3)
    return sheet


# ── cell helpers ──────────────────────────────────────────────────────────────

def parse_cell_ref(ref: str) -> tuple[int, int]:
    """A1-style → (col, row), both 0-indexed."""
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        die(f"Invalid cell reference {ref!r}. Use A1-style notation, e.g. B3.")
    col_str, row_str = m.groups()
    col = 0
    for ch in col_str.upper():
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col - 1, int(row_str) - 1


def get_cell_value(cell):
    from odf.text import P  # type: ignore[import]
    vtype = cell.getAttribute("valuetype")
    if vtype is None:
        return None
    if vtype == "float":
        v = float(cell.getAttribute("value"))
        return int(v) if v == int(v) else v
    if vtype in ("percentage", "currency"):
        return float(cell.getAttribute("value"))
    if vtype == "date":
        return cell.getAttribute("datevalue")
    if vtype == "time":
        return cell.getAttribute("timevalue")
    if vtype == "boolean":
        return cell.getAttribute("booleanvalue") == "true"
    if vtype == "string":
        parts = [str(p) for p in cell.getElementsByType(P)]
        return "\n".join(parts)
    return None


def make_cell(value, vtype: str | None = None):
    from odf.table import TableCell  # type: ignore[import]
    from odf.text import P  # type: ignore[import]
    cell = TableCell()
    if value is None:
        return cell
    if vtype == "float" or (
        vtype is None and isinstance(value, (int, float)) and not isinstance(value, bool)
    ):
        cell.setAttribute("valuetype", "float")
        cell.setAttribute("value", str(value))
        cell.addElement(P(text=str(value)))
    elif vtype == "bool" or isinstance(value, bool):
        cell.setAttribute("valuetype", "boolean")
        cell.setAttribute("booleanvalue", "true" if value else "false")
        cell.addElement(P(text="TRUE" if value else "FALSE"))
    else:
        cell.setAttribute("valuetype", "string")
        cell.addElement(P(text=str(value)))
    return cell


# ── sheet reading ─────────────────────────────────────────────────────────────

def expand_row(tr) -> list:
    from odf.table import TableCell  # type: ignore[import]
    row: list = []
    for cell in tr.getElementsByType(TableCell):
        # cap repeat to avoid pathological files with millions of empty cols
        n = min(int(cell.getAttribute("numbercolumnsrepeated") or 1), 1024)
        v = get_cell_value(cell)
        row.extend([v] * n)
    while row and row[-1] is None:
        row.pop()
    return row


def read_rows(sheet) -> list[list]:
    from odf.table import TableRow  # type: ignore[import]
    rows: list[list] = []
    for tr in sheet.getElementsByType(TableRow):
        n = int(tr.getAttribute("numberrowsrepeated") or 1)
        row = expand_row(tr)
        # skip large blocks of empty padding rows (common at sheet bottom)
        if n > 1 and not any(v is not None for v in row):
            continue
        for _ in range(n):
            rows.append(row[:])
    while rows and not any(v is not None for v in rows[-1]):
        rows.pop()
    return rows


def build_table(name: str, data: list[list]):
    from odf.table import Table, TableRow  # type: ignore[import]
    table = Table(name=name)
    for row_data in data:
        tr = TableRow()
        for val in row_data:
            tr.addElement(make_cell(val))
        table.addElement(tr)
    return table


def replace_sheet(doc, old_sheet, new_table) -> None:
    parent = old_sheet.parentNode
    parent.insertBefore(new_table, old_sheet)
    parent.removeChild(old_sheet)


# ── commands ──────────────────────────────────────────────────────────────────

def cmd_list_sheets(args) -> None:
    doc = load_doc(args.file)
    names = [s.getAttribute("name") for s in get_sheets(doc)]
    emit({"sheets": names, "count": len(names)})


def cmd_read_sheet(args) -> None:
    doc = load_doc(args.file)
    sheet = require_sheet(doc, args.sheet)
    rows = read_rows(sheet)
    total = len(rows)
    offset = args.offset
    page = rows[offset : offset + args.limit] if args.limit else rows[offset:]
    emit({
        "sheet": args.sheet,
        "rows": page,
        "total_rows": total,
        "returned": len(page),
        "offset": offset,
    })


def cmd_get_cell(args) -> None:
    doc = load_doc(args.file)
    sheet = require_sheet(doc, args.sheet)
    col, row_idx = parse_cell_ref(args.cell)
    rows = read_rows(sheet)
    if row_idx >= len(rows):
        emit({"sheet": args.sheet, "cell": args.cell, "value": None})
    row = rows[row_idx]
    emit({"sheet": args.sheet, "cell": args.cell, "value": row[col] if col < len(row) else None})


def cmd_set_cell(args) -> None:
    doc = load_doc(args.file)
    sheet = require_sheet(doc, args.sheet)
    col, row_idx = parse_cell_ref(args.cell)

    rows = read_rows(sheet)
    while len(rows) <= row_idx:
        rows.append([])
    row = rows[row_idx]
    while len(row) <= col:
        row.append(None)

    value: object = args.value
    if args.type == "float":
        try:
            value = float(args.value)
        except ValueError:
            die(f"Cannot convert {args.value!r} to float.")
    elif args.type == "bool":
        value = args.value.lower() in ("true", "1", "yes")

    row[col] = value
    rows[row_idx] = row

    # Rebuild the sheet from the modified data matrix.
    # Note: this replaces the sheet element — cell formatting is not preserved.
    replace_sheet(doc, sheet, build_table(args.sheet, rows))
    save_doc(doc, args.file)
    emit({"success": True, "sheet": args.sheet, "cell": args.cell, "value": value})


def cmd_append_rows(args) -> None:
    doc = load_doc(args.file)
    sheet = require_sheet(doc, args.sheet)

    try:
        new_rows = json.loads(args.rows)
    except json.JSONDecodeError as exc:
        die(f"Invalid JSON for --rows: {exc}")

    if not isinstance(new_rows, list):
        die("--rows must be a JSON array of arrays.")

    from odf.table import TableRow  # type: ignore[import]
    for i, row_data in enumerate(new_rows):
        if not isinstance(row_data, list):
            die(f"Row {i} must be a JSON array.")
        tr = TableRow()
        for val in row_data:
            tr.addElement(make_cell(val))
        sheet.addElement(tr)

    save_doc(doc, args.file)
    emit({"success": True, "sheet": args.sheet, "rows_added": len(new_rows)})


def cmd_create(args) -> None:
    from odf.opendocument import OpenDocumentSpreadsheet  # type: ignore[import]
    from odf.table import Table  # type: ignore[import]

    p = Path(args.file)
    if p.exists() and not args.overwrite:
        die(f"File already exists: {args.file}. Use --overwrite to replace it.")

    doc = OpenDocumentSpreadsheet()
    sheet_names = args.sheets or ["Sheet1"]
    for name in sheet_names:
        doc.spreadsheet.addElement(Table(name=name))
    save_doc(doc, str(p))
    emit({"success": True, "path": str(p.resolve()), "sheets": sheet_names})


def cmd_add_sheet(args) -> None:
    from odf.table import Table  # type: ignore[import]
    doc = load_doc(args.file)
    if find_sheet(doc, args.sheet):
        die(f"Sheet {args.sheet!r} already exists.", code=3)
    doc.spreadsheet.addElement(Table(name=args.sheet))
    save_doc(doc, args.file)
    emit({"success": True, "sheet": args.sheet})


def cmd_rename_sheet(args) -> None:
    doc = load_doc(args.file)
    sheet = require_sheet(doc, args.sheet)
    if find_sheet(doc, args.new_name):
        die(f"Sheet {args.new_name!r} already exists.", code=3)
    sheet.setAttribute("name", args.new_name)
    save_doc(doc, args.file)
    emit({"success": True, "old_name": args.sheet, "new_name": args.new_name})


def cmd_delete_sheet(args) -> None:
    if not args.confirm:
        die("--confirm is required to delete a sheet. This cannot be undone.")
    doc = load_doc(args.file)
    sheet = require_sheet(doc, args.sheet)
    all_sheets = list(get_sheets(doc))
    if len(all_sheets) == 1:
        die("Cannot delete the only sheet in a file.")
    doc.spreadsheet.removeChild(sheet)
    save_doc(doc, args.file)
    emit({"success": True, "deleted": args.sheet})


def cmd_file_info(args) -> None:
    doc = load_doc(args.file)
    sheets_info = []
    for sheet in get_sheets(doc):
        rows = read_rows(sheet)
        max_cols = max((len(r) for r in rows), default=0)
        sheets_info.append({
            "name": sheet.getAttribute("name"),
            "rows": len(rows),
            "cols": max_cols,
        })
    emit({
        "path": str(Path(args.file).resolve()),
        "sheets": sheets_info,
        "sheet_count": len(sheets_info),
    })


# ── argument parser ───────────────────────────────────────────────────────────

EXAMPLES = """
examples:
  uv run scripts/ods.py file-info data.ods
  uv run scripts/ods.py list-sheets data.ods
  uv run scripts/ods.py read-sheet data.ods --sheet Sales --limit 50
  uv run scripts/ods.py get-cell data.ods --sheet Sales --cell B3
  uv run scripts/ods.py set-cell data.ods --sheet Sales --cell B3 --value 42 --type float
  uv run scripts/ods.py append-rows data.ods --sheet Sales --rows '[["Alice",1200],["Bob",950]]'
  uv run scripts/ods.py create report.ods --sheets Summary Data
  uv run scripts/ods.py add-sheet data.ods --sheet Notes
  uv run scripts/ods.py rename-sheet data.ods --sheet Notes --new-name Archive
  uv run scripts/ods.py delete-sheet data.ods --sheet Archive --confirm
"""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ods.py",
        description="Read and write ODS (OpenDocument Spreadsheet) files.",
        epilog=EXAMPLES,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version="ods.py 1.0")
    sub = parser.add_subparsers(dest="command", metavar="COMMAND")
    sub.required = True

    # list-sheets
    p = sub.add_parser(
        "list-sheets",
        help="List all sheet names.",
        description="List all sheet names in an ODS file.",
        epilog="Output: {sheets: [...], count: N}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.set_defaults(func=cmd_list_sheets)

    # read-sheet
    p = sub.add_parser(
        "read-sheet",
        help="Read rows from a sheet.",
        description="Read rows from a named sheet. Use --limit and --offset to paginate large sheets.",
        epilog="Output: {sheet, rows: [[...], ...], total_rows, returned, offset}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Sheet name.")
    p.add_argument("--offset", type=int, default=0, metavar="N",
                   help="Skip the first N rows (default: 0).")
    p.add_argument("--limit", type=int, metavar="N",
                   help="Return at most N rows. Omit to return all.")
    p.set_defaults(func=cmd_read_sheet)

    # get-cell
    p = sub.add_parser(
        "get-cell",
        help="Get a single cell value.",
        description="Get the value of a cell using A1-style notation.",
        epilog="Output: {sheet, cell, value}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Sheet name.")
    p.add_argument("--cell", required=True, metavar="REF",
                   help="Cell reference in A1 notation, e.g. B3.")
    p.set_defaults(func=cmd_get_cell)

    # set-cell
    p = sub.add_parser(
        "set-cell",
        help="Set a single cell value.",
        description=(
            "Set the value of a cell. "
            "Warning: rebuilds the sheet element — cell formatting is not preserved."
        ),
        epilog="Output: {success, sheet, cell, value}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Sheet name.")
    p.add_argument("--cell", required=True, metavar="REF",
                   help="Cell reference in A1 notation, e.g. B3.")
    p.add_argument("--value", required=True, help="New value for the cell.")
    p.add_argument("--type", choices=["string", "float", "bool"], default="string",
                   help="Value type (default: string). Use float for numbers.")
    p.set_defaults(func=cmd_set_cell)

    # append-rows
    p = sub.add_parser(
        "append-rows",
        help="Append rows to a sheet.",
        description="Append one or more rows to the end of a sheet.",
        epilog='Output: {success, sheet, rows_added}\nExample: --rows \'[["Alice", 30], ["Bob", 25]]\'',
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Sheet name.")
    p.add_argument("--rows", required=True, metavar="JSON",
                   help='JSON array of arrays, e.g. [["Alice", 30], ["Bob", 25]].')
    p.set_defaults(func=cmd_append_rows)

    # create
    p = sub.add_parser(
        "create",
        help="Create a new ODS file.",
        description="Create a new empty ODS file with one or more sheets.",
        epilog="Output: {success, path, sheets}",
    )
    p.add_argument("file", help="Path for the new .ods file.")
    p.add_argument("--sheets", nargs="+", metavar="NAME",
                   help="Sheet names to create (default: Sheet1).")
    p.add_argument("--overwrite", action="store_true",
                   help="Overwrite if the file already exists.")
    p.set_defaults(func=cmd_create)

    # add-sheet
    p = sub.add_parser(
        "add-sheet",
        help="Add a new sheet to an existing file.",
        description="Add a new empty sheet to an ODS file.",
        epilog="Output: {success, sheet}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Name for the new sheet.")
    p.set_defaults(func=cmd_add_sheet)

    # rename-sheet
    p = sub.add_parser(
        "rename-sheet",
        help="Rename a sheet.",
        description="Rename an existing sheet.",
        epilog="Output: {success, old_name, new_name}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Current sheet name.")
    p.add_argument("--new-name", required=True, dest="new_name", help="New sheet name.")
    p.set_defaults(func=cmd_rename_sheet)

    # delete-sheet
    p = sub.add_parser(
        "delete-sheet",
        help="Delete a sheet (requires --confirm).",
        description=(
            "Permanently delete a sheet from an ODS file. "
            "Cannot delete the last remaining sheet. Requires --confirm."
        ),
        epilog="Output: {success, deleted}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.add_argument("--sheet", required=True, help="Name of the sheet to delete.")
    p.add_argument("--confirm", action="store_true",
                   help="Required. Confirms you intend to permanently delete this sheet.")
    p.set_defaults(func=cmd_delete_sheet)

    # file-info
    p = sub.add_parser(
        "file-info",
        help="Show file metadata.",
        description="Show metadata: sheet names, row counts, and column counts.",
        epilog="Output: {path, sheet_count, sheets: [{name, rows, cols}, ...]}",
    )
    p.add_argument("file", help="Path to the .ods file.")
    p.set_defaults(func=cmd_file_info)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
