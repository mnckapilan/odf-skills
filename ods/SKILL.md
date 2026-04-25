---
name: ods
description: >
  Work with ODS (OpenDocument Spreadsheet) files. Use when the user wants to
  read, write, create, or modify .ods spreadsheet files — listing sheets,
  reading rows, getting or setting cells, appending data, or managing sheets.
license: MIT
compatibility: Requires uv and Python 3.11+
allowed-tools: Bash(uv:*)
---

# ODS skill

Read and write ODS spreadsheet files using `scripts/ods.py`.

**Prerequisite:** [`uv`](https://docs.astral.sh/uv/) must be installed.
Install it with: `curl -LsSf https://astral.sh/uv/install.sh | sh`

## Available scripts

- **`scripts/ods.py`** — All ODS operations via subcommands (see below).

Run `uv run scripts/ods.py --help` for the full command list, or
`uv run scripts/ods.py <command> --help` for per-command usage.

## Commands

| Command | What it does |
|---|---|
| `file-info <file>` | Sheet names, row/col counts — start here when exploring a file |
| `list-sheets <file>` | Sheet names only |
| `read-sheet <file> --sheet NAME` | All rows as JSON; supports `--offset N` and `--limit N` |
| `get-cell <file> --sheet NAME --cell A1` | Single cell value |
| `set-cell <file> --sheet NAME --cell A1 --value V [--type float\|string\|bool]` | Set a cell |
| `append-rows <file> --sheet NAME --rows JSON` | Append rows |
| `create <file> [--sheets S1 S2 …] [--overwrite]` | Create a new ODS file |
| `add-sheet <file> --sheet NAME` | Add a sheet |
| `rename-sheet <file> --sheet OLD --new-name NEW` | Rename a sheet |
| `delete-sheet <file> --sheet NAME --confirm` | Delete a sheet |

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Success |
| 1 | Invalid arguments |
| 2 | File not found |
| 3 | Sheet not found |
| 5 | ODF parse or write error |

## Workflow guidance

**Exploring a file** — always run `file-info` first to understand the
structure, then `read-sheet` for data. For large sheets use `--limit` and
`--offset` to page through rows rather than reading everything at once.

**Writing data** — `append-rows` is efficient for adding multiple rows.
`set-cell` is fine for single updates but rebuilds the sheet element, so
cell-level formatting (fonts, colours, borders) is not preserved. Formulas
are also replaced by their last computed values.

**Destructive operations** — `delete-sheet` requires `--confirm`. Always
check with the user before passing this flag.

**Idempotent creation** — `create` fails if the file exists unless
`--overwrite` is passed. Prefer checking first with `file-info` if
unsure whether the file already exists.

## Examples

```bash
# Explore a file
uv run scripts/ods.py file-info data.ods

# Read the first 20 rows of a sheet
uv run scripts/ods.py read-sheet data.ods --sheet "Sales" --limit 20

# Read rows 20–39 (second page)
uv run scripts/ods.py read-sheet data.ods --sheet "Sales" --offset 20 --limit 20

# Get a cell
uv run scripts/ods.py get-cell data.ods --sheet "Sales" --cell B3

# Set a numeric cell
uv run scripts/ods.py set-cell data.ods --sheet "Sales" --cell B3 --value 42 --type float

# Append rows
uv run scripts/ods.py append-rows data.ods --sheet "Sales" --rows '[["Alice", 1200], ["Bob", 950]]'

# Create a new file with two sheets
uv run scripts/ods.py create report.ods --sheets "Summary" "Data"

# Rename a sheet
uv run scripts/ods.py rename-sheet data.ods --sheet "Sheet1" --new-name "Sales"

# Delete a sheet (ask user before passing --confirm)
uv run scripts/ods.py delete-sheet data.ods --sheet "Scratch" --confirm
```
