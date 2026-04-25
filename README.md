# odf-skills

Agent skills for working with [OpenDocument Format](https://en.wikipedia.org/wiki/OpenDocument) files — the open standard behind LibreOffice, Google Docs exports, and more.

Skills follow the [Agent Skills](https://agentskills.io) open standard and work with any compatible agent.

## Contents

- [Installation](#installation)
- [Skills](#skills)
  - [ods — OpenDocument Spreadsheet](#ods--opendocument-spreadsheet)
  - [odt — OpenDocument Text](#odt--opendocument-text)
- [Running tests](#running-tests)
- [Repository layout](#repository-layout)

## Installation

Requires [uv](https://docs.astral.sh/uv/):

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Install a skill by curling its two files into `~/.claude/skills/` (or `.claude/skills/` for a project-local install):

**ods:**
```bash
mkdir -p ~/.claude/skills/ods/scripts
curl -fsSL https://raw.githubusercontent.com/mnckapilan/odf-skills/main/ods/SKILL.md -o ~/.claude/skills/ods/SKILL.md
curl -fsSL https://raw.githubusercontent.com/mnckapilan/odf-skills/main/ods/scripts/ods.py -o ~/.claude/skills/ods/scripts/ods.py
```

**odt:**
```bash
mkdir -p ~/.claude/skills/odt/scripts
curl -fsSL https://raw.githubusercontent.com/mnckapilan/odf-skills/main/odt/SKILL.md -o ~/.claude/skills/odt/SKILL.md
curl -fsSL https://raw.githubusercontent.com/mnckapilan/odf-skills/main/odt/scripts/odt.py -o ~/.claude/skills/odt/scripts/odt.py
```

Then invoke with `/ods` or `/odt` in your agent. No further setup needed — the Python dependency ([odfpy](https://github.com/eea/odfpy)) is declared as a [PEP 723](https://peps.python.org/pep-0723/) inline dependency and installed automatically by `uv` on first use.

## Skills

### `ods` — OpenDocument Spreadsheet

Read and write `.ods` spreadsheet files — the format used by LibreOffice Calc, Google Sheets (on export), and other open-source office tools.

**Commands:**

| Command | What it does |
|---|---|
| `file-info` | Sheet names, row/col counts |
| `list-sheets` | Sheet names only |
| `read-sheet` | All rows as JSON; supports `--offset` and `--limit` for pagination |
| `get-cell` | Single cell value by A1-style reference |
| `set-cell` | Set a cell value (string, float, or bool) |
| `append-rows` | Append one or more rows |
| `create` | Create a new ODS file with named sheets |
| `add-sheet` | Add a sheet to an existing file |
| `rename-sheet` | Rename a sheet |
| `delete-sheet` | Delete a sheet (requires `--confirm`) |

**Caveats:**
- `set-cell` rebuilds the sheet element — cell-level formatting (fonts, colours, borders) and formulas are not preserved

---

### `odt` — OpenDocument Text

Read and write `.odt` document files — the format used by LibreOffice Writer and other open-source word processors.

Paragraphs and headings are addressed by 0-based index. Use `file-info` to see counts and `read-text` to browse content before editing.

**Commands:**

| Command | What it does |
|---|---|
| `file-info` | Title, paragraph/heading counts, word count |
| `read-text` | All blocks as JSON; supports `--offset` and `--limit` for pagination |
| `get-paragraph` | Single block by index |
| `set-paragraph` | Replace a block's text (optionally change its style) |
| `append-paragraph` | Append a block at the end |
| `insert-paragraph` | Insert a block before a given index |
| `delete-paragraph` | Delete a block (requires `--confirm`) |
| `list-headings` | All headings with level and index |
| `find-replace` | Find and replace text across all blocks; supports `--dry-run` |
| `create` | Create a new ODT file (optionally set a title) |

**Supported styles** for `--style`: `default`, `h1`–`h6`, `title`, `subtitle`, or any named paragraph style.

**Caveats:**
- `set-paragraph` and `find-replace` rebuild modified blocks — inline formatting (bold, italic, etc.) within changed blocks is not preserved
- Only direct paragraphs and headings are indexed; content nested inside tables or text frames is not included

---

## Running tests

```bash
# All skills
uv run --with pytest pytest ods/tests/ odt/tests/ -v

# One skill
uv run --with pytest pytest ods/tests/ -v
uv run --with pytest pytest odt/tests/ -v
```

## Repository layout

```
odf-skills/
├── ods/                  # ODS skill
│   ├── SKILL.md
│   ├── scripts/ods.py
│   ├── tests/
│   └── evals/
└── odt/                  # ODT skill
    ├── SKILL.md
    ├── scripts/odt.py
    ├── tests/
    └── evals/
```

Each skill is self-contained and can be installed independently.
