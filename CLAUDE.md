# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A collection of Claude Code skills for working with ODF (OpenDocument Format) files. Each skill lives in its own subdirectory and can be installed by copying that directory to `~/.claude/skills/<name>/`.

Current skills: `ods/` (OpenDocument Spreadsheet), `odt/` (OpenDocument Text).

## Running tests

```bash
# All tests (both skills)
uv run --with pytest pytest ods/tests/ odt/tests/ -v

# One skill only
uv run --with pytest pytest ods/tests/ -v
uv run --with pytest pytest odt/tests/ -v

# Single test
uv run --with pytest pytest ods/tests/test_ods.py::TestReadSheet::test_limit -v
uv run --with pytest pytest odt/tests/test_odt.py::TestReadText::test_reads_all_blocks -v
```

Tests require `uv` in PATH. No other setup needed — odfpy is declared as an inline dependency in the script and installed automatically.

Each skill directory has `__init__.py` files (in `<format>/` and `<format>/tests/`) so that both test suites can run together without conftest naming conflicts. Test files use relative imports (`from .conftest import run`).

## Skill structure

Each skill directory follows this layout:

```
<format>/
├── SKILL.md          # Frontmatter + instructions loaded by Claude Code
├── scripts/
│   └── <format>.py  # Self-contained CLI script (PEP 723 inline deps, run via uv)
├── evals/
│   ├── evals.json   # Eval test cases (prompts, expected outputs, assertions)
│   └── files/       # Committed fixture files for evals
└── tests/
    ├── conftest.py  # Shared fixtures (create temp files via the script itself)
    └── test_<format>.py
```

## Scripts

Scripts are self-contained via [PEP 723](https://peps.python.org/pep-0723/) inline dependency declarations and run with `uv run`. Always invoke as `uv run scripts/<format>.py <command>`, never `python scripts/<format>.py`.

All commands output JSON to stdout and errors to stderr. Exit codes: 0 success, 1 invalid args, 2 file not found, 3 not found (sheet/block), 5 ODF error.

### odfpy attribute API

odfpy's `getAttribute`/`setAttribute` accept lowercase hyphen-stripped names, not namespace-prefixed strings:
- `"valuetype"` not `"office:value-type"`
- `"numbercolumnsrepeated"` not `"table:number-columns-repeated"`
- `"name"` not `"table:name"` (on Table elements)

Direct dict access (`element.attributes[ns_tuple]`) works but bypasses grammar validation — prefer the string API.

### odfpy element types

odfpy's `P`, `H`, `Table`, etc. are **factory functions**, not classes. `isinstance(elem, P)` does not work. Identify element types by qualified name instead:

```python
from odf.namespaces import TEXTNS
elem.qname == (TEXTNS, 'p')   # paragraph
elem.qname == (TEXTNS, 'h')   # heading
```

## Tests vs evals

**Tests** (`tests/`) are deterministic CI tests for the script CLI. Each test creates fresh files via `tmp_path` fixtures and checks exact JSON output and exit codes.

**Evals** (`evals/`) assess skill quality — does Claude use the script correctly given realistic user prompts? These require spawning Claude with/without the skill and an LLM judge to grade assertions. Results go in `<format>-workspace/` (gitignored). Not wired into CI.

## Adding a new skill

1. Create `<format>/` with the structure above
2. Add empty `__init__.py` to `<format>/` and `<format>/tests/`
3. Use `from .conftest import run` in the test file (relative import)
4. Add `<format>-workspace/` to `.gitignore`
5. Add a test job to `.github/workflows/test.yml`
