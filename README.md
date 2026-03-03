# vCard Converter

A flexible, modular tool for converting contacts between CSV, plain text, and vCard (`.vcf`) formats — with a rich CLI and a Streamlit web UI.

## Features

- **CSV → vCard**: Standard CSVs, Google Contacts Takeout CSVs, columns in any order
- **VCF → VCF**: Merge and reformat multiple vCard files
- **Free text → vCard**: Paste WhatsApp group exports, tab-separated rows, or loose prose
- **Rich CLI** with interactive prompts when flags are missing, progress bar, preview table
- **Streamlit web UI** with editable contacts table before download
- **Duplicate detection** (`--dedupe`)
- **No double-space phone bug** — fixed from the original `Convert2.py`
- vCard 3.0 (default) and 2.1 output
- JSON error reports for skipped rows

## Quick Start

```bash
# Install dependencies (uses uv)
uv sync

# Or with pip
pip install -r requirements.txt
pip install -e .
```

## CLI Usage

```bash
# Convert a standard CSV (prompts for anything missing)
uv run python -m vcard_convert.cli convert -i contacts.csv -o out.vcf

# Google Contacts Takeout CSV
uv run python -m vcard_convert.cli convert -i google_export.csv --from google-csv

# Different country code
uv run python -m vcard_convert.cli convert -i contacts.csv --country-code +1

# Free-form text file
uv run python -m vcard_convert.cli convert -i data.txt --from text

# Paste from stdin
uv run python -m vcard_convert.cli paste

# Merge multiple VCF files, removing duplicates
uv run python -m vcard_convert.cli merge a.vcf b.vcf -o merged.vcf --dedupe

# Show stats for any contact file
uv run python -m vcard_convert.cli info contacts.vcf

# Skip confirmation prompt
uv run python -m vcard_convert.cli convert -i contacts.csv -y
```

### Supported paste / text formats

```
Raj Kumar ~ +91 98765 43210          ← WhatsApp group export
Priya Sharma	8765432109             ← Tab-separated
Arun, 9123456789                     ← Comma-separated
Call Neha at 9000011112              ← Loose prose
```

## Web UI

```bash
uv run streamlit run vcard_convert/web/app.py
```

Three tabs:
1. **Upload File** — drag-drop CSV or VCF
2. **Paste Text** — paste messy data, review parsed contacts in an editable table, download VCF
3. **Google Contacts** — upload Google Takeout CSV directly

## Project Structure

```
vcard_convert/          ← main package
├── core/
│   ├── models.py       ← Contact, ParseResult dataclasses
│   ├── phone.py        ← PhoneNormalizer (fixes double-space bug)
│   ├── parsers/
│   │   ├── base.py     ← AbstractParser
│   │   ├── csv_parser.py   ← CSV + Google Contacts CSV
│   │   ├── vcf_parser.py   ← vCard 2.1/3.0 reader
│   │   └── text_parser.py  ← Free-form text parser
│   └── writers/
│       ├── base.py     ← AbstractWriter
│       └── vcf_writer.py   ← vCard 3.0/2.1 writer
├── cli/
│   ├── main.py         ← argparse CLI with rich output
│   └── __main__.py
└── web/
    └── app.py          ← Streamlit web UI
tests/                  ← pytest unit tests (46 tests)
legacy/                 ← original convert.py / Convert2.py (kept for reference)
Files/                  ← sample input files
output/                 ← default output directory
```

## Running Tests

```bash
uv run pytest tests/ -v
```

## CLI Reference

| Command | Description |
|---|---|
| `convert -i FILE` | Convert CSV / VCF / text to vCard |
| `merge FILE [FILE…]` | Merge multiple VCF files |
| `paste` | Read from stdin (paste workflow) |
| `info FILE` | Show stats without converting |

### Common flags

| Flag | Default | Description |
|---|---|---|
| `--from FORMAT` | auto-detect | `csv`, `google-csv`, `vcf`, `text` |
| `--country-code CC` | `+91` | Default country dialing code |
| `--vcf-version VER` | `3.0` | vCard version (`2.1` or `3.0`) |
| `--dedupe` | off | Remove contacts with duplicate phone numbers |
| `-y / --yes` | off | Skip confirmation prompts |
| `--output-dir DIR` | `output/` | Where to write output files and error reports |
