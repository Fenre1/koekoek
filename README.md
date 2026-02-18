# Koekoek Timeline Generator

Generate horizontal, vertical, or combined timeline HTML views from an Excel file.

## Requirements

- Python 3.8+
- `pandas`
- Excel reader engine (typically `openpyxl`)

Install:

```bash
pip install pandas openpyxl
```

## Quick Start (CLI)

Use `timeline_cli.py`:

```bash
python timeline_cli.py horizontal -i "20260115 Tijdlijn.xlsx" -o "timeline_horizontal.html"
python timeline_cli.py vertical   -i "20260115 Tijdlijn.xlsx" -o "timeline_vertical_filterable.html"
python timeline_cli.py combined   -i "20260115 Tijdlijn.xlsx" -o "timeline_combined.html"
```

## Notebook / Cell-by-Cell Workflow (`#%%`)

If you prefer debugging step-by-step in VS Code/Jupyter style, use:

- `timeline_notebook.py`

This file has:

1. a config cell (`mode`, `excel_path`, `output_path`)
2. a function cell (`generate_timeline(...)`)
3. a run cell that executes the selected mode

Run cells individually to troubleshoot input/output and rendering changes.

## Project Structure

- `timeline_core.py`: shared parsing, normalization, entity colors, source handling
- `timeline_horizontal.py`: horizontal renderer
- `timeline_vertical_filterable.py`: vertical filterable renderer
- `timeline_combined_viewer.py`: combined viewer generator (horizontal + vertical)
- `timeline_cli.py`: CLI entrypoint
- `timeline_notebook.py`: notebook-style runner with `#%%` cells

## Expected Excel Columns

Required:

- `Datum`
- `Starttijd`
- `Eindtijd`
- `Zekerheid (ja/nee)`
- `Entiteit(en) (splits op met |)`
- `Gebeurtenis`
- `Geverifieerd`

Optional:

- `Bron`

## Notes

- Multiple entities should be separated by `|`.
- For links in `Bron`, values starting with `http://`, `https://`, or `www.` are rendered as clickable links.
- Unknown/missing values are handled with safe defaults where possible.
