# python-keyword-script

Small CLI tool to read an Excel file, filter rows that contain provided keywords, and create/append results to an output Excel workbook.

## Install

Create a virtualenv and install dependencies:

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
python kw_processor.py --input input.xlsx --output results.xlsx --keywords "error,fail" --columns Message,Details
```

Options:

- `--input` / `-i`: input Excel file
- `--output` / `-o`: output Excel file (created or appended)
- `--keywords` / `-k`: comma-separated keywords
- `--keywords-file`: text file with one keyword per line
- `--columns` / `-c`: optional, comma-separated columns to search
- `--sheet`: sheet name to write to (default: `results`)
- `--dry-run`: do not write output, only report matches

## Testing

Run tests with pytest:

```bash
pytest -q
```
