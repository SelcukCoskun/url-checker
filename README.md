# Local URL Checker

A local web app that checks URLs from a file and shows HTTP status codes, redirects, final URLs, page titles, and short summaries.

- Supports XLSX, XLSM, XLS, ODS, CSV, TSV, and TXT files.
- Scans only the file you select (up to 100 URLs and 10 MB).
- Filter results and export them to XLSX.
- Runs on your computer at `127.0.0.1`; uploaded files are processed in memory.

## Requirements

Python 3.10 or later.

## Run on Windows

In Command Prompt, from the project folder:

```cmd
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python run.py
```

## Run on Linux

From the project folder:

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
python run.py
```

The app opens in your browser. Stop it with `Ctrl+C`.
