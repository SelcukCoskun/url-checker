# Local URL Checker

A local web app that checks URLs from a file and displays HTTP status codes, redirects, final URLs, page titles, and short summaries. Supports XLSX, XLSM, XLS, ODS, CSV, TSV, and TXT. Only the selected file is scanned (up to 100 URLs and 10 MB); results can be filtered and exported to XLSX.

## Requirements

Git and Python 3.10 or later.

## Windows

Open Command Prompt and run:

```cmd
git clone https://github.com/SelcukCoskun/url-checker.git
cd url-checker\aktif-pasif
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python run.py
```

## Linux

Open a terminal and run:

```bash
git clone https://github.com/SelcukCoskun/url-checker.git
cd url-checker/aktif-pasif
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
python run.py
```

The app opens in your browser and listens only on `127.0.0.1`. Stop it with `Ctrl+C`
