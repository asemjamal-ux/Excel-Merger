# Excel Merger

Merge multiple Excel workbooks (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`) into one file, with a
`Source File` column showing where each row came from. Options: pick a sheet by name,
skip header rows, drop columns (by letter, index or name), remove duplicates, drop empty rows.

## Web version — `public/`

A fully static single-page app. All parsing and merging runs **in the browser** via
[SheetJS](https://sheetjs.com); your files are never uploaded anywhere.


## Python versions

- `app.py` — Flask web server variant (`pip install -r requirements.txt && python app.py`).
- `excel_merger.py` — desktop GUI (customtkinter).
