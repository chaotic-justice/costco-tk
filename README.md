# run costco analyzer on your pc

## HOW TO

download as a zip from github. unzip it, open a terminal, check out to this folder location, run this command. make sure you have `uv` installed first.

```bash
uv run pyinstaller --onefile --windowed pdf_counter.py
```

## making a change to the store numbers

- update store_numbers.csv
- goto utils/csv_string.py, delete everything inside the triple quotes, copy paste the entire csv content there
- rerun that `uv` command above
