# Scripts to generate reports from a RockMaker databases

## Installing dependencies

```bash
pip install --user XlsxWriter
pip install --user python-tds
```

## Configuration

Copy the file db.example.cfg to db.cfg and customise it to use your own database credentials.

## Running

```bash
python plates_to_xlsx.py month 2018 01
```
