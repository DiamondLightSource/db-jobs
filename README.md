# Scripts to generate reports from a RockMaker databases

## Installing dependencies

```bash
pip install --user XlsxWriter
pip install --user python-tds
```

## Configuration

Copy the file config.example.cfg to config.cfg and customise it to use your own database credentials and email settings.

## Running

```bash
python plates_to_xlsx.py month 2018 01
```
