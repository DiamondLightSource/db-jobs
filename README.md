# Generate RockMaker plate reports

## Features

- Reads config file
- Reads command-line arguments
- Queries the database to get the plates registered per month (or year), the number of times each has been imaged, plus more
- Writes the query result to a spreadsheet (.xlsx)
- Emails the spreadsheet as an attachment to the recipient email addresses defined in the config file

## Installing dependencies

```bash
pip install --user XlsxWriter
pip install --user python-tds
```

## Configuration

Copy the file config.example.cfg to config.cfg and customise it to use your own database credentials and email settings.

## Running

```bash
python rmaker_plates_to_xlsx.py month 2018 01
```
