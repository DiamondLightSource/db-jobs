# Database reporting scripts

## Features

- Reads config file
- Reads command-line arguments
- Queries the database to get a result set
- Writes the query result set to a spreadsheet file (.xlsx)
- Emails the spreadsheet as an attachment to the recipient email addresses defined in the config file

## Installing dependencies

```bash
pip install --user XlsxWriter
pip install --user python-tds
pip install --user mysql-connector
```

## Configuration

Copy the file config.example.cfg to config.cfg and customise it to use your own database credentials and email settings.

## Examples

```bash
# report on plates and imagings from a RockMaker database
python rmaker_plates_to_xlsx.py month 2018 10
# report on plates and imagings from an ISPyB database
python ispyb_plates_to_xlsx.py month 2018 09
```
