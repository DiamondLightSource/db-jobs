# Database reporting scripts

## Features

- Reads config file
- Reads command-line arguments
- Queries the database to get a result set
- Writes the query result set to a spreadsheet file (.xlsx or .csv)
- If recipient email addresses are defined in the config file, then emails the spreadsheet as an attachment to these

## Installing dependencies

```bash
pip install --user XlsxWriter
pip install --user python-tds
pip install --user mysql-connector
```

## Configuration

Copy the file config.example.cfg to config.cfg and customise it to use your own database credentials and email settings.

## Example usage

```bash
# report on plates and imagings from a RockMaker database
python rmaker_plates_to_xlsx.py month 2018 10
# report on plates and imagings from an ISPyB database
python ispyb_plates_to_xlsx.py month 2018 09
```

## Developing new reports

You will need:
- An SQL template string for your database query
- A list with the column headers you want to use in the report
- If your database system is not yet supported, extend the DBReport class and implement the `create_report` method. (See `mariadbreport.py` or `mssqlreport.py` for examples.)
- If you don't have the database credentials yet in the `config.cfg` file, add them to the file under a new section.

See `ispyb_plates_to_xlsx.py` or `rmaker_plates_to_xlsx.py` for examples of how to put it all together.
