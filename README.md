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
pip install --user psycopg2
```

## Configuration

Copy the file `config.example.cfg` to `config.cfg` and customise it to use your own database credentials and email settings.

## Example usage

```bash
# report on plates and imagings from a RockMaker database
python runreport.py RockMakerPlateReport month 2018 10
# report on plates and imagings from an ISPyB database
python runreport.py ISPyBPlatesReport month 2018 09
```

## Developing new reports

You will need to add this to your 'config.cfg' file:
- An SQL template string for your database query
- A list with the column headers you want to use in the report
- The database credentials

If your database system is not yet supported, amend the `create_report` in the DBReport class.

See the report sections in `config.example.cfg` for examples of how to put it all together.
