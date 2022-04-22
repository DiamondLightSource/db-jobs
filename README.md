# Database reporting scripts

## Features

- Reads config file
- Reads command-line arguments
- Queries the database to get a result set
- Writes the query result set to a spreadsheet file (.xlsx or .csv)
- If recipient email addresses are defined in the config file, then emails the spreadsheet as an attachment (or just a path name, depending on your configuration).

## Installing dependencies

```bash
pip install --user XlsxWriter
pip install --user python-tds
pip install --user mysql-connector
pip install --user psycopg2
```

## Configuration

Copy the files `config.example.cfg` to `config.cfg` and `datasources.example.cfg` to `datasources.cfg`, then customise them to use your own database credentials and email settings.

## Example usage

```bash
# report on plates and imagings from a RockMaker database
python runreport.py RockMakerPlateReport month 2018 10
# report on plates and imagings from an ISPyB database
python runreport.py ISPyBPlatesReport month 2018 09
```

## Developing new reports

You will need to add this to the 'reports.cfg' file:
- An SQL template string for your database query
- A list with the column headers you want to use in the report

You also need to add database credentials to the `datasources.cfg` file and email settings to the `config.cfg` file.

If your database system is not yet supported, amend the `create_report` in the DBReport class.

See the .cfg files for examples.
