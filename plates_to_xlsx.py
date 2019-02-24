#
# Copyright 2019 Karl Levik
#

# Our imports:
import xlsxwriter
import pytds
import logging
from logging.handlers import RotatingFileHandler
import datetime
import sys

# Trick to make it work with both Python 2 and 3:
try:
  import configparser
except ImportError:
  import ConfigParser as configparser

# Configure logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('* %(asctime)s [id=%(thread)d] <%(levelname)s> %(message)s')
hdlr = RotatingFileHandler(filename='plates_to_xslx.log', maxBytes=1000000, backupCount=10)
hdlr.setFormatter(formatter)
logging.getLogger().addHandler(hdlr)

# Get input parameters 
start_year = sys.argv[1]  # e.g. 2018
start_month = sys.argv[2] # e.g. 02 
interval = sys.argv[3]    # accepted values: year or month
start_date = '%s/%s/01' % (start_year, start_month)

# Query to retrieve all plates registered and the number of times each has been imaged, within the reporting time frame:
sql = """SELECT pl.Barcode as "barcode", 
    tn4.Name as "project", 
    pl.DateDispensed as "date dispensed", 
    count(it.DateImaged) as "imagings", 
    u.Name as "user name", 
    g.Name as "group name", 
    c.Name as "plate type"
FROM Plate pl 
    INNER JOIN Experiment e ON pl.ExperimentID = e.ID
    INNER JOIN Containers c ON e.ContainerID = c.ID 
    INNER JOIN Users u ON e.userID = u.ID
    INNER JOIN GroupUser gu ON u.ID = gu.UserID
    INNER JOIN Groups g ON g.ID = gu.GroupID
    INNER JOIN TreeNode tn1 ON pl.TreeNodeID = tn1.ID 
    INNER JOIN TreeNode tn2 ON tn1.ParentID = tn2.ID
    INNER JOIN TreeNode tn3 ON tn2.ParentID = tn3.ID
    INNER JOIN TreeNode tn4 ON tn3.ParentID = tn4.ID
    LEFT OUTER JOIN ExperimentPlate ep ON ep.PlateID = pl.ID
    LEFT OUTER JOIN ImagingTask it ON it.ExperimentPlateID = ep.ID
WHERE pl.DateDispensed >= convert(date, '%s', 111) AND pl.DateDispensed <= dateadd(%s, 1, convert(date, '%s', 111)) 
    AND ((it.DateImaged >= convert(date, '%s', 111) AND it.DateImaged <= dateadd(%s, 1, convert(date, '%s', 111))) OR it.DateImaged is NULL)
	AND g.Name <> 'AllRockMakerUsers'
GROUP BY pl.Barcode, 
    tn4.Name, 
    pl.DateDispensed,  
    u.Name, 
    g.Name, 
    c.Name
ORDER BY pl.DateDispensed ASC
""" % (start_date, interval, start_date, start_date, interval, start_date)

# Get the database credentials from the config file:
configuration_file = 'db.cfg'
config = configparser.RawConfigParser(allow_no_value=True)
if not config.read(configuration_file):
    msg = 'No configuration found at %s' % configuration_file
    logging.getLogger().exception(msg)
    raise AttributeError(msg)

credentials = None
if not config.has_section('RockMakerDB'):
    msg = 'No "RockMakerDB" section in configuration found at %s' % configuration_file
    logging.getLogger().exception(msg)
    raise AttributeError(msg)
else:
    credentials = dict(config.items('RockMakerDB'))

# Connect to the database, create a cursor, actually execute the query, and write the results to an xlsx file:
with pytds.connect(**credentials) as conn:
    with conn.cursor() as c:
        c.execute(sql)

        workbook = xlsxwriter.Workbook('report_%s_%s-%s.xlsx' % (interval, start_year, start_month))
        worksheet = workbook.add_worksheet()

        i = 0
        worksheet.write(0, 0, 'barcode')
        worksheet.write(0, 1, 'project')
        worksheet.write(0, 2, 'date dispensed')
        worksheet.write(0, 3, 'imagings')
        worksheet.write(0, 4, 'user name')
        worksheet.write(0, 5, 'group name')
        worksheet.write(0, 6, 'plate type')

        for row in c.fetchall():
            i = i + 1
            j = 0
            for col in row:
                worksheet.write(i, j, col)
                j = j + 1

        workbook.close()

