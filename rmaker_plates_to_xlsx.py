#
# Copyright 2019 Karl Levik
#

# Our imports:
import xlsxwriter
import pytds
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta, date
import sys, os, copy
# Trick to make it work with both Python 2 and 3:
try:
  import configparser
except ImportError:
  import ConfigParser as configparser


def make_sql(headers):
    fmt = copy.deepcopy(headers)
    fmt.append(start_date)
    fmt.append(interval)

    return """SELECT pl.Barcode as "{0}",
        tn4.Name as "{1}",
        pl.DateDispensed as "{2}",
        count(it.DateImaged) as "{3}",
        u.Name as "{4}",
        u.ID,
        STUFF((
              SELECT ', ' + g.Name
              FROM Groups g
                INNER JOIN GroupUser gu ON g.ID = gu.GroupID
              WHERE u.ID = gu.UserID AND g.Name <> 'AllRockMakerUsers'
              FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') as "{5}",
        c.Name as "{6}",
        st.Temperature as "{7}",
        itemp.Temperature as "{8}"
    FROM Plate pl
        INNER JOIN Experiment e ON pl.ExperimentID = e.ID
        INNER JOIN Containers c ON e.ContainerID = c.ID
        INNER JOIN Users u ON e.userID = u.ID
        INNER JOIN TreeNode tn1 ON pl.TreeNodeID = tn1.ID
        INNER JOIN TreeNode tn2 ON tn1.ParentID = tn2.ID
        INNER JOIN TreeNode tn3 ON tn2.ParentID = tn3.ID
        INNER JOIN TreeNode tn4 ON tn3.ParentID = tn4.ID
        INNER JOIN SetupTemp st ON e.SetupTempID = st.ID
        INNER JOIN IncubationTemp itemp ON e.IncubationTempID = itemp.ID
        LEFT OUTER JOIN ExperimentPlate ep ON ep.PlateID = pl.ID
        LEFT OUTER JOIN ImagingTask it ON it.ExperimentPlateID = ep.ID
    WHERE pl.DateDispensed >= convert(date, '{9}', 111) AND pl.DateDispensed < dateadd({10}, 1, convert(date, '{9}', 111))
        AND ((it.DateImaged >= convert(date, '{9}', 111) AND it.DateImaged < dateadd({10}, 1, convert(date, '{9}', 111))) OR it.DateImaged is NULL)
    GROUP BY pl.Barcode,
        tn4.Name,
        pl.DateDispensed,
        u.Name,
        u.ID,
        c.Name,
        st.Temperature,
        itemp.Temperature
    ORDER BY pl.DateDispensed ASC
    """.format(*fmt)

def send_email(filepath, sender, recipients, interval, start_date):
    message = MIMEMultipart()
    message['Subject'] = 'RockMaker plate report for %s starting %s' % (interval, start_date)
    message['From'] = sender
    message['To'] = recipients
    body = 'Please find the report attached.'
    message.attach(MIMEText(body, 'plain'))

    with open(filepath, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header(
        'Content-Disposition',
        'attachment; filename= %s' % filename,
    )

    message.attach(part)
    text = message.as_string()

    if recipients is not None and recipients != "":
        try:
            server = smtplib.SMTP('localhost', 25) # or 587?
            #server.login('youremailusername', 'password')

            # Send the mail
            recipients_list = []
            for i in recipients.split(','):
                recipients_list.append(i.strip())
            server.sendmail(sender, recipients_list, text)
        except:
            err_msg = 'Failed to send email'
            logging.getLogger().exception(err_msg)
            print(err_msg)

        logging.getLogger().debug('Email sent')


def read_config():
    # Get the database credentials and email settings from the config file:
    configuration_file = os.path.join(sys.path[0], 'config.cfg')
    config = configparser.RawConfigParser(allow_no_value=True)
    if not config.read(configuration_file):
        msg = 'No configuration found at %s' % configuration_file
        logging.getLogger().error(msg)
        raise AttributeError(msg)

    credentials = None
    if not config.has_section('RockMakerDB'):
        msg = 'No "RockMakerDB" section in configuration found at %s' % configuration_file
        logging.getLogger().error(msg)
        raise AttributeError(msg)
    else:
        credentials = dict(config.items('RockMakerDB'))

    sender = None
    recipients = None
    if not config.has_section('Email'):
        msg = 'No "Email" section in configuration found at %s' % configuration_file
        logging.getLogger().error(msg)
        raise AttributeError(msg)
    else:
        email_settings = dict(config.items('Email'))
        sender = email_settings['sender']
        recipients = email_settings['recipients']

    return (credentials, sender, recipients)


# Configure logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('* %(asctime)s [id=%(thread)d] <%(levelname)s> %(message)s')
hdlr = RotatingFileHandler(filename='/tmp/rmaker_plates_to_xlsx.log', maxBytes=1000000, backupCount=10)
hdlr.setFormatter(formatter)
logging.getLogger().addHandler(hdlr)

# Get input parameters, otherwise use default values
interval = 'month'

today = date.today()
first = today.replace(day=1)
prev_date = first - timedelta(days=1)

start_year = prev_date.year
start_month = prev_date.month
if len(sys.argv) > 1:
    interval = sys.argv[1]
    if interval not in ('month', 'year'):
        err_msg = 'interval must be "month" or "year"'
        logging.getLogger().error(err_msg)
        raise AttributeError(err_msg)
if len(sys.argv) > 2:
    start_year = sys.argv[2]  # e.g. 2018
if len(sys.argv) > 3:
    start_month = sys.argv[3] # e.g. 02
start_date = '%s/%s/01' % (start_year, start_month)

# Query to retrieve all plates registered and the number of times each has been imaged, within the reporting time frame:
field_names = ['barcode', 'project', 'date dispensed', 'imagings', 'user name',
    'group name', 'plate type', 'setup temp', 'incub. temp']

sql = make_sql(field_names)

(credentials, sender, recipients) = read_config()

filename = None
# Connect to the database, create a cursor, actually execute the query, and write the results to an xlsx file:
with pytds.connect(
    dsn = credentials['dsn'],
    database = credentials['database'],
    user = credentials['user'],
    password = credentials['password'],
    as_dict = True
    ) as conn:
    with conn.cursor() as c:
        c.execute(sql)

        filename = 'rmaker_report_%s_%s-%s.xlsx' % (interval, start_year, start_month)
        filedir = '/tmp'
        filepath = os.path.join(filedir, filename)
        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

        # Pre-populate the max lengths for each column
        # with the lenth of the header
        max_lengths = []
        for field_name in field_names:
            max_lengths.append(len(field_name))

        # Populate the worksheet columns with values from the DB result set.
        # Keep track of the max lengths for each column.
        i = 0
        for row in c.fetchall():

            i += 1
            j = 0
            for field_name in field_names:
                field_value = row[field_name]

                if isinstance(field_value, datetime):
                    worksheet.write(i, j, field_value, date_format)
                    s = str(field_value)
                    # disregard chars after dot when finding length
                    max_lengths[j] = len(s[:s.rfind('.')])
                else:
                    worksheet.write(i, j, field_value)
                    if len(str(field_value)) > max_lengths[j]:
                        max_lengths[j] = len(str(field_value))

                j += 1

        # Populate the column headers in the worksheet.
        # Set the column widths to the max length used in each column.
        j = 0
        for field_name in field_names:
            worksheet.write(0, j, field_name, bold)
            worksheet.set_column(j, j, max_lengths[j] + 1)
            j += 1

        workbook.close()
        msg = 'Report available at %s' % filepath
        print(msg)
        logging.getLogger().debug(msg)

if filepath is not None and sender is not None and recipients is not None:
    send_email(filepath, sender, recipients, interval, start_date)
