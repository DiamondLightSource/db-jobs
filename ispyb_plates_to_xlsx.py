#
# Copyright 2019 Karl Levik
#

# Our imports:
import xlsxwriter
import mysql.connector
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta, date
import sys, os

# Trick to make it work with both Python 2 and 3:
try:
  import configparser
except ImportError:
  import ConfigParser as configparser

# Configure logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('* %(asctime)s [id=%(thread)d] <%(levelname)s> %(message)s')
hdlr = RotatingFileHandler(filename='/tmp/ispyb_plates_to_xlsx.log', maxBytes=1000000, backupCount=10)
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
start_date = '%s-%s-01' % (start_year, start_month)

# Query to retrieve all plates registered and the number of times each has been imaged, within the reporting time frame:
sql = """SELECT c.barcode as "barcode",
    concat(p.proposalCode, p.proposalNumber, '-', bs.visit_number) as "session",
    c.bltimeStamp as "date dispensed",
    count(ci.completedTimeStamp) as "imagings",
    concat(pe.givenName, ' ', pe.familyName) as "user name",
    l.name as "lab name",
    c.containerType as "plate type",
    i.temperature as "imager temp"
FROM Container c
    INNER JOIN BLSession bs ON c.sessionId = bs.sessionId
    INNER JOIN Proposal p ON bs.proposalId = p.proposalId
    INNER JOIN Person pe ON c.ownerId = pe.personId
    INNER JOIN Imager i ON c.imagerId = i.imagerId
    LEFT OUTER JOIN Laboratory l ON pe.laboratoryId = l.laboratoryId
    LEFT OUTER JOIN ContainerInspection ci ON c.containerId = ci.containerId
WHERE c.bltimeStamp >= '%s' AND c.bltimeStamp < date_add('%s', INTERVAL 1 %s)
    AND ((ci.completedTimeStamp >= '%s' AND ci.completedTimeStamp < date_add('%s', INTERVAL 1 %s)) OR ci.completedTimeStamp is NULL)
GROUP BY c.barcode,
    concat(p.proposalCode, p.proposalNumber, '-', bs.visit_number),
    c.bltimeStamp,
    concat(pe.givenName, ' ', pe.familyName),
    l.name,
    c.containerType,
    i.temperature
ORDER BY c.bltimeStamp ASC
""" % (start_date, start_date, interval, start_date, start_date, interval)

# Get the database credentials from the config file:
configuration_file = os.path.join(sys.path[0], 'config.cfg')
config = configparser.RawConfigParser(allow_no_value=True)
if not config.read(configuration_file):
    msg = 'No configuration found at %s' % configuration_file
    logging.getLogger().error(msg)
    raise AttributeError(msg)

credentials = None
if not config.has_section('ISPyBDB'):
    msg = 'No "ISPyBDB" section in configuration found at %s' % configuration_file
    logging.getLogger().error(msg)
    raise AttributeError(msg)
else:
    credentials = dict(config.items('ISPyBDB'))

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

filename = None
# Connect to the database, create a cursor, actually execute the query, and write the results to an xlsx file:
#with mysql.connector.connect(**credentials) as conn:
conn = mysql.connector.connect(host=credentials['host'], database=credentials['database'], user=credentials['user'], password=credentials['password'], port=int(credentials['port']))
if conn.is_connected():
    c = conn.cursor()
    if c is not None:
        c.execute(sql)

        filename = 'ispyb_report_%s_%s-%s.xlsx' % (interval, start_year, start_month)
        filedir = '/tmp'
        filepath = os.path.join(filedir, filename)
        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

        worksheet.set_column('A:A', 13)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('E:G', 20)
        worksheet.set_column('H:H', 11)

        worksheet.write('A1', 'barcode', bold)
        worksheet.write('B1', 'project', bold)
        worksheet.write('C1', 'date dispensed', bold)
        worksheet.write('D1', 'imagings', bold)
        worksheet.write('E1', 'user name', bold)
        worksheet.write('F1', 'group name', bold)
        worksheet.write('G1', 'plate type', bold)
        worksheet.write('H1', 'imager temp', bold)

        i = 0
        for row in c.fetchall():
            i = i + 1
            j = 0
            for col in row:
                if j != 2:
                    worksheet.write(i, j, col)
                else:
                    worksheet.write(i, j, col, date_format)
                j = j + 1

        workbook.close()
        msg = 'Report available at %s' % filepath
        print(msg)
        logging.getLogger().debug(msg)

if filepath is not None and sender is not None and recipients is not None:
    message = MIMEMultipart()
    message['Subject'] = 'ISPyB plate report for %s starting %s' % (interval, start_date)
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
