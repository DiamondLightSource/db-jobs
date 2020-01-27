#
# Copyright 2019 Karl Levik
#
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta, date
import xlsxwriter
import sys, os, copy

# Trick to make it work with both Python 2 and 3:
try:
  import configparser
except ImportError:
  import ConfigParser as configparser

class DBReport():
    """Utility methods to create a report and send it as en email attachment"""

    def __init__(self, working_dir, fileprefix, config_file, db_section, email_section=None, log_level=logging.DEBUG, filesuffix='xlsx'):
        self.get_parameters()
        self.working_dir = working_dir
        self.fileprefix = fileprefix
        self.filesuffix = filesuffix
        self.filename = '%s%s_%s-%s.%s' % (fileprefix, self.interval, self.start_year, self.start_month, filesuffix)
        self.set_logging(log_level)
        self.read_config(config_file, db_section, email_section)

    def get_parameters(self):
        # Get input parameters, otherwise use default values
        self.interval = 'month'

        today = date.today()
        first = today.replace(day=1)
        prev_date = first - timedelta(days=1)

        if len(sys.argv) > 1:
            self.interval = sys.argv[1]

        if len(sys.argv) >= 1:
            if self.interval == 'month':
                self.start_year = prev_date.year
                self.start_month = prev_date.month
            elif self.interval == 'year':
                self.start_year = prev_date.year - 1
                self.start_month = prev_date.month
            else:
                err_msg = 'interval must be "month" or "year"'
                logging.getLogger().error(err_msg)
                raise AttributeError(err_msg)

            if len(sys.argv) > 2:
                self.start_year = sys.argv[2]  # e.g. 2018
                if len(sys.argv) > 3:
                    self.start_month = sys.argv[3] # e.g. 02

        self.start_date = '%s/%s/01' % (self.start_year, self.start_month)

    def set_logging(self, level):
        """Configure logging"""
        filepath = os.path.join(self.working_dir, '%s%s_%s-%s.log' % (self.fileprefix, self.interval, self.start_year, self.start_month))
        logger = logging.getLogger()
        logger.setLevel(level)
        formatter = logging.Formatter('* %(asctime)s [id=%(thread)d] <%(levelname)s> %(message)s')
        hdlr = RotatingFileHandler(filename=filepath, maxBytes=1000000, backupCount=10)
        hdlr.setFormatter(formatter)
        logging.getLogger().addHandler(hdlr)

    def read_config(self, config_file, db_section, email_section=None):
        # Get the database credentials and email settings from the config file:
        configuration_file = os.path.join(sys.path[0], config_file)
        config = configparser.RawConfigParser(allow_no_value=True)
        if not config.read(configuration_file):
            msg = 'No configuration found at %s' % configuration_file
            logging.getLogger().error(msg)
            raise AttributeError(msg)

        self.credentials = None
        if not config.has_section(db_section):
            msg = 'No "RockMakerDB" section in configuration found at %s' % configuration_file
            logging.getLogger().error(msg)
            raise AttributeError(msg)
        else:
            self.credentials = dict(config.items(db_section))

        self.sender = None
        self.recipients = None
        if email_section is None or not config.has_section(email_section):
            msg = 'No "%s" section in configuration found at %s' % (email_section, configuration_file)
            logging.getLogger().error(msg)
            raise AttributeError(msg)
        else:
            email_settings = dict(config.items(email_section))
            self.sender = email_settings['sender']
            self.recipients = email_settings['recipients']

        return True

    def make_sql(self, sql_template, headers):
        """Create proper SQL from the template - merge the headers in as aliases"""
        self.headers = headers
        fmt = copy.deepcopy(headers)
        fmt.append(self.start_date)
        fmt.append(self.interval)
        self.sql = sql_template.format(*fmt)

    def create_xlsx(self, result_set, worksheet_name=None):
        filepath = os.path.join(self.working_dir, self.filename)

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet(worksheet_name)

        bold = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

        # Pre-populate the max lengths for each column
        # with the lenth of the header
        max_lengths = []
        for header in self.headers:
            max_lengths.append(len(header))

        # Populate worksheet columns with values from DB result set.
        # Keep track of the max lengths for each column.
        i = 0
        for row in result_set:

            i += 1
            j = 0
            for header in self.headers:
                field_value = row[header]

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
        for header in self.headers:
            worksheet.write(0, j, header, bold)
            worksheet.set_column(j, j, max_lengths[j] + 1)
            j += 1

        workbook.close()
        msg = "Report available at %s" % filepath
        print(msg)
        logging.getLogger().debug(msg)

    def create_csv(self, result_set, worksheet_name=None):
        filepath = os.path.join(self.working_dir, self.filename)

        with open(filepath, 'w') as f:

            # Write comma-separated column headers.
            i = 1
            for header in self.headers:
                f.write(header)
                if i < len(self.headers):
                    f.write(",")
                i += 1

            if len(self.headers) > 0:
                f.write("\n")

            # Write each row with comma-separated fields.
            for row in result_set:
                i = 1
                for header in self.headers:
                    field_value = str(row[header])
                    if "," in field_value:
                        field_value = "\"%s\"" % field_value
                    f.write(field_value)
                    if i < len(self.headers):
                        f.write(",")
                    i += 1
                f.write("\n")

        msg = "Report available at %s" % filepath
        print(msg)
        logging.getLogger().debug(msg)


    def send_email(self, report_name):
        if self.working_dir is not None and self.filename is not None and self.sender is not None and self.recipients is not None:
            filepath = os.path.join(self.working_dir, self.filename)

            message = MIMEMultipart()
            message['Subject'] = '%s for %s starting %s' % (report_name, self.interval, self.start_date)
            message['From'] = self.sender
            message['To'] = self.recipients
            body = 'Please find the report attached.'
            message.attach(MIMEText(body, 'plain'))

            with open(filepath, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())

            encoders.encode_base64(part)

            part.add_header(
                'Content-Disposition',
                'attachment; filename= %s' % self.filename,
            )

            message.attach(part)
            text = message.as_string()

            if self.recipients is not None and self.recipients != "":
                try:
                    server = smtplib.SMTP('localhost', 25) # or 587?
                    #server.login('youremailusername', 'password')

                    # Send the mail
                    recipients_list = []
                    for i in self.recipients.split(','):
                        recipients_list.append(i.strip())
                    server.sendmail(self.sender, recipients_list, text)
                except:
                    err_msg = 'Failed to send email'
                    logging.getLogger().exception(err_msg)
                    print(err_msg)

                logging.getLogger().debug('Email sent')
