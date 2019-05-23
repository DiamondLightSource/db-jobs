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
import sys, os
# Trick to make it work with both Python 2 and 3:
try:
  import configparser
except ImportError:
  import ConfigParser as configparser

class DBReports():
    """Utility methods to create a report and send it as en email attachment"""

    def __init__(self, reportname, filedir, fileprefix):
        self.get_parameters()
        self.reportname = reportname
        self.filedir = filedir
        self.fileprefix = fileprefix
        self.filename = '%s%s_%s-%s.xlsx' % (fileprefix, self.interval, self.start_year, self.start_month)

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

    def set_logging(self):
        """Configure logging"""
        filepath = os.path.join(self.filedir, '%s_%s_%s-%s.log' % (self.fileprefix, self.interval, self.start_year, self.start_month))
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('* %(asctime)s [id=%(thread)d] <%(levelname)s> %(message)s')
        hdlr = RotatingFileHandler(filename=filepath, maxBytes=1000000, backupCount=10)
        hdlr.setFormatter(formatter)
        logging.getLogger().addHandler(hdlr)

    def read_config(self):
        # Get the database credentials and email settings from the config file:
        configuration_file = os.path.join(sys.path[0], 'config.cfg')
        config = configparser.RawConfigParser(allow_no_value=True)
        if not config.read(configuration_file):
            msg = 'No configuration found at %s' % configuration_file
            logging.getLogger().error(msg)
            raise AttributeError(msg)

        self.credentials = None
        if not config.has_section('RockMakerDB'):
            msg = 'No "RockMakerDB" section in configuration found at %s' % configuration_file
            logging.getLogger().error(msg)
            raise AttributeError(msg)
        else:
            self.credentials = dict(config.items('RockMakerDB'))

        self.sender = None
        self.recipients = None
        if not config.has_section('Email'):
            msg = 'No "Email" section in configuration found at %s' % configuration_file
            logging.getLogger().error(msg)
            raise AttributeError(msg)
        else:
            email_settings = dict(config.items('Email'))
            self.sender = email_settings['sender']
            self.recipients = email_settings['recipients']

        return True

    def send_email(self):
        if self.filedir is not None and self.filename is not None and self.sender is not None and self.recipients is not None:
            filepath = os.path.join(self.filedir, self.filename)

            message = MIMEMultipart()
            message['Subject'] = '%s plate report for %s starting %s' % (self.reportname, self.interval, self.start_date)
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
