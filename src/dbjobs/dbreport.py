#
# Copyright 2019 Karl Levik
#
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import logging
from datetime import datetime, timedelta, date
import xlsxwriter
import sys, os, copy
import pytds
import mysql.connector
import psycopg2
from dbjobs.dbjob import DBJob

import configparser

class DBReport(DBJob):
    """Utility methods to create a report and send it as en email attachment"""

    def __init__(self, job_section=None, conf_dir=None, log_level=logging.DEBUG):
        super().__init__(job_section=job_section, conf_dir=conf_dir, log_level=log_level)
        self.filesuffix = self.job['file_suffix']

    def get_start_date(self, interval=None, start_year=None, start_month=None):
        # Get input parameters, otherwise use default values
        today = date.today()
        first = today.replace(day=1)
        prev_date = first - timedelta(days=1)

        if interval:
            self.interval = interval
        else:
            self.interval = "month"

        if self.interval == "month": 
            self.start_year = prev_date.year
            self.start_month = prev_date.month

        elif self.interval == "year":
            self.start_year = prev_date.year - 1
            self.start_month = prev_date.month

        else:
            self.error('interval must be "month" or "year"')
        
        if start_year:
            self.start_year = start_year
        if start_month:
            self.start_month = start_month

        return f'{self.start_year}/{self.start_month}/01'

    def read_config(self, job_section, conf_dir):
        super().read_config(job_section, conf_dir=conf_dir)
        self.sender = self.config['sender']
        self.recipients = self.config['recipients']

    def make_sql(self, interval, start_year, start_month):
        """Create proper SQL from the template - merge the headers in as aliases"""
        self.headers = self.job['sql_headers'].split(',')
        fmt = copy.deepcopy(self.headers)

        self.start_date = self.get_start_date(interval, start_year, start_month)
        fmt.append(self.start_date)
        fmt.append(self.interval)
        nowstr = str(datetime.now().strftime('%Y%m%d-%H%M%S'))
        self.filename = '%s%s_%s-%s_%s%s' % (self.fileprefix, self.interval, self.start_year, self.start_month, nowstr, self.filesuffix)
        self.sql = self.job['sql'].format(*fmt)

    def run_job(self):
        rs = super().run_job()
        if self.job['file_suffix'] == ".xlsx":
            self.create_xlsx(rs)
        elif self.job['file_suffix'] == ".csv":
            self.create_csv(rs)
        else:
            msg = 'Unknown file suffix: %s' % self.job['file_suffix']
            logging.getLogger().error(msg)
            raise AttributeError(msg)

    def create_xlsx(self, result_set):
        filepath = os.path.join(self.working_dir, self.filename)

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet(self.job['worksheet_name'])

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

    def create_csv(self, result_set):
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


    def send_email(self, subject):
        report_name = self.job["fullname"]
        attach_report = True
        if self.config["attach"].lower() == "no":
            attach_report = False
        if self.working_dir is not None and self.filename is not None and self.sender is not None and self.recipients is not None:
            filepath = os.path.join(self.working_dir, self.filename)

            message = MIMEMultipart()
            message['Subject'] = subject
            message['From'] = self.sender
            message['To'] = self.recipients

            if attach_report:
                body = 'Please find the report attached.'
            else:
                body = 'Please find the report at %s.' % filepath
            message.attach(MIMEText(body, 'plain'))

            if attach_report:
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
