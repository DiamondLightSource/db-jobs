from dbreport import DBReport
import xlsxwriter
import mysql.connector
import os
from datetime import datetime
import logging

class MariaDBReport(DBReport):

    def create_report(self):
        filepath = os.path.join(self.filedir, self.filename)

        # Connect to database, create cursor, execute query, write results to xlsx file:
        conn = mysql.connector.connect(host=self.credentials['host'], database=self.credentials['database'], user=self.credentials['user'], password=self.credentials['password'], port=int(self.credentials['port']))
        if conn.is_connected():
            c = conn.cursor(dictionary=True)
            if c is not None:
                c.execute(self.sql)

                workbook = xlsxwriter.Workbook(filepath)
                worksheet = workbook.add_worksheet()

                bold = workbook.add_format({'bold': True})
                date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

                # Pre-populate the max lengths for each column
                # with the lenth of the header
                max_lengths = []
                for header in self.headers:
                    max_lengths.append(len(header))

                # Populate the worksheet columns with values from the DB result set.
                # Keep track of the max lengths for each column.
                i = 0
                for row in c.fetchall():

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
