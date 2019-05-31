from dbreport import DBReport
import mysql.connector
import logging

class MariaDBReport(DBReport):

    def create_report(self, worksheet_name=None):
        # Connect to database, create cursor, execute query, write results to xlsx file:
        conn = mysql.connector.connect(host=self.credentials['host'], database=self.credentials['database'], user=self.credentials['user'], password=self.credentials['password'], port=int(self.credentials['port']))
        if conn.is_connected():
            c = conn.cursor(dictionary=True)
            if c is not None:
                c.execute(self.sql)

                self.create_xlsx(c.fetchall(), worksheet_name)
            conn.disconnect()
