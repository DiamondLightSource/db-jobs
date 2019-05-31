from dbreport import DBReport
import pytds
import logging

class MSSQLReport(DBReport):

    def create_report(self, worksheet_name=None):
        # Connect to database, create cursor, execute query, write results to xlsx file:
        with pytds.connect(
            dsn = self.credentials['dsn'],
            database = self.credentials['database'],
            user = self.credentials['user'],
            password = self.credentials['password'],
            as_dict = True
            ) as conn:
            with conn.cursor() as c:
                c.execute(self.sql)

                self.create_xlsx(c.fetchall(), worksheet_name)
