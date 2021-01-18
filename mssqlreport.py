from dbreport import DBReport
import pytds
import logging

class MSSQLReport(DBReport):

    def create_report(self):
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

                if self.report['format'] == "xlsx":
                    self.create_xlsx(c.fetchall())
                elif self.report['format'] == "csv":
                    self.create_csv(c.fetchall())
                else:
                    msg = 'Unknown format: %s' % self.report['format']
                    logging.getLogger().error(msg)
                    raise AttributeError(msg)
