#
# Copyright 2019 Karl Levik
#

import xlsxwriter
import pytds
from dbreports import DBReports
import logging
from datetime import datetime, timedelta, date
import os, copy

def make_sql(start_date, interval):
    """SQL query to retrieve all plates registered and the number of times each has been imaged, within the reporting time frame"""

    headers = ['barcode', 'project', 'date dispensed', 'imagings',
    'user name', 'group name', 'plate type', 'setup temp', 'incub. temp']

    fmt = copy.deepcopy(headers)
    fmt.append(start_date)
    fmt.append(interval)

    sql = """SELECT pl.Barcode as "{0}",
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

    return (sql, headers)


def create_report(filedir, filename, credentials, sql, field_names):
    filepath = os.path.join(filedir, filename)

    # Connect to database, create cursor, execute query, write results to xlsx file:
    with pytds.connect(
        dsn = credentials['dsn'],
        database = credentials['database'],
        user = credentials['user'],
        password = credentials['password'],
        as_dict = True
        ) as conn:
        with conn.cursor() as c:
            c.execute(sql)

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
            msg = "Report available at %s" % filepath
            print(msg)
            logging.getLogger().debug(msg)

r = DBReports("RockMaker", "/tmp", "rmaker_report_")
r.set_logging()
(sql, field_names) = make_sql(r.start_date, r.interval)
r.read_config()
create_report(r.filedir, r.filename, r.credentials, sql, field_names)
r.send_email()
