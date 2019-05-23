#
# Copyright 2019 Karl Levik
#

from mssqlreport import MSSQLReport
import logging

headers = ['barcode', 'project', 'date dispensed', 'imagings',
'user name', 'group name', 'plate type', 'setup temp', 'incub. temp']

# SQL query to retrieve all plates registered and the number of times each has been imaged, within the reporting time frame
sql_template = """SELECT pl.Barcode as "{0}",
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
"""

r = MSSQLReport("RockMaker", "/tmp", "rmaker_report_")
r.set_logging(logging.DEBUG)
r.make_sql(sql_template, headers)
r.read_config("config.cfg")
r.create_report()
r.send_email()
