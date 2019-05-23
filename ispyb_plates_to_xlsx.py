#
# Copyright 2019 Karl Levik
#

# Our imports:
from mariadbreport import MariaDBReport
import logging

headers = ['barcode', 'session', 'date dispensed', 'imagings',
'user name', 'lab name', 'plate type', 'imager temp']

# Query to retrieve all plates registered and the number of times each has been imaged, within the reporting time frame:
sql_template = """SELECT c.barcode as "{0}",
    concat(p.proposalCode, p.proposalNumber, '-', bs.visit_number) as "{1}",
    c.bltimeStamp as "{2}",
    count(ci.completedTimeStamp) as "{3}",
    concat(pe.givenName, ' ', pe.familyName) as "{4}",
    l.name as "{5}",
    c.containerType as "{6}",
    i.temperature as "{7}"
FROM Container c
    INNER JOIN BLSession bs ON c.sessionId = bs.sessionId
-- LEFT OUTER JOIN Dewar ... LEFT OUTER JOIN Shipping ... LEFT OUTER JOIN LabContact LEFT OUTER JOIN Laboratory ??????????????
    INNER JOIN Proposal p ON bs.proposalId = p.proposalId
    INNER JOIN Person pe ON c.ownerId = pe.personId
    INNER JOIN Imager i ON c.imagerId = i.imagerId
    LEFT OUTER JOIN Laboratory l ON pe.laboratoryId = l.laboratoryId
    LEFT OUTER JOIN ContainerInspection ci ON c.containerId = ci.containerId
WHERE c.bltimeStamp >= '{8}' AND c.bltimeStamp < date_add('{8}', INTERVAL 1 {9})
    AND ((ci.completedTimeStamp >= '{8}' AND ci.completedTimeStamp < date_add('{8}', INTERVAL 1 {9})) OR ci.completedTimeStamp is NULL)
GROUP BY c.barcode,
    concat(p.proposalCode, p.proposalNumber, '-', bs.visit_number),
    c.bltimeStamp,
    concat(pe.givenName, ' ', pe.familyName),
    l.name,
    c.containerType,
    i.temperature
ORDER BY c.bltimeStamp ASC
"""

r = MariaDBReport("ISPyB", "/tmp", "ispyb_report_")
r.set_logging(logging.DEBUG)
r.make_sql(sql_template, headers)
r.read_config("config.cfg", "ISPyBDB")
r.create_report()
r.send_email()
