#
# Copyright 2019 Karl Levik
#

# Our imports:
from mariadbreport import MariaDBReport
import logging

headers = ['beamline', 'start time', 'end time', 'visit', 'system', 'component', 'subcomponent', 'lost hrs', 'total samples', 'collected samples', 'title', 'local contact(s)', 'URL' ]

sql_template = """SELECT bs.beamLineName "{0}",
f.beamtimelost_starttime "{1}",
f.beamtimelost_endtime "{2}",
concat(p.proposalCode, p.proposalNumber, '-', bs.visit_number) "{3}",
s.name "{4}",
c.name "{5}",
sc.name "{6}",
(f.beamtimelost_endtime - f.beamtimelost_starttime)/3600 "{7}",
(SELECT count(*) FROM Dewar d INNER JOIN Container con ON con.dewarId = d.dewarId INNER JOIN BLSample bls ON bls.containerId = con.containerId WHERE d.firstExperimentId = bs.sessionId) "{8}",
(SELECT count(distinct(dc.blSampleId)) FROM DataCollection dc WHERE bs.sessionId = dc.sessionId) "{9}",
f.title "{10}",
group_concat(per.givenName, ' ', per.familyName SEPARATOR ', ') "{11}",
concat('https://ispyb.diamond.ac.uk/faults/fid/', f.faultId) "{12}"
FROM BF_fault f
  INNER JOIN BLSession bs USING(sessionId)
  INNER JOIN Proposal p USING(proposalId)
  LEFT OUTER JOIN BF_subcomponent sc USING(subcomponentId)
  LEFT OUTER JOIN BF_component c USING(componentId)
  LEFT OUTER JOIN BF_system s USING(systemId)
  LEFT OUTER JOIN Session_has_Person shp ON shp.sessionId = bs.sessionId AND shp.role LIKE 'Local Contact%'
  LEFT OUTER JOIN Person per ON per.personId = shp.personId
WHERE f.beamtimelost_starttime > now() - INTERVAL 14 DAY
GROUP BY bs.beamLineName, f.beamtimelost_starttime, f.beamtimelost_endtime, concat(p.proposalCode, p.proposalNumber, '-', bs.visit_number), s.name, c.name, sc.name, f.title, f.faultId
ORDER BY bs.beamLineName, f.beamtimelost_starttime
"""

format = "csv"

r = MariaDBReport(
    "/tmp",
    "fault_report_",
    "config.cfg",
    "ISPyBDB",
    email_section="FaultEmails",
    log_level=logging.DEBUG,
    filesuffix=format
)
r.make_sql(sql_template, headers)
r.create_report("faults", format=format)
r.send_email("Fault report", attach_report=False)
