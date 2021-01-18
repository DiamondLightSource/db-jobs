#
# Copyright 2019 Karl Levik
#

from mariadbreport import MariaDBReport
import logging

r = MariaDBReport(
    "config.cfg",
    db_section="ISPyBDB",
    report_section="FaultReport",
    email_section="FaultEmails",
    log_level=logging.DEBUG,
)
r.make_sql()
r.create_report()
r.send_email("Fault report", attach_report=False)
