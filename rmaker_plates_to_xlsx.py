#
# Copyright 2019 Karl Levik
#

from mssqlreport import MSSQLReport
import logging

r = MSSQLReport(
    "config.cfg",
    db_section="RockMakerDB",
    report_section="RockMakerPlateReport",
    email_section="RockMakerPlatesEmails",
    log_level=logging.DEBUG
)
r.make_sql()
r.create_report()
r.send_email("RockMaker plate report")
