#
# Copyright 2021 Karl Levik
#

from dbreport import DBReport
import logging

r = DBReport(log_level=logging.DEBUG)
r.make_sql()
r.run_job()
r.send_email()
