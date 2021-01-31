#
# Copyright 2021 Karl Levik
#

from dbreport import DBReport
import logging

import sys

r = DBReport(log_level=logging.DEBUG)
r.make_sql()
r.create_report()
r.send_email()
