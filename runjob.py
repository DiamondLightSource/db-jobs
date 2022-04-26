#
# Copyright 2022 Karl Levik
#

from dbjob import DBJob
import logging

j = DBJob(log_level=logging.DEBUG)
j.run_job()
