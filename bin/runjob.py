#
# Copyright 2022 Karl Levik
#

import argparse
import logging
import dbjobs

parser = argparse.ArgumentParser()
parser.add_argument("-j", "--job", help = "Specify job")
parser.add_argument("-d", "--dir", help = "Specify config dir")

args = parser.parse_args()

j = dbjobs.create("job", job_section=args.job, conf_dir=args.dir, log_level=logging.DEBUG)
j.run_job()
