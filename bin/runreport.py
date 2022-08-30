#!/usr/bin/env python

#
# Copyright 2021 Karl Levik
#
import argparse
import logging
import dbjobs

parser = argparse.ArgumentParser()
parser.add_argument("-j", "--job", help = "Specify job")
parser.add_argument("-d", "--dir", help = "Specify config dir")
parser.add_argument("-i", "--interval", help = "Specify interval")
parser.add_argument("-m", "--start_month", help = "Specify start month")
parser.add_argument("-y", "--start_year", help = "Specify start year")

args = parser.parse_args()
#r = DBReport(job_section=args.job, interval=args.interval, start_year=args.start_year, start_month=self.start_month, log_level=logging.DEBUG)
r = dbjobs.create("report", job_section=args.job, conf_dir=args.dir, log_level=logging.DEBUG)
r.make_sql(args.interval, args.start_year, args.start_month)
r.run_job()
r.send_email(f'{args.job} for {r.interval} starting {r.start_date}')
