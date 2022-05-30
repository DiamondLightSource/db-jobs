#
# Copyright 2022 Karl Levik
#
import logging
import atexit
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta, date
import sys, os, copy, os.path
from dbjobs.base import Base
try:
    import pytds
except ImportError:
    pytds = None
try:
    import mysql.connector
except ImportError:
    mysql = None
try:
    import psycopg2
except ImportError:
    psycopg2 = None

import configparser

class DBJob(Base):
    """Utility methods to execute a database query with logging based on 
    configuration and command-line parameters"""

    def __init__(self, job_section=None, conf_dir=None, log_level=logging.DEBUG):
        if job_section is None:
            self.error("Job section is required.")

        self.read_config(job_section, conf_dir)
        self.working_dir = self.config['directory']
        self.fileprefix = self.job['file_prefix']
        self.set_logging(level = log_level, filepath = os.path.join(self.working_dir, f"{self.fileprefix}.log"))
        logging.getLogger().info("DBJob %s started" % self.job['fullname'])
        atexit.register(self.clean_up)
        self.sql = self.job['sql']

    def clean_up(self):
        if hasattr(self, 'job'):
            logging.getLogger().info("DBJob %s completed" % self.job['fullname'])
            logging.shutdown()

    def run_job(self):
        rs = None
        if self.datasource["dbtype"] == "MariaDB":
            rs = self.run_mariadb_job()
        elif self.datasource["dbtype"] == "MSSQL":
            rs = self.run_mssql_job()
        elif self.datasource["dbtype"] == "PostgreSQL":
            rs = self.run_postgresql_job()
        else:
            msg = "Unknown dbtype: %s" % self.datasource["dbtype"]
            logging.getLogger().error(msg)
            raise AttributeError(msg)
        return rs

    def run_mariadb_job(self):
        """Connect to database, create cursor, execute query, disconnect, return
        result set."""

        if mysql is None:
            logging.getLogger().error("mysql.connector not found")
            sys.exit(1)

        rs = None
        conn = mysql.connector.connect(host=self.datasource['host'],
            database=self.datasource['database'], 
            user=self.datasource['user'], 
            password=self.datasource['password'], 
            port=int(self.datasource['port'])
        )
        if conn.is_connected():
            c = conn.cursor(dictionary=True)
            if c is not None:
                c.execute(self.sql)
                if self.job['sql_type'] in ('read', 'read-write'):
                    rs = c.fetchall()
            conn.disconnect()
        return rs

    def run_mssql_job(self):
        """Connect to database, create cursor, execute query, disconnect, return
        result set."""

        if pytds is None:
            logging.getLogger().error("pytds not found")
            sys.exit(1)

        rs = None
        with pytds.connect(
            dsn = self.datasource['dsn'],
            database = self.datasource['database'],
            user = self.datasource['user'],
            password = self.datasource['password'],
            as_dict = True
            ) as conn:
            with conn.cursor() as c:
                c.execute(self.sql)
                if self.job['sql_type'] in ('read', 'read-write'):
                    rs = c.fetchall()
        return rs

    def run_postgresql_job(self):
        """Connect to database, create cursor, execute query, disconnect, return
        result set."""

        if psycopg2 is None:
            logging.getLogger().error("psycopg2 not found")
            sys.exit(1)

        rs = None
        with psycopg2.connect(host=self.datasource['host'],
            dbname=self.datasource['dbname'],
            user=self.datasource['user'],
            password=self.datasource['password'],
            port=int(self.datasource['port'])
        ) as conn:
            conn.set_session(readonly=True, autocommit=True)
            if not conn.closed:
                with conn.cursor(dictionary=True) as c:
                    if c is not None:
                        c.execute(self.sql)
                        if self.job['sql_type'] in ('read', 'read-write'):
                            rs = c.fetchall()
        return rs

    def set_logging(self, level=None, filepath=None):
        """Configure logging"""
        logger = logging.getLogger()
        logger.setLevel(level)
        formatter = logging.Formatter('* %(asctime)s [id=%(thread)d] <%(levelname)s> %(message)s')
        hdlr = RotatingFileHandler(filename=filepath, maxBytes=1000000, backupCount=10)
        hdlr.setFormatter(formatter)
        logging.getLogger().addHandler(hdlr)

    def get_section_items(self, conf_file, conf_section):
        config_file = os.path.join(sys.path[0], conf_file)
        config = configparser.RawConfigParser(allow_no_value=True)
        if not config.read(config_file):
            msg = 'No configuration found at %s' % config_file
            logging.getLogger().error(msg)
            raise AttributeError(msg)

        if not config.has_section(conf_section):
            msg = 'No section %s in configuration found at %s' % (conf_section, config_file)
            logging.getLogger().error(msg)
            raise AttributeError(msg)

        return dict(config.items(conf_section))

    def read_config(self, job_section, conf_dir=None):
        """Read the email settings, job configuration and DB credentials from
        the config.cfg, reports.cfg and datasources.cfg config files"""

        conf_file = "config.cfg"
        jobs_file = "jobs.cfg"
        ds_file = "datasources.cfg"

        if conf_dir:
            conf_file = os.path.join(conf_dir, conf_file)
            jobs_file = os.path.join(conf_dir, jobs_file)
            ds_file = os.path.join(conf_dir, ds_file)

        self.config = self.get_section_items(conf_file, job_section)
        self.job = self.get_section_items(jobs_file, job_section)
        ds_section = self.job['datasource']
        self.datasource = self.get_section_items(ds_file, ds_section)

        return True
