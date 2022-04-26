#
# Copyright 2022 Karl Levik
#
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta, date
import sys, os, copy
import pytds
import mysql.connector
import psycopg2

# Trick to make it work with both Python 2 and 3:
try:
  import configparser
except ImportError:
  import ConfigParser as configparser

class DBJob():
    """Utility methods to execute a database query with logging based on 
    configuration and command-line parameters"""

    def __init__(self, log_level=logging.DEBUG):
        if len(sys.argv) <= 1:
            msg = "No parameters"
            logging.getLogger().error(msg)
            raise AttributeError(msg)

        self.read_config(sys.argv[1])
        nowstr = str(datetime.now().strftime('%Y%m%d-%H%M%S'))
        self.working_dir = self.config['directory']
        self.fileprefix = self.job['file_prefix']
        self.set_logging(level = log_level, filepath = os.path.join(self.working_dir, '%s.log' % self.fileprefix))
        self.sql = self.job['sql']

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
        rs = None
        with psycopg2.connect(host=self.datasource['host'], dbname=self.datasource['dbname'], user=self.datasource['user'], password=self.datasource['password'], port=int(self.datasource['port'])) as conn:
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

    def read_config(self, job_section):
        """Read the email settings, job configuration and DB credentials from
        the config.cfg, reports.cfg and datasources.cfg config files"""

        self.config = self.get_section_items("config.cfg", job_section)
        self.job = self.get_section_items("jobs.cfg", job_section)
        ds_section = self.job['datasource']
        self.datasource = self.get_section_items("datasources.cfg", ds_section)

        return True
