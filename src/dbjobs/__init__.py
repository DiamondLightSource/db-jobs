__version__ = "0.5.0"

def create(type, job_section, conf_dir, log_level):
    if type == "job":
        from dbjobs.dbjob import DBJob
        return DBJob(job_section=job_section, conf_dir=conf_dir, log_level=log_level)
    elif type == "report":
        from dbjobs.dbreport import DBReport
        return DBReport(job_section=job_section, conf_dir=conf_dir, log_level=log_level)
    else:
        raise AttributeError(
            f"{type} is not a supported dbjob type. Supported types are 'job' and 'report'"
        )
