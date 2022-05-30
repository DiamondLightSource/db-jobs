# pytest configuration file

import os
import pytest
import dbjobs
import logging


@pytest.fixture(scope="session")
def testconfdir():
    """Return the directory path to the configuration files."""
    config_dir = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "conf")
    )
    if not os.path.exists(config_dir):
        pytest.skip(
            "No configuration dir found. Skipping end-to-end tests."
        )
    return config_dir

@pytest.fixture
def testdbreport(testconfdir):
    """Return a DBReport object for the test configuration."""
    yield dbjobs.create("report", job_section="FaultReport", conf_dir=testconfdir, log_level=logging.DEBUG)
