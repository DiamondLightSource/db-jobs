import time


def test_report(testdbreport):
    testdbreport.make_sql("month", "2022", "01")
    testdbreport.run_job()
    # testdbreport.send_email()

    # assert that output file exists in expected dir and so on ...

