__author__ = "David Marienburg"
__version__ = "1.11"

"""
This script is for processing the Housing Services - Housing Outcomes v2.0 report that is used by
the follow-ups specialistself.

This script should identify when it is being run and adjust its date parameters to reflect being run
on the first of the month or after it.  This well prevent future staff members from needing to
fiddle with the code every time they run the report.
"""

import pandas as pd
from datetime import date
from datetime import datetime
from calendar import monthrange
from dateutil.relativedelta import relativedelta
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class RunDate:
    """
    This class is not currently used by the CreateRequiredFollowUps class but I would like to
    eventually make it so that the sheetnames output by that class are modified by the relation to
    the current month.
    """
    def __init__(self):
        self.today = datetime.now().date()
        self.check_date()

    def check_date(self):
        if self.today.day <= 5:
            last_month = self.today + relativedelta(months=-1)
            end_of_month = date(
                year=last_month.year,
                month=last_month.month,
                day=monthrange(last_month.year, last_month.month)[1]
            )
            return end_of_month
        else:
            end_of_month = date(
                year=self.today.year,
                month=self.today.month,
                day=monthrange(self.today.year, self.today.month)[1]
            )
            return end_of_month


class CreateRequiredFollowUps:
    def __init__(self, file_path):
        self.raw_data = pd.read_excel(file_path)
        self.run_date = RunDate()
        self.month_range = set(
            [value.strftime("%B") for value in self.raw_data["Follow Up Due Date(2512)"]]
        )
        self.current_month = datetime.now().month

    def process(self):
        data = self.raw_data
        writer = pd.ExcelWriter(
            asksaveasfilename(title="Save the Required Follow-ups Report"),
            engine="xlsxwriter"
        )
        for month in self.month_range:
            month_data = data[
                (data["Follow Up Due Date(2512)"].dt.strftime("%B") == month) &
                data["Actual Follow Up Date(2518)"].isna()
            ].drop_duplicates(subset="Client Uid")
            month_data.to_excel(writer, sheet_name="{} Follow-Ups".format(month), index=False)
        data.to_excel(writer, sheet_name="Raw Data", index=False)
        writer.save()

if __name__ == "__main__":
    run = CreateRequiredFollowUps(askopenfilename(title="Open the Housing Outcomes v2.0 Report"))
    run.process()
