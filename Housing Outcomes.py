__author__ = "David Marienburg"
__version__ = "1.12"

"""
This script is for processing the Housing Services - Housing Outcomes v2.0 report that is used by
the follow-ups specialist.

This script should identify when it is being run and adjust its date parameters to reflect being run
on the first of the month or after it.  This well prevent future staff members from needing to
fiddle with the code every time they run the report.
"""

# import required libraries
import pandas as pd
from datetime import datetime
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class CreateRequiredFollowUps:
    def __init__(self, file_path):
        # read the excel report into a pandas data frames
        self.raw_fu_data = pd.read_excel(file_path, sheet_name="FollowUps")
        self.raw_placement_data = pd.read_excel(file_path, sheet_name="Placements")
        self.raw_address_data = pd.read_excel(file_path, sheet_name="Addresses")
        # create a immutable list of unique months during which follow-ups are
        # due
        self.month_range = set(
            [value.strftime("%B") for value in self.raw_fu_data["Follow Up Due Date(2512)"]]
        )
        # create month and year name variables for the name of the processed
        # report
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year

    def process(self):
        # create a local copy of the self.raw_data data frame then merge that
        # copy with the address and placement data frames to ensure that all
        # followups are related to a TPI placement and that the Addresses
        # provided are the newest addresses.
        data = self.raw_fu_data.merge(
            self.raw_address_data.sort_values(
                by=["Client Unique Id", "Date Added (61-date_added)"],
                ascending=False
            ).drop_duplicates(subset="Client Unique Id"),
            how="left",
            on="Client Unique Id"
        ).merge(
            self.raw_placement_data,
            how="inner",
            left_on=["Client Unique Id", "Initial Placement/Eviction Prevention Date(2515)"],
            right_on=["Client Unique Id", "Placement Date(3072)"]
        )
        # initiate the ExcelWriter object variable
        writer = pd.ExcelWriter(
            asksaveasfilename(
                title="Save the Required Follow-ups Report",
                defaultextension=".xlsx",
                initialfile="Required Follow-ups for {} {}".format(
                    self.current_month,
                    self.current_year
                )
            ),
            engine="xlsxwriter"
        )
        # loop through the values of the self.month_range set creating dataframes
        # where the value of Follow Up Due Date(2512) column is equal to the set
        # item's value creating excel sheets for each of these data drames
        for month in self.month_range:
            month_data = data[
                (data["Follow Up Due Date(2512)"].dt.strftime("%B") == month) &
                data["Actual Follow Up Date(2518)"].isna()
            ].drop_duplicates(
                subset="Client Unique Id"
            ).drop(
                ["Client Unique Id", "Client Uid_y", "Placement Date(3072)"],
                axis=1
            )
            month_data.to_excel(
                writer,
                sheet_name="{} Follow-Ups".format(month),
                index=False
            )
        # create an excel sheet containing the raw data
        data.to_excel(writer, sheet_name="Raw Data", index=False)
        # save the spreadsheet
        writer.save()

if __name__ == "__main__":
    run = CreateRequiredFollowUps(
        askopenfilename(title="Open the Housing Outcomes v2.2 Report")
    )
    run.process()
