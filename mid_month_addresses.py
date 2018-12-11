"""
This script is for processing the Housing Services - Housing Outcome v2.0
report that is used by the follow-ups specialist(s).


This script should output a list of participants who were not contacted by
the follow-up specialist and will provide these persons' addresses.
"""

import pandas as pd
from datetime import datetime as dt
from tkinter.filedialog import askopenfilename as aofn
from tkinter.filedialog import asksaveasfilename as asafn

class CreateAddressList:
    def __init__(self, file_path):
        self.raw_data = pd.read_excel(file_path)
        self.current_month = dt.now().month
        self.current_year = dt.now().year

    def process(self):
        data = self.raw_data[
            ~(self.raw_data["Follow-Up Status(2729)"] == "Client contacted") &
            ~(self.raw_data["Follow-Up Status(2729)"] == "Other verifiable source contacted") &
            (self.raw_data["Follow Up Due Date(2512)"].dt.month == self.current_month) &
            (self.raw_data["Follow Up Due Date(2512)"].dt.year == self.current_year)
        ].sort_values(
            by=["Client Uid", "Follow Up Due Date(2512)"],
            ascending=False
        ).drop_duplicates(
            subset="Client Uid",
            keep="first"
        )
        writer = pd.ExcelWriter(
            asafn(
                title="Save the Non-Contacted Follow-Ups Report",
                initialfile="Non-Cotacted Follow-Ups.xlsx",
                defaultextension=".xlsx"
            ),
            engine="xlsxwriter"
        )
        data.to_excel(writer, sheet_name="Data", index=False)
        writer.save()

if __name__ == "__main__":
    run = CreateAddressList(
        aofn(title="Open the Housing Outcomes v2.0 (Mid Month) Report")
    )
    run.process()
