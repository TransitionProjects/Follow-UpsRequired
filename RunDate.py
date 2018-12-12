__author__ == "David Marienburg"
__version__ == ".1"

from datetime import date
from datetime import datetime
from calendar import monthrange
from dateutil.relativedelta import relativedelta

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
