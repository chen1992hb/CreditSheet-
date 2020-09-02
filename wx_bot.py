from wxpy import *
import datetime

today = datetime.datetime.now()
first_day = datetime.datetime(today.year, today.month,1)
print(str(first_day))