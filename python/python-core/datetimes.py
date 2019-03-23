# import our libraries
import time
import datetime

# get today's date
today = date.today()
print(today)

# create a custom date
future_date = date(2020, 1, 31)
print(future_date)

# let's create a time stamp
time_stamp = time.time()
print(time_stamp)

# create a date from a timestamp
date_stamp = date.fromtimestamp(time_stamp)
print(date_stamp)

# get components of a date
print(date_stamp.year)
print(date_stamp.month)
print(date_stamp.day)
