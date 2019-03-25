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

# ------------------------- PART TWO --------------------------

from datetime import datetime, date, time

# create a date and a time
my_date = date(2019, 3, 22)
my_time = time(12, 30)

# create a datetime
my_datetime = datetime.combine(my_date, my_time)
print(my_datetime)

# get the different components
print(my_datetime.year)
print(my_datetime.month)
print(my_datetime.day)
print(my_datetime.hour)
print(my_datetime.minute)
