"""
    Python offers the capability for manipulating, time, date, and datetime objects. All the manipulation you'll need to do is through the `datetime` library.

    In this tutorial, we will explore the `date` object, the `time` object and the `datetime` object. Additionally, we will see how to solve everyday
    problems that require using some type of duration with the `datetime` library. 
"""

# first we need to import all of our objects, the best way to import everything is using the form alias.
from datetime import date, time, datetime, timezone, timedelta

# let's start with a date. If we want to create a new date object, we can use the `date` object and pass through some arguments
my_date = date(year = 2020, month = 3, day = 1)
print("Here is my new date: {date}".format(date = my_date))


# with a date, I can grab it's different `parts`
print("\nThe YEAR for my new date is {my_year}".format(my_year=my_date.year))
print("The MONTH for my new date is {my_month}".format(my_month=my_date.month))
print("The DAY for my new date is {my_day}".format(my_day=my_date.day))


# I can also replace different parts, for example just the year
my_new_date = my_date.replace(year = 2021)
print("\nHere is my new date, with just the year replaced: {my_new_year}".format(my_new_year = my_new_date))

# or I can also replace all the parts.
my_new_date = my_date.replace(year = 2021, month = 5, day = 5)
print("Here is my new date, with all the parts replaced: {my_new_date}".format(my_new_date = my_new_date))


'''
    I can also grab some additional context information about my date.
'''

# for example, let's see the week day, remember it's 0 (Monday) to 6 (Sunday).
print("\nThe WEEKDAY for my new date is: {my_weekday}".format(my_weekday=my_date.weekday()))

# let's see the ISO week day, remember it's 1 (Monday) to 7 (Sunday).
print("The ISO WEEKDAY for my new date is: {my_isoweekday}".format(my_isoweekday=my_date.isoweekday()))

# let's see the Proleptic Gregorian date
print("The PROLETPIC GREGORIAN DATE for my new date is: {my_ordinal}".format(my_ordinal=my_date.toordinal()))

# let's see the string representation of my date
print("The STRING REPRESENTATION for my new date is: {my_string_date}".format(my_string_date=my_date.ctime()))

# let's see the ISO FORMAT representation of my date
print("The ISO FORMAT REPRESENTATION for my new date is: {my_string_date}".format(my_string_date=my_date.isoformat()))

# let's see the ISO FORMAT representation of my date
print("The ISO CALENDAR for my new date is: {my_string_date}".format(my_string_date=str(my_date.isocalendar())))

# let's see the ISO FORMAT representation of my date
print("The TIME TUPLE for my new date is: {my_string_date}".format(my_string_date=str(my_date.timetuple())))


'''
    Sometimes we want today's date, that's easy to get with the `today()` method.
'''
today_is = date.today()
print("\nHere is today's date: {today_is}".format(today_is = today_is))

# if you want the earliest possible date you can create, then use the `min` attribute.
print("Here is the earliest possible date I can create using the Python `date` object: {min_date}".format(min_date = date.min))

# if you want the latest possible date you can create, then use the `max` attribute.
print("Here is the latest possible date I can create using the Python `date` object: {max_date}".format(max_date = date.max))