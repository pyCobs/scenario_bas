import datetime
from datetime import date


def addYears(d, years):
    try:
#Return same day of the current year
        return d.replace(year = d.year + years)
    except ValueError:
#If not same day, it will return other, i.e.  February 29 to March 1 etc.
        return d + (date(d.year + years, 1, 1) - date(d.year, 1, 1))
