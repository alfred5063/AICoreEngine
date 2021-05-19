#!/usr/bin/env python
#
# Created by Xixuan on 29/08/2018
#
# The file contains useful tools (functions).
#

from datetime import datetime
from pandas import isnull, Timestamp
import re


def ct(plain=False):
    if plain:
        return datetime.utcnow().strftime('%Y_%m_%d_%H_%M_%S_%f')[:-3]
    else:
        return datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]

def empty_column_check(table):
    return False # TODO: Implement this check properly
    '''
    index = [i for i in range(len(table[0])) if table[0] == ""]
    if len(index) > 0:
        for col in index:
            if all([row[col] == "" for row in table]):
                for row in table:
                    del row[col]
        return True
    return False'''

def is_date(date):
    try:
        day, month, year = date.split('/')
        datetime(int(year), int(month), int(day))
    except ValueError:
        return False
    return True

def date_processing(date):
    if isnull(date):
        return (None, False)
    elif type(date) == datetime:
        return (date, False)
    elif date == "":
        return (None, False)
    elif type(date) == Timestamp:
        return (datetime.combine(date.date(), datetime.min.time()), False)
    elif is_date(date):
        return (datetime.strptime(date, "%d/%m/%Y"), False)
    elif re.sub("[0-9]+[/|-][0-9]+[/|-][0-9]+","",date) == "":
        date = datetime.strptime(re.sub("[/|-]","/",date), "%d/%m/%Y")
        return (date, isnull(date))
    elif re.sub("[0-9]+-[0-9]+-[0-9]+ [0-9]+:[0-9]+:[0-9]+","",date)=="":
        date = datetime.strptime(re.findall("[0-9]+-[0-9]+-[0-9]+", date), "%d-%m-%Y")
        return (date, isnull(date))
    elif re.sub("[0-9]*","", date) == "":
        date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date) - 2)
        return (date, isnull(date))
    else:
        return (None, True)


def date_formatting(table, date_formats):
    errorlists = []
    for format in date_formats:
        if format.casefold() in map(str.casefold, table[0].keys()):
            temp = [date_processing(row[format]) for row in table]
            errorlists.append([item[1] == True for item in temp])
            for row, date in zip (table, temp):
                row[format] = date[0]
    errors = [any(tuple) for tuple in zip(*errorlists)]
    if (any(errors)):
        return errors
    return None



def empty_remark(str):
    if re.sub("^[ ]*$", "", str) == "":
        return ""
    else:
        return str