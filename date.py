from datetime import datetime as dt


def now_day():
    day = dt.now().strftime('%d')
    return day

def now_month():
    month = dt.now().strftime('%m')
    return month

def now_year():
    year = dt.now().strftime('%Y')
    return year