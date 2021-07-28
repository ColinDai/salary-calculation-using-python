import datetime

#  本函数的功能为计算两个打卡时间之间的上班时间
def time_sub(datetime_start, datetime_end):
    minute1 = datetime_start.time().minute
    minute2 = datetime_end.time().minute
    minute1_fl = minute1 // 15
    new_datetime1_hour = int(minute1_fl / 2) * 0.5 + int(minute1_fl % 2) * 0.5
    minute2_fl = minute2 // 15
    new_datetime2_hour = int(minute2_fl / 2) * 0.5 + int(minute2_fl % 2) * 0.5
    datetime_start_min00 = datetime.datetime(datetime_start.year, datetime_start.month, datetime_start.day,
                                             datetime_start.hour, 0, 0)
    datetime_end_min00 = datetime.datetime(datetime_end.year, datetime_end.month, datetime_end.day,
                                           datetime_end.hour, 0, 0)
    hours1 = datetime_start_min00.__rsub__(datetime_end_min00).total_seconds() / 60 / 60
    if datetime_start.time()<datetime.time(12,0,0) and datetime_end.time()>=datetime.time(13,0,0):
        hours = hours1 + new_datetime2_hour - new_datetime1_hour-1
        return hours
    elif datetime_end.time()<datetime.time(13,0,0)  or datetime_start.time()>datetime.time(12,0,0) or \
            (datetime_start.time().hour==12 and datetime_end.time().hour==13):
        hours = hours1 + new_datetime2_hour - new_datetime1_hour
        return hours


