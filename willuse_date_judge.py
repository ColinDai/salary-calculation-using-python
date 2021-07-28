import chinese_calendar as ri_li
import datetime



#  本函数的功能为判断当前日期的加班工资系数为：1.5,2或者3

def times_salary(date):
    triple_salary = [datetime.date(2021, 1, 1), datetime.date(2021, 2, 12), datetime.date(2021, 2, 13),
                     datetime.date(2021, 2, 14),
                     datetime.date(2021, 4, 4), datetime.date(2021, 5, 1), datetime.date(2021, 6, 14),
                     datetime.date(2021, 9, 21),
                     datetime.date(2021, 10, 1), datetime.date(2021, 10, 2), datetime.date(2021, 10, 3)]
    if ri_li.is_holiday(date):
        if date in triple_salary:
            return 3
        else:
            return 2
    else:
        return 1.5


