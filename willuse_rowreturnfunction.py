


#  本函数的功能为：输入工作表，返回一个字典，字典的key为员工姓名，值为工作表中员工打卡时间所在的行数

def rowreturn(wsheet):
    name_row= {}
    for row in range(5, wsheet.max_row, 2):
        name_row[wsheet.cell(row, 11).value] = row

    return name_row
