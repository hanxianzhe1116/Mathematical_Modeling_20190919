##!/usr/bin/env python3

import time
import pandas as pd
import openpyxl

'''
    将unix时间转换为日期格式
    param：unix_time(1537860800.0)
    return: time(2018-09-25 15:33:20)

'''
def unix_time_to_time(unix_time):
    time_array = time.localtime(unix_time)

    time_style = time.strftime("%Y-%m-%d %H:%M:%S",time_array)

    return time_style

'''
    将正常的时间转为unix time
    param: 2018-09-25 15:33:20
    return: 1537860800.0
'''
def unix_time(dt):
    time_array = time.strptime(dt,"%Y-%m-%d %H:%M:%S")
    time_stamp = time.mktime(time_array)
    return time_stamp

'''
    读取excel数据
    param: excel_data_url(excel文件路径)
    return: 增删改查等操作
'''
def open_excel(excel_data_url):
    excel_sheet = openpyxl.load_workbook(excel_data_url)
    # print(excel_sheet.sheetnames)

    excel_data = excel_sheet['Sheet1']

    for row in excel_data.rows:
        for cell in row:
            print(cell.value,end=',')
        print()

    # excel_active = excel_data.active
    #
    # col_start_time = excel_active['起飞时间']
    # print(col_start_time)
    # for start_time_item in col_start_time:
    #     print(start_time_item)


def main():
    excel_data_url = './data/Schedules.xlsx'
    open_excel(excel_data_url)


    # time_now = '2016-04-23 02:02:00'
    # unix_t = unix_time(time_now)
    # print(unix_t)

    # unix_t = 1461399900
    # time_style = unix_time_to_time(unix_t)
    # print(time_style)



if __name__ == '__main__':
    main()
