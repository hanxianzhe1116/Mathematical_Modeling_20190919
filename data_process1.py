import pandas as pd
import datetime as dt


data = pd.read_excel("data1.xlsx", sheet_name='Sheet1')


def open_excel(excel_data_url):
    #excel_sheet = openpyxl.load_workbook(excel_data_url)
    global endrow
    data = pd.read_excel(excel_data_url, sheet_name='Sheet1')
    num = 0
    start = []
    endrow = 0
    for row in range(2, len(data)):
        #for col in range(1, 14):
            # if sh.cell(row=row, column=12).value != 0 and sh.cell(row=row, column=2).value == 0:
                # 获取某行某列单元格元素
            if data.ix(row, 14) !=0 and data.ix(row, 2) == 0:
                start.append(row)
                num = num + 1   # 共有多少个异常数据
                endrow = row    # 取最后一个异常的行数
            elif data.ix(row, 2) != 0 and num <= 180:
                # print('1111')
                num = 0
                start = []
            #print(sh.cell(row=row, column=col).value, end=',')
        #print()
    startrow = start[0]
    #     print(startrow)
    #     print(endrow)
    # print(num)

    # 删除GPS车速为0的数据（大于180条）
    count = 0
    for row in range(startrow, endrow+1):
        count = count + 1
        print(count)
        data.drop(row, axis=0)
        # sh.delete_rows(startrow)
    # wb.save(excel_data_url)


def main():
    excel_data_url = 'testdata.xlsx'
    open_excel(excel_data_url)


if __name__ == '__main__':
    main()