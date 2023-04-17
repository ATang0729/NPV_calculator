"""use win32com to calculate NPV"""

import win32com
from win32com.client import Dispatch, constants
import os
import pandas as pd


# 获取当前脚本路径
def getScriptPath():
    nowpath = os.path.split(os.path.realpath(__file__))[0]
    print(nowpath)
    return nowpath


def main(excel_name='敏感性分析1.1(1).xlsx'):
    circular_table = []

    app = win32com.client.Dispatch('Excel.Application')

    # 后台运行，不显示，不警告
    app.Visible = 0
    app.DisplayAlerts = 0

    WorkBook = app.Workbooks.Open(os.path.join(getScriptPath(), excel_name))
    WorkSheet = WorkBook.Worksheets('基本信息')

    # process in batch
    age_list = range(20, 66)  # age range from 20 to 65
    gen_list = [0, 1]  # gender list
    time_flags = [0, 1]
    sa_levels = [1, 2, 3]
    for age in age_list:
        print('age: {}'.format(age))
        circular_table_row = []
        for gender in gen_list:
            for flag in time_flags:
                for level in sa_levels:
                    # set age
                    WorkSheet.Range('B2').Value = age
                    # set gender
                    WorkSheet.Range('B3').Value = gender
                    # set time flag
                    WorkSheet.Range('B4').Value = flag
                    # set sa level
                    WorkSheet.Range('B5').Value = level
                    # read NPV
                    NPV = WorkSheet.Range('E3').Value
                    # print(WorkSheet.Range('B2').Value,
                    #       WorkSheet.Range('B3').Value,
                    #       WorkSheet.Range('B4').Value,
                    #       WorkSheet.Range('B5').Value,
                    #       WorkSheet.Range('B6').Value,
                    #       NPV)
                    circular_table_row.append(NPV)
        circular_table.append(circular_table_row)
    # convert to DataFrame
    circular_table = pd.DataFrame(circular_table)
    circular_table.to_excel('result.xlsx', index=False, header=False)

    WorkBook.Close()
    app.Quit()


if __name__ == '__main__':
    main()
