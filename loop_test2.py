# coding=utf-8

import xlwings as xw
import pandas as pd


def main(excel_name='敏感性分析1.1(1).xlsx'):
    # initial an error log, if the counted NPV is assumed to be wrong, the error info. will be recorded in error.xlsx
    # error_log = [['age', 'gender', 'time_flag', 'sa_level', 'the "age" where the cash_flow is blank']]
    # initial the `circular table`, which is used to record the calculated NPV
    circular_table = []

    # load data.xlsx
    wb = xw.Book(excel_name)
    # load sheet
    sht = wb.sheets['基本信息']

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
                    sht.range('B2').value = age
                    # set gender
                    sht.range('B3').value = gender
                    # set time flag
                    sht.range('B4').value = flag
                    # set sa level
                    sht.range('B5').value = level
                    # read NPV
                    NPV = sht.range('E3').value
                    # print(sht.range('B2').value,
                    #       sht.range('B3').value,
                    #       sht.range('B4').value,
                    #       sht.range('B5').value,
                    #       sht.range('B6').value,
                    #       NPV)
                    circular_table_row.append(NPV)
        circular_table.append(circular_table_row)
    # convert to DataFrame
    circular_table = pd.DataFrame(circular_table)
    # print(circular_table)
    # save to excel
    circular_table.to_excel('circular_table.xlsx', index=False, header=False)


if __name__ == '__main__':
    main()
