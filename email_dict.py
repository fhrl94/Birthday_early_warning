import pprint

import xlrd


# TODO 人员范围未处理
def to_send_email():
    send_email = {}
    workbook = xlrd.open_workbook(filename='主管及以上名单.xlsx')
    for sheet, one in enumerate(workbook.sheet_names()):
        for i in range(1, workbook.sheet_by_name(one).nrows):
            if workbook.sheet_by_name(one).cell_value(i, 2) != "":
                # print(workbook.sheet_by_name(one).cell_value(i, 2))
                # print(workbook.sheet_by_name(one).cell_value(i, 6))
                send_email[workbook.sheet_by_name(one).cell_value(i, 2)] = workbook.sheet_by_name(one).cell_value(i, 6)
    # print(send_email)
    # pprint.pprint(send_email)
    return send_email

# if __name__ == '__main__':
#     to_send_email()

