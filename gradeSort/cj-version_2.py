# !/usr/bin/python
# _*_coding: utf-8 _*_

import json
import sys
from openpyxl import Workbook

reload(sys)
sys.setdefaultencoding('utf-8')

def convertToTitle(n):
    """
    :type n: int
    :rtype: str
    #需要注意26时：26%26为0 也就是0为A 所以使用n-1  A的ASCII码为65
    """
    result = ""
    while n != 0:
        result = chr((n - 1) % 26 + 65) + result
        n = (n - 1) / 26
    return result

def main():
    #读取json文件
    with open('CJ.json','r') as f:
        CJ_list=json.load(f)
    #导出到excel表
    #创建Excel
    wb=Workbook()
    #创建表
    sheet=wb.active
    sheet.title='CJ'

    sheet['B1'].value='学分'
    sheet['A2'].value='学号'
    sheet['B2'].value='姓名'
    row_i=3
    kccol_i=3
    kc_list=list()
    for s_dict in CJ_list:
        sheet['A'+str(row_i)].value=s_dict[u'学号']
        sheet['B'+str(row_i)].value=s_dict[u'姓名']
        kc_mc=''
        for kc in s_dict['grade_list']:
            #过滤掉*和#，去掉两头空格
            if '#' in kc[0]:
                kc_mc=kc[0].replace('#','').strip()
            elif '*' in kc[0]:
                kc_mc=kc[0].replace('*','').strip()
            else:
                kc_mc=kc[0]
            print kc_mc
            if kc_mc not in kc_list:
                sheet[convertToTitle(kccol_i) + str(2)].value = kc_mc#录入课程名称
                sheet[convertToTitle(kccol_i) + str(1)].value = kc[1]#录入学分
                kccol_i = kccol_i + 1
                kc_list.append(kc_mc)
            col_i=kc_list.index(kc_mc)+3
            #读取单元格数字，若为空白则为0
            if sheet[convertToTitle(col_i) + str(row_i)].value == None or kc[2] > float(sheet[convertToTitle(col_i) + str(row_i)].value) :
                sheet[convertToTitle(col_i) + str(row_i)].value=str(kc[2])
        row_i=row_i+1


    wb.save('cj.xlsx')

main()
print convertToTitle(37)