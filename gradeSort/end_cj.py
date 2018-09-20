# !/usr/bin/python
# _*_coding: utf-8 _*_
import json
import sys
from openpyxl import load_workbook

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

def last():
    wb=load_workbook('cj.xlsx')
    sheet=wb.active
    sheet_z=wb.create_sheet('sheet_z')
    i=3
    while(sheet[convertToTitle(i)+str(2)].value != None):
        i=i+1
    col_end=i-1#从3到end都有数据
    i=3
    while (sheet['B' + str(i)].value != None):
        i = i + 1
    row_end=i-1
    for j in range(3,row_end+1):#每一个同学的
        sum_xf=0
        wtgkc_list=list()
        for i in range(3,col_end+1):#每一个成绩print sheet[convertToTitle(i)+str(j)].value
            if sheet[convertToTitle(i)+str(j)].value != None and float(sheet[convertToTitle(i)+str(j)].value) >=60:
                sum_xf = sum_xf + float(sheet[convertToTitle(i)+str(1)].value)
            elif sheet[convertToTitle(i)+str(j)].value != None and float(sheet[convertToTitle(i)+str(j)].value) <=60:
                print i
        print sheet['B'+str(j)].value,str(sum_xf)
        sheet_z['A'+str(j-2)].value=sheet['A'+str(j)].value
        sheet_z['C'+str(j-2)].value=str(sum_xf)
    wb.save('cj.xlsx')
last()