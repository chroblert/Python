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
    i=3
    while(sheet[convertToTitle(i)+str(2)]!= None):
        i=i+1
    col_end=i-1#从3到end都有数据
    i=3
    while (sheet['B' + str(i)] != None):
        i = i + 1
    row_end=i-1
    for j in range(3,row_end+1):#每一个同学
        for i in range(3,col_end+1):
            if sheet[convertToTitle(i)+str(j)]

last()