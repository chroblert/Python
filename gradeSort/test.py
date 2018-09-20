#!/usr/bin/python
# _*_coding: utf-8 _*_

import re
import chardet
import json
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

def htmlToJson():
    with open('CJ_sort1.html','r') as f:
        line1=f.read()
    type=chardet.detect(line1)
    data=line1.decode('GB2312',errors='ignore')
    #matchObj = re.finditer(r'</style>.*?</div>(.*?)<br />.*?/>(.*?)<br />.*?class="tableStyleTd">(.*?)</td>.*?StyleTd">(.*?)</td>.*?StyleTd">(.*?)</td>.*?<style>', data, re.M | re.I|re.S)
    #matchObj = re.finditer(r'class="tableStyleTd">(.*?)</td>.*?StyleTd">(.*?)</td>.*?StyleTd">(.*?)</td>', data, re.M | re.I|re.S)

    matchObj = re.finditer(r'</style>.*?95%;".*?</div>.*?:(.*?)<br />.*?/>.*?:(.*?)<br />(.*?)<style>', data, re.M | re.I|re.S)
    CJ_list=[]
    if matchObj:
        for m in matchObj:
            stu_dict = {}
            # print m.group(1),
            # print m.group(2),
            stu_dict.setdefault('学号',m.group(1).strip())
            stu_dict.setdefault('姓名', m.group(2).strip())
            cj_data=m.group(3)
            grade_list = []
            for i in re.findall(r'class="tableStyleTd">(.*?)</td>.*?StyleTd">(.*?)</td>.*?StyleTd">(.*?)</td>', cj_data, re.M | re.I|re.S):
                temp_list=[]
                for j in i:
                    #过滤
                    if '<strong>#</strong>' in j:
                        j=j.replace('<strong>#</strong>','#')
                    if '<strong>*</strong>' in j:
                        j=j.replace('<strong>*</strong>','*')
                    temp_list.append(j.strip())
                    # print j,
                # print
                grade_list.append(temp_list)
                stu_dict.setdefault('grade_list', grade_list)
            CJ_list.append(stu_dict)
            # print
    else:
        print "No match"

    #导出到json文件中
    file_name="CJ.json"
    with open(file_name,'w') as f:
        json.dump(CJ_list,f,ensure_ascii=False,indent=4)

htmlToJson()