#!/usr/bin/python
# _*_coding: utf-8 _*_

import re
import chardet
import json
import sys
from openpyxl import Workbook

reload(sys)
sys.setdefaultencoding('utf-8')
line = '''<html>
 <head></head>
 <body>
  <table align="center" style="margin: 7px 1% -1px 2%; border: currentColor; border-image: none; width: 95%;" cellspacing="0" cellpadding="0" valign="top"> 
   <tbody> 
    <tr> 
     <td style="width: 100%; text-align: center;" colspan="10"><strong>安徽理工大学学业成绩表</strong></td> 
    </tr> 
    <tr> 
     <td style="width: 100%; text-align: center;" colspan="10"> 学院:计算机科学与工程学院&nbsp;&nbsp; 专业:信息安全&nbsp;&nbsp; 班级:信息安全15-2&nbsp;&nbsp; 入学日期:2012-09-07&nbsp;&nbsp; 毕业日期:2019-06-30&nbsp;&nbsp; 身份证号:310105199407181618 </td> 
    </tr> 
    <tr> 
     <td class="tableStyleTd2" valign="top" style="width: 10%; text-align: left; font-size: 10px;"> 
      <div style="text-align: center;">
       <img width="70" height="80" title="尤海天" src="/eams/showAvatar.action?user.name=2012303163" /> 
      </div> 学号 : 2012303163<br /> &nbsp;<br /> 姓名 : 尤海天<br /> &nbsp;<br /> 所在年级 : 2015<br /> &nbsp;<br /> 学制 : 4<br /> &nbsp;<br /> 层次 : 本科 <br /> 
      <hr style="margin: 0px auto; padding: 0px; height: 1px; color: rgb(0, 0, 0); overflow: hidden;" /> 
      <div style="text-align: left;">
       修读学分统计
      </div> &nbsp;&nbsp;<br /> 必修课 : 3<br /> &nbsp;&nbsp;<br /> 公共必修课程 : 87<br /> &nbsp;&nbsp;<br /> 学科专业必修 : 24.5<br /> &nbsp;&nbsp;<br /> 专业核心课程 : 20.5<br /> &nbsp;&nbsp;<br /> 实践教学环节 : 6<br /> &nbsp;&nbsp;<br /> 公共选修课程 : 12<br /> &nbsp;&nbsp;<br /> 跨学科选修 : 2<br /> &nbsp;&nbsp;<br /> 专业任选课程 : 22<br /> 
      <hr style="margin: 0px auto; padding: 0px; height: 1px; color: rgb(0, 0, 0); overflow: hidden;" /> 
      <div style="text-align: center; margin-top: 62px;">
       学院公章
       <br />2018-09-19
      </div> 
      <hr style="margin: 0px auto; padding: 0px; width: 110px; height: 1px; color: rgb(0, 0, 0); overflow: hidden;" /> 
      <div style="text-align: center; margin-top: 100px;">
       教务处公章
       <br />2018-09-19
      </div> </td> 
     <td class="tableStyleTd4" valign="top" style="width: 90%; border-left-color: currentColor; border-left-width: 0px; border-left-style: none;"> 
      <table align="center" class="tableStyle" style="width: 100%; border-left-color: currentColor; border-left-width: 0px; border-left-style: none; border-collapse: collapse;"> 
       <tbody>
        <tr> 
         <td width="30%"> 
          <table class="tableStyle" id="mainTable2" style="width: 100%;"> 
           <tbody>
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd" style="line-height: 101%;">课程名称</td> 
             <td width="12%" align="center" class="tableStyleTd" style="line-height: 101%;">课程学分</td> 
             <td width="12%" class="tableStyleTd" style="line-height: 101%;">课程成绩</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="25%" class="tableStyleTd" colspan="3">2012-2013学年第1学期</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 手球 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">80</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> C语言程序设计Ⅰ </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">47</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 大学英语（一） </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 形势与政策（一） </td> 
             <td width="12%" class="tableStyleTd">0</td> 
             <td width="12%" class="tableStyleTd">80</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 中国近现代史纲要 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">92</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 思想道德修养与法律基础 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">74</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 计算机科学导论 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">76</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> 高等数学Ⅰ（上） </td> 
             <td width="12%" class="tableStyleTd">5.5</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="25%" class="tableStyleTd" colspan="3">2012-2013学年第2学期</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 大学生心理健康教育 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">65</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 信息安全导论 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">64</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 形势与政策（二） </td> 
             <td width="12%" class="tableStyleTd">0</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>*</strong> C语言程序设计2 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>*</strong> 面向对象程序设计 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">50</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 程序设计训练2 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">87</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>*</strong> 离散数学 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 大学物理实验(上) </td> 
             <td width="12%" class="tableStyleTd">1.5</td> 
             <td width="12%" class="tableStyleTd">64</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 中国文化导论 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">75</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 足球 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">75</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 马克思主义基本原理 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">70</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> 高等数学Ⅰ（下） </td> 
             <td width="12%" class="tableStyleTd">6</td> 
             <td width="12%" class="tableStyleTd">63</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> 大学物理（上） </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">72</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> 大学英语（二） </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">80</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="25%" class="tableStyleTd" colspan="3">2016-2017学年第1学期</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 线性代数 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">68</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 大学物理实验Ⅰ（下） </td> 
             <td width="12%" class="tableStyleTd">1.5</td> 
             <td width="12%" class="tableStyleTd">74</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 数据结构课程设计 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">79</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 形势与政策（三） </td> 
             <td width="12%" class="tableStyleTd">0</td> 
             <td width="12%" class="tableStyleTd">75</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 大学英语（三） </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 毛泽东思想中国特色社会主义概论-上 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">90</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 数字逻辑 </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">63</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 模拟电子技术 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">71</td> 
            </tr> 
           </tbody>
          </table> </td> 
         <td width="30%"> 
          <table class="tableStyle" id="mainTable2" style="width: 100%;"> 
           <tbody>
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd" style="line-height: 101%;">课程名称</td> 
             <td width="12%" align="center" class="tableStyleTd" style="line-height: 101%;">课程学分</td> 
             <td width="12%" class="tableStyleTd" style="line-height: 101%;">课程成绩</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 数据结构 </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">75</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 认识实习 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">76</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>*</strong> 概率论与数理统计 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> 大学物理（下）Ⅰ </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">92</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="25%" class="tableStyleTd" colspan="3">2016-2017学年第2学期</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 追寻幸福：西方伦理史视角(尔雅) </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">95.08</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> Java程序设计(双语) </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 毛泽东思想中国特色社会主义概论-下 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">81</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 组成原理课程设计 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">78</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 数据库概论 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">68</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 计算机组成原理 </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">69</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 数据库系统原理实验及课程设计 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">83</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 信息安全的数学基础 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">84</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 汇编语言程序设计 </td> 
             <td width="12%" class="tableStyleTd">2.5</td> 
             <td width="12%" class="tableStyleTd">71</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 从“愚昧”到“科学”-科学技术简史(尔雅) </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">98.94</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> Jsp程序设计 </td> 
             <td width="12%" class="tableStyleTd">2.5</td> 
             <td width="12%" class="tableStyleTd">63</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 大学英语（四） </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">66</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 形势与政策（四） </td> 
             <td width="12%" class="tableStyleTd">0</td> 
             <td width="12%" class="tableStyleTd">81</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="25%" class="tableStyleTd" colspan="3">2017-2018学年第1学期</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> 编译原理 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">75</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> Android系统开发 </td> 
             <td width="12%" class="tableStyleTd">2.5</td> 
             <td width="12%" class="tableStyleTd">82</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 操作系统 </td> 
             <td width="12%" class="tableStyleTd">4</td> 
             <td width="12%" class="tableStyleTd">81</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 舞蹈鉴赏(尔雅) </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">94.47</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> C语言程序设计 </td> 
             <td width="12%" class="tableStyleTd">5</td> 
             <td width="12%" class="tableStyleTd">41</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 信息论与编码技术 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">77</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 形势与政策（五） </td> 
             <td width="12%" class="tableStyleTd">0</td> 
             <td width="12%" class="tableStyleTd">81</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 操作系统课程设计 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd">70</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 现代密码学 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd">63</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 排球 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 文学人类学概说(尔雅) </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">95.57</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 计算机网络 </td> 
             <td width="12%" class="tableStyleTd">3.5</td> 
             <td width="12%" class="tableStyleTd">72</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="25%" class="tableStyleTd" colspan="3">2017-2018学年第2学期</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 个人理财规划(尔雅) </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd">94.97</td> 
            </tr> 
           </tbody>
          </table> </td> 
         <td width="30%"> 
          <table class="tableStyle" id="mainTable2" style="width: 100%;"> 
           <tbody>
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd" style="line-height: 101%;">课程名称</td> 
             <td width="12%" align="center" class="tableStyleTd" style="line-height: 101%;">课程学分</td> 
             <td width="12%" class="tableStyleTd5" style="line-height: 101%;">课程成绩</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 网络安全与病毒防范 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd5">75.6</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 形势与政策（六） </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd5">82</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 中国哲学概论(尔雅) </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd5">96.3</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 信息隐藏技术 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd5">72.6</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> Linux管理系统及开发 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd5">68</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 武术（散打） </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd5">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 知识产权法 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd5">79</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 操作系统安全 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd5">61</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 软件工程 </td> 
             <td width="12%" class="tableStyleTd">2</td> 
             <td width="12%" class="tableStyleTd5">80</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 就业指导 </td> 
             <td width="12%" class="tableStyleTd">1</td> 
             <td width="12%" class="tableStyleTd5">77</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> 数据安全与恢复技术 </td> 
             <td width="12%" class="tableStyleTd">3</td> 
             <td width="12%" class="tableStyleTd5">60</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td width="76%" class="tableStyleTd"> <strong>#</strong> C语言程序设计 </td> 
             <td width="12%" class="tableStyleTd">3.5</td> 
             <td width="12%" class="tableStyleTd5">0.1</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
            <tr align="center" class="tableStyleTitle"> 
             <td class="tableStyleTd5" colspan="3">空白</td> 
            </tr> 
           </tbody>
          </table> </td> 
        </tr> 
       </tbody> 
      </table> 
      <table align="center" class="tableStyle" id="mainTable2" style="width: 100%; margin-right: 0px; margin-bottom: -1px; margin-left: -1px; border-left-color: currentColor; border-left-width: 0px; border-left-style: none; border-collapse: collapse;"> 
       <tbody>
        <tr align="center" class="tableStyleTitle"> 
         <td class="tableStyleTd1" style="border-top-color: currentColor; border-top-width: 0px; border-top-style: none;" colspan="6">毕业设计（论文）题目：</td> 
         <td class="tableStyleTd1" style="border-top-color: currentColor; border-top-width: 0px; border-top-style: none;" colspan="4">毕业设计（论文）成绩：</td> 
        </tr> 
        <tr align="center" class="tableStyleTitle"> 
         <td class="tableStyleTd1" colspan="10">培养计划学分要求：</td> 
        </tr> 
        <tr align="center" class="tableStyleTitle"> 
         <td class="tableStyleTd1" colspan="10">历年学分平均绩点：1.7 &nbsp;学分加权平均成绩：67&nbsp;&nbsp; 注： 课程名称前标注<strong>*</strong>的为补考成绩,标记<strong>#</strong>的为重学成绩。</td> 
        </tr> 
        <tr align="center" class="tableStyleTitle"> 
         <td class="tableStyleTd1" colspan="5">学籍异动情况： 恢复学籍 Sep 26, 2016 </td> 
         <td class="tableStyleTd1" colspan="2">毕（结）业证书号：</td> 
         <td class="tableStyleTd1" colspan="3">学位证书号：</td> 
        </tr> 
       </tbody> 
      </table> </td> 
    </tr> 
   </tbody> 
  </table>
 </body>
</html>'''

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

