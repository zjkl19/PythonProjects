# -*- coding: utf-8 -*-
"""
Created on Fri Nov  3 15:33:45 2017

说明:
'$'表示测站
'#'表示结束
excel中A列表示编号
测站：B列表示后视点编号,C列表示后视点读数值
测点: B列表示读数值

@author: ldn
"""

import os

import xlrd
import xlwt

rfileName='LevelingSurvey.xls'
wfileName='LevelingSurveyResult.xls'
sheetName='Sheet1'

data=xlrd.open_workbook(rfileName) #打开Exce1文件
sh=data.sheet_by_name(sheetName) #获得需要的表单


dic={}    #空字典,字典存储真实高程

pKey=sh.cell_value(0,0)   #点号
pVal=[sh.cell_value(0,1),sh.cell_value(0,2)]    #第2个值和第3个值
dic[pKey]=pVal[1]

pointer=1    #行指针
pKey=sh.cell_value(pointer,0)   #点号
pVal=[sh.cell_value(pointer,1),sh.cell_value(pointer,2)]    #第2个值和第3个值


while pKey!='#':    #判断是否到最后一个点
    if pKey[0]=='$':    #后视点
        aft=pVal[1]    #后视点读数
        h=dic[pVal[0]] #后视点高程
    else:
        dic[pKey]=h+aft-pVal[0]
    
    pointer=pointer+1
    pKey=sh.cell_value(pointer,0)   #点号
    pVal=[sh.cell_value(pointer,1),sh.cell_value(pointer,2)]    #第2个值和第3个值


path=os.sys.path[0]
w=xlwt.Workbook()
ws=w.add_sheet(sheetName)
i=0
for key, value in dic.items():
    ws.write(i,5,key)
    ws.write(i,6,value)
    i=i+1
    
localPath=os.path.join(path,wfileName)
w.save(localPath)
