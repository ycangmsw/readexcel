#!/usr/bin/env python
#-*- coding:utf-8 -*-

import sys
import conn_sql
import xlrd
from datetime import datetime

reload(sys)
sys.setdefaultencoding('utf8')
myfile=u'售后记录2017.xlsx'

#定义列名常量
DATA=1       #日期
NUMBER=2     #任务单号
CLIENT=4     #客户单位名称
CLIADDR=5    #客户地址
CLINAME=6    #联系人名
PHONE=7      #联系电话
NAME=8       #销售员名
CRATYPE=9    #起重机类型/型号
METTYPE=10      #仪表类型/型号
TYPE=11      #售后类型
FAUDESCRIBE=12  #故障描述
FAUANALYZE=13   #故障分析
FAURESULT=14    #处理结果

db=conn_sql.mydb()
db.conn_db()
workbook=xlrd.open_workbook(myfile)
sheet=workbook.sheet_by_index(4)
for row in range(sheet.nrows):
    if row != 0 :
        as_date = xlrd.xldate.xldate_as_datetime(sheet.row(row)[DATA].value,0)
        as_number = str(sheet.row(row)[NUMBER].value)
        if '.' in as_number:
            as_number = as_number[:-2]

        as_type = sheet.row(row)[TYPE].value
        as_flagtype = "结束"
        as_name = sheet.row(row)[NAME].value
        as_client = sheet.row(row)[CLIENT].value
        as_clientname = sheet.row(row)[CLINAME].value
        as_clientphone = str(sheet.row(row)[PHONE].value)
        if '.' in as_clientphone:
            as_clientphone = as_clientphone[:-2]

        as_clientaddress = sheet.row(row)[CLIADDR].value
        as_cranetype = sheet.row(row)[CRATYPE].value
        as_metertype = sheet.row(row)[METTYPE].value
        as_cranemodel = "不详"
        as_metermodel = "不详"
        as_faultdescribe = sheet.row(row)[FAUDESCRIBE].value
        as_faultanalyze = sheet.row(row)[FAUANALYZE].value
        as_faultresult = sheet.row(row)[FAURESULT].value
        tup=(as_date,as_number,as_type,as_flagtype,as_name,as_client,as_clientname,as_clientphone,as_clientaddress,as_cranetype,as_metertype,as_cranemodel,as_metermodel,as_faultdescribe,as_faultanalyze,as_faultresult) 
        print u'行数',':',row,':',tup
        print
        cur=db.sql_insert(tup)
        if not cur:
            sys.exit(0)
        else:
            db.sql_close(cur)
db.sql_finish()

print u'结束'
