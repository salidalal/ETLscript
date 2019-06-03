
# !/usr/bin/env python3
# coding: utf-8
from datetime import datetime
import xlrd
import xlwt
style = xlwt.easyxf('align: horiz center')
succ = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                   'font: colour white, bold True;'
                   'align: horiz center')
fail = xlwt.easyxf('pattern: pattern solid, fore_colour red;'
                   'font: colour white, bold True;'
                   'align: horiz center')
titleStyle = xlwt.easyxf("font: bold True; align: horiz center")
format = '%Y-%m-%d %H:%M:%S.%f'



def addToList(file,envList):
    xlFile = xlrd.open_workbook(file)

    for env in xlFile.sheets():
        row = env.row_values(1,0,14)
        envList[row[1]]+=[row]

def printToFile(outputLog,envName,envLogs):

    sheet = outputLog
    sheet = outputLog.add_sheet(month+" analysis")
    row = 0
    col = 0

    sheet.write(row, col, "OBI Host Name: ", titleStyle)
    sheet.write(row + 1, col, self.getOBIhost, style)
    sheet.write(row, col + 1, "DWH Host Name: ", titleStyle)
    sheet.write(row + 1, col + 1, self.getDWHhost, style)
    sheet.write(row, col + 2, "Last ETL status: ", titleStyle)






