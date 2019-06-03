
# !/usr/bin/env python3
# coding: utf-8
from datetime import *
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


def calcAvg(envList):
    totalTime= timedelta()
    counterSucc = 0
    counterFailed = 0

    for day in envList:
        if("SUCCESS" in day[2]):
            counterSucc+=1
            (h,s,m) = day[4].split(':')
            totalTime += timedelta(hours=int(h), minutes=int(m), seconds=int(s))
        else:
            counterFailed+=1
    avg = 0
    if counterSucc>0:
        print(totalTime)
        avg =  totalTime.total_seconds()/(counterSucc)
        print(timedelta(seconds=avg))
    return (timedelta(seconds=avg), counterSucc, counterFailed)

def printToFile(outputLog,envName,envLogs,month):

    sheet = outputLog
    sheet = outputLog.add_sheet(envName+" analysis")
    row = 0
    col = 0

    sheet.write(row, col, "OBI Host Name: ", titleStyle)
    sheet.write(row + 1, col, envLogs[envName][0][0], style)
    sheet.write(row, col + 1, "DWH Host Name: ", titleStyle)
    sheet.write(row + 1, col + 1, envLogs[envName][0][1], style)

    #returns tuple - 0: avg time, 1: num of successed etls , 2: num of failed etls
    calc = calcAvg(envLogs[envName])
    sheet.write(row, col + 2, "ETL avg. time: (passed)", titleStyle)
    sheet.write(row + 1, col + 2,str(calc[0]) , style)

    sheet.write(row, col + 3, "Number of passed ETLs", titleStyle)
    sheet.write(row + 1, col + 3, calc[1], style)

    sheet.write(row, col + 4, "Number of failed ETLs", titleStyle)
    sheet.write(row + 1, col + 4, calc[2], style)


    #settings cols size
    for i in range(6):
        sheet.col(i).width = (15) * 367



