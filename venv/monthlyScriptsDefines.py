
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

    def convert(secs):
        return str(timedelta(seconds=secs))


    totalTime= timedelta()
    time=0
    counterSucc = 0
    counterFailed = 0

    for day in envList:

        if("SUCCESS" in day[2]):
            counterSucc+=1
            d=0
            if "day" in day[4]:
                d=int(day[4].split()[0])
                day[4] = day[4].split(',')[1][1:]


            (h,m,s) = day[4].split(':')
            h = str(int(h) + d*24)
            s=(s.split('.')[0])
            time += int(s)+ int(m)*60 + int(h)*60*60
            totalTime += timedelta(hours=int(h), minutes=int(m), seconds=int(s))

        elif "Failed" in day[2]:
            counterFailed+=1
    avg = 0
    if counterSucc>0:
        #print(totalTime)
        avg = totalTime.total_seconds()/(counterSucc)

    return (convert(avg), counterSucc, counterFailed)

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


    row+=3

    sheet.write(row , col, "Starting time: ", titleStyle)
    sheet.write(row , col + 1, "Status: ", titleStyle)
    sheet.write(row , col + 2, "Total time: ", titleStyle)
    sheet.write(row , col + 3, "OBI\DWH ver: ", titleStyle)

    for i in range(len(envLogs[envName])-1):
        sheet.write(row + i + 1, col, envLogs[envName][i][3], style)
        sheet.write(row + i + 1, col + 1, envLogs[envName][i][2], style)
        sheet.write(row + i + 1, col + 2, envLogs[envName][i][4], style)
        sheet.write(row + i + 1, col + 3, envLogs[envName][i][5], style)


    #settings cols size
    for i in range(6):
        sheet.col(i).width = (15) * 367



