#!/usr/bin/env python

from datetime import datetime
import pandas as pd
import xlwt
import numpy as np
import functools
import ETLscrtipFuncs

#defines
months = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]
format = '%Y-%m-%d %H:%M:%S'
ErrorFormat = '%d/%m/%Y %H:%M:%S'
fix = "None"
etls = []
start = -1
fail = False





#reading files
exeLOG = pd.read_excel('Execution Log.xlsx', sheet_name='Sheet1')
log = open("ETL.log", 'r')
CheckApps = open("CheckApps.txt", "r").readlines()
setupDWH = open("SetupDWH.root.log", "r").readlines()
ErrorLog = pd.read_excel('ETL Error Log_ETL Error Log.xlsx', sheet_name="ETL Error Log")






#retrieve
dwh = CheckApps[3].split()[2][1:-1]
ver = CheckApps[9].split()[1]

for line in CheckApps:
    if "fix" in line.lower():
        fix = line.split()[1]
        break


for line in setupDWH:
    if "Host name:" in line:
        obi = line.split()[2]
        break


msg = exeLOG.values.T[3].tolist()[-2:1:-1]
status = exeLOG.values.T[2].tolist()[-2:1:-1]
task = exeLOG.values.T[1].tolist()[-2:1:-1]
timeStamp = exeLOG.values.T[0].tolist()[-2:1:-1]

time = ErrorLog.values.T[0].tolist()[8:]
cause = ErrorLog.values.T[4].tolist()[8:]


for i in range(len(msg)):
    if not fail:
        if task[i] == "Main Job":
            if "starting" in msg[i]:
                if start != -1:
                    etls += [ETL(start, i - 1)]
                    fail = False
                    start = -1

                else:
                    start = i

            elif "finished" in msg[i]:
                etls += [ETL(start, i)]
                fail = False
                start = -1

        if status[i] == "FAILURE":
            fail = True

    else:
        if task[i] == "Main Job":
            if "starting" in msg[i]:
                etls += [ETL(start, i - 1)]
                fail = False
                start = i
                end = -1

etls = etls[::-1]


for etl in etls:
    for i in range(len(time)):
        if etl.log[0][0] <= time[i] and etl.log[0][-1] >= time[i]:
            etl.err += [cause[i]]


curMonth = etls[0].log[0][0].month






def printToFile():
    style = xlwt.easyxf('align: horiz center')
    succ = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                       'font: colour white, bold True;'
                       'align: horiz center')
    fail = xlwt.easyxf('pattern: pattern solid, fore_colour red;'
                       'font: colour white, bold True;'
                       'align: horiz center')
    titleStyle = xlwt.easyxf("font: bold True; align: horiz center")

    row = 1
    col = 1

    def printDet(p):

        sheet1.write(p + 1, col, "Version: ", titleStyle)
        sheet1.write(p + 1, col + 1, ver, style)
        sheet1.write(p + 1, col + 2, "Fix: ", titleStyle)
        sheet1.write(p + 1, col + 3, fix, style)  # need to check this
        sheet1.write(p + 2, col, "Environment: ", titleStyle)
        sheet1.write(p + 2, col + 1, "obi :", titleStyle)
        sheet1.write(p + 2, col + 2, obi, style)
        sheet1.write(p + 3, col + 1, "dwh :", titleStyle)
        sheet1.write(p + 3, col + 2, dwh, style)
        sheet1.write(p + 4, col, "Oracle: ", titleStyle)
        sheet1.write(p + 4, col + 1, "147.234.159.200", style)
        sheet1.write(p + 4, col + 2, "Nms: ", titleStyle)
        sheet1.write(p + 4, col + 3, ver, style)

        return 8

    def printETL(p, etl):

        # calling the printing details func
        p += printDet(p)

        if len(etls[etl].flags) == 0:
            return
        sheet1.write(p, col, "ETL total time: ", titleStyle)
        sheet1.write(p, col + 1, str(etls[etl].totalTime()), style)
        if etls[etl].isPassed():
            msg = "Passed"
            s = succ
        else:
            msg = "Failed"
            s = fail
        sheet1.write(p, col + 2, msg, s)

        sheet1.write(p + 1, 0, "ETL flags: ", titleStyle)
        sheet1.write(p + 2, 0, "recovery mode: ", titleStyle)
        sheet1.write(p + 2, 1, etls[etl].flags[0], style)
        sheet1.write(p + 2, 2, "skip mode: ", titleStyle)
        sheet1.write(p + 2, 3, etls[etl].flags[1], style)
        sheet1.write(p + 2, 4, "transfer only(to): ", titleStyle)
        sheet1.write(p + 2, 5, etls[etl].flags[1], style)

        nonlocal row
        p += 4

        sheet1.write(p, col, "Timestamp", titleStyle)
        sheet1.write(p, col + 1, "Task", titleStyle)
        sheet1.write(p, col + 2, "Status", titleStyle)
        sheet1.write(p, col + 3, "Elaboration", titleStyle)
        p += 1
        row += 15

        for i in range(len(etls[etl].log[0])):
            sheet1.write(p, 1, str(etls[etl].log[0][i]), style)
            sheet1.write(p, 2, str(etls[etl].log[1][i]), style)
            if etls[etl].log[2][i] == "WARNING":
                sheet1.write(p, 3, str(etls[etl].log[2][i]), xlwt.easyxf('pattern: pattern solid, fore_colour yellow;'
                                                                         'font: colour black, bold True;'
                                                                         'align: horiz center'))
            elif etls[etl].log[2][i] == "FAILURE":
                sheet1.write(p, 3, str(etls[etl].log[2][i]), fail)
            elif etls[etl].log[2][i] == "SUCCESS":
                sheet1.write(p, 3, str(etls[etl].log[2][i]), succ)
            else:
                sheet1.write(p, 3, str(etls[etl].log[2][i]), style)

            sheet1.write(p, 4, str(etls[etl].log[3][i]), style)
            p += 1
            row += 1
        sheet1.write(p + 1, 1, set(etls[etl].err), style)
        row += 5

    for i in range(10):
        if (etls[i][0][0].month != curMonth):
            for i in range(6):
                sheet1.col(i).width = (15) * 367
            sheet1.col(4).width = (16) * 367

            sheet1 = log.add_sheet(months[etls[i][0][0].month])
            #nonlocal
        print(etls[i].err)
        printETL(row, i)

    sheet2.write(1, 1, "Average total time", titleStyle)
    sheet2.write(1, 2, str(calcAVG()), titleStyle)

    log.save("log.xls")




