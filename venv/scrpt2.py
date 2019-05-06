#!/usr/bin/env python

from datetime import datetime
import pandas as pd
import xlwt
import numpy as np
import functools

format = '%Y-%m-%d %H:%M:%S'
ErrorFormat = '%d/%m/%Y %H:%M:%S'
months = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]

my_sheet_name = 'Sheet1'
df = pd.read_excel('Execution Log.xlsx', sheet_name=my_sheet_name)
log = open("ETL.log", 'r')
setupDWH = open("CheckApps.txt", "r").readlines()

dwh = setupDWH[3].split()[2][1:-1]
ver = setupDWH[9].split()[1]
fix = "None"
for line in setupDWH:
    if "fix" in line.lower():
        fix = line.split()[1]
        break

setupDWH = open("SetupDWH.root.log", "r").readlines()
for line in setupDWH:
    if "Host name:" in line:
        obi = line.split()[2]
        break

# log = open ("/var/DWH/logs/ETL.log",'r')
lines = log.readlines()

msg = df.values.T[3].tolist()[-2:1:-1]
status = df.values.T[2].tolist()[-2:1:-1]
task = df.values.T[1].tolist()[-2:1:-1]
timeStamp = df.values.T[0].tolist()[-2:1:-1]


def calcTime(start, end):
    return end - start


class ETL:
    def __init__(self, start, end):
        self.start = start
        self.end = end
        self.log = []
        self.getMat(start, end)
        self.flags = self.getFlags()
        self.err = []

    def getMat(self, start, end):
        self.log += [timeStamp[start: end + 1]]
        for i in range(len(self.log[0])):
            if type(self.log[0][i]) == str:
                self.log[0][i] = datetime.strptime(self.log[0][i][:19], format)
        self.log += [task[start: end + 1]]
        self.log += [status[start: end + 1]]
        self.log += [msg[start: end + 1]]

    def totalTime(self):
        start = self.log[0][0]
        end = self.log[0][-1]
        return calcTime(start, end)

    def isPassed(self):
        return self.log[1][-1] == "Main Job" and self.log[2][-1] == "SUCCESS"

    def hasErrors(self):
        return "WARNING" in self.log[2]

    def getErrorsIndexs(self):
        locate = []
        if self.hasErrors():
            for i in range(len(self.log[2])):
                if self.log[2][i] == "WARNING":
                    locate += [i]
        return locate

    def getErrors(self):
        if (not self.hasErrors()):
            return []
        err = []
        for x in self.getErrorsIndexs():
            err += [self.log[1][x]]
        return set(err)

    def getFlags(self):
        f = []
        for line in lines:
            if str(self.log[0][0])[:-2] in line and "ETL parameters:" in line:
                line = iter(line.split())
                for w1 in line:
                    if "mode" in w1:
                        f += [next(line)[:-1]]
                    elif "skip" in w1:
                        f += [next(line)[:-1]]
                    elif "only(to)" in w1:
                        f += [next(line)]

        return f


etls = []
start = -1
fail = False

df = pd.read_excel('ETL Error Log_ETL Error Log.xlsx', sheet_name="ETL Error Log")
time = df.values.T[0].tolist()[8:]
# time = list(map(lambda x:datetime.strptime(x,ErrorFormat) ,df.values.T[0].tolist()[8:]))
cause = df.values.T[4].tolist()[8:]

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


def calcAVG():
    total = functools.reduce((lambda x, y: x + y), map(lambda x: x.totalTime(), etls))
    return total / len(etls)


log = xlwt.Workbook(encoding="utf-8")
sheet1 = log.add_sheet("Last ETL process")
sheet2 = log.add_sheet("statistics")

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
            nonlocal
        print(etls[i].err)
        printETL(row, i)

    sheet2.write(1, 1, "Average total time", titleStyle)
    sheet2.write(1, 2, str(calcAVG()), titleStyle)

    log.save("log.xls")


printToFile()


