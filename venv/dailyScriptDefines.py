#!/usr/bin/env python

import xlwt
import datetime

style = xlwt.easyxf('align: horiz center')
succ = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                       'font: colour white, bold True;'
                       'align: horiz center')
fail = xlwt.easyxf('pattern: pattern solid, fore_colour red;'
                       'font: colour white, bold True;'
                       'align: horiz center')
titleStyle = xlwt.easyxf("font: bold True; align: horiz center")
format = '%Y-%m-%d %H:%M:%S.%f'


def calcTime(start, end):
    return end - start



class Environment:
    def __init__(self, envfil):
        self.envfile = envfil


    ######### DWH #########
    @property
    def getDWHhost(self):
        for line in self.envfile:
            if "DWH hostname" in line:
                return line.split()[2]

    @property
    def getDWHFix(self):
        for line in self.envfile:
            if "DW" in line and "FixApp" in line:
                return line.split()[0]
        return "None"

    @property
    def getDWHver(self):
        for line in self.envfile:
            if "ECILdwh" in line:
                return line.split('=')[1]

    @property
    def getParams(self):
        for line in self.envfile:
            if "mirror_db_param" in line:
                return line.split('=')[1]

    @property
    def getInc(self):
        for line in self.envfile:
            if "inc_load" in line:
                return line.split('=')[1]

    @property
    def getRecovery(self):
        for line in self.envfile:
            if "recoverymode" in line:
                return line.split('=')[1]

    @property
    def getFailMng(self):
        for line in self.envfile:
            if "fail" in line:
                return line.split('=')[1]

    @property
    def getSkip(self):
        for line in self.envfile:
            if "skip=" in line:
                return line.split('=')[1]

    @property
    def getTO(self):
        for line in self.envfile:
            if "(to)" in line:
                return line.split('=')[1]


    ######### OBI #########
    @property
    def getOBIhost(self):
        for line in self.envfile:
            if "OBI host" in line:
                return line.split()[2]

    @property
    def getOBIFix(self):
        for line in self.envfile:
            if "OB" in line and "FixApp" in line:
                return line.split()[0]
        return "None"

    @property
    def getOBIver(self):
        for line in self.envfile:
            if "ECILobiee" in line:
                return line.split('=')[1]

    @property
    def getObieeInst1(self):
        for line in self.envfile:
            if "ECILobieeInst1" in line:
                return line.split('=')[1]

    @property
    def getObieeInst2(self):
        for line in self.envfile:
            if "ECILobieeInst1" in line:
                return line.split('=')[1]




    ######### ETL #########
    @property
    def getETLStart(self):
        return self.sortETL()[0][0]

    @property
    def getETLEnd(self):
        return self.sortETL()[-1][0]

    @property
    def getETL(self):
        for i in range(len(self.envfile)):
            if "excutaion log" in self.envfile[i]:
                return self.envfile[i+2:]

    def sortETL(self):
        from collections import OrderedDict

        etl={}
        ret = []
        for line in self.getETL:
            etl.update( {datetime.datetime.strptime ( line.split(',')[0],format): line.split(',')[2:]}  )
        for key in (OrderedDict( sorted(etl.items(), key=lambda x: x) )):
            ret+=[(key,etl[key])]
        return ret

    @property
    def etlTime(self):
        return calcTime(self.getETLStart,self.getETLEnd)

    @property
    def isPass(self):
        return self.sortETL()[-1][1][1]

    def warning(self):
        for row in self.sortETL():
            if row[1][1]!="START" and row[1][1]!="SUCCESS":
                return true
        return false

    def printToFile(self,outputLog):
        sheet = outputLog.add_sheet(self.getOBIhost)



        row = 3
        col = 0

        #obi prints
        sheet.write(row, col, "OBI: ", titleStyle)
        sheet.write(row, col+1, "Host Name: ", titleStyle)
        sheet.write(row, col + 2, self.getOBIhost, style)
        sheet.write(row, col + 3, "OBI Version: ", titleStyle)
        sheet.write(row, col + 4, self.getOBIver, style)
        sheet.write(row, col + 5, "OBI Fix: ", titleStyle)
        sheet.write(row, col + 6, self.getOBIFix, style)
        sheet.write(row, col + 7, "ECILobieeInst1: ", titleStyle)
        sheet.write(row, col + 8, self.getObieeInst1, style)
        sheet.write(row, col + 9, "ECILobieeInst2: ", titleStyle)
        sheet.write(row, col + 10, self.getObieeInst2, style)

        row = 5


        #dwh prints
        sheet.write(row, col, "DWH: ", titleStyle)
        sheet.write(row, col+1, "Host Name: ", titleStyle)
        sheet.write(row, col + 2, self.getDWHhost, style)
        sheet.write(row, col + 3, "OBI Version: ", titleStyle)
        sheet.write(row, col + 4, self.getDWHver, style)
        sheet.write(row, col + 5, "OBI Fix: ", titleStyle)
        sheet.write(row, col + 6, self.getDWHFix, style)

        row = 7


        #flags
        sheet.write(row , col, "Flags: ", titleStyle)
        sheet.write(row, col + 1, "Inc Load : ", titleStyle)
        sheet.write(row, col + 2, self.getInc, style)
        sheet.write(row, col + 3, "Failure Management: ", titleStyle)
        sheet.write(row, col + 4, self.getFailMng, style)
        sheet.write(row, col + 5, "Skip Mode: ", titleStyle)
        sheet.write(row, col + 6, self.getSkip, style)
        sheet.write(row, col + 7, "Skip Mode: ", titleStyle)
        sheet.write(row, col + 8, self.getSkip, style)
        sheet.write(row, col + 9, "Transfer Only: ", titleStyle)
        sheet.write(row, col + 10, self.getTO, style)
        sheet.write(row+1, col , "mirror DB params: ", titleStyle)
        sheet.write(row+1, col + 1, self.getParams, style)

        row=11
        sheet.write(row , col, "ETL: ", titleStyle)
        sheet.write(row, col + 2, "Starting time: ", titleStyle)
        sheet.write(row, col + 3, str(self.getETLStart)[:19], style)
        sheet.write(row, col + 4, "Ending time: ", titleStyle)
        sheet.write(row, col + 5, str(self.getETLEnd)[:19], style)
        sheet.write(row, col + 6, "Total time: ", titleStyle)
        sheet.write(row, col + 7, str(self.etlTime)[:7], style)

        if(self.isPass =="SUCCESS"):
            sheet.write(row, col + 1,self.isPass, succ)
        else:
            sheet.write(row, col + 1,"ETL Failed", fail)

        row+=1
        for r in self.sortETL():
            sheet.write(row, col+1, r[1][0], style)
            sheet.write(row, col + 2, r[1][1], style)
            sheet.write(row, col + 3, r[1][2], style)
            row+=1

        for i in range(12):
            sheet.col(i).width = (15) * 367
