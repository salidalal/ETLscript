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
        return self.sortETL()[-1][1][1] == "SUCCESS" and self.sortETL()[0][1][0] == "MAIN JOB"

    def warning(self):
        for row in self.sortETL():
            if row[1][1]!="WARNING":
                return true
        return false

    def printToFile(self,outputLog):
        sheet = outputLog.add_sheet(self.getOBIhost)



        row = 1
        col = 0








        sheet.write(row, col, "OBI Host Name: ", titleStyle)
        sheet.write(row + 1, col, self.getOBIhost, style)
        sheet.write(row, col+1, "DWH Host Name: ", titleStyle)
        sheet.write(row + 1, col + 1, self.getDWHhost, style)
        sheet.write(row, col+2, "Last ETL status: ", titleStyle)


        if(self.warning and self.isPass):
            sheet.write(row + 1, col + 2, "SUCCESS with WARNING", succ)
        elif (self.isPass):
            sheet.write(row + 1, col + 2, "ETL SUCCESS", succ)
        else:
            sheet.write(row + 1, col + 2, "ETL Failed", fail)

        col+=3
        sheet.write(row, col, "Starting time: ", titleStyle)
        sheet.write(row + 1, col, str(self.getETLStart)[:19], style)
        sheet.write(row, col + 1, "Total time: ", titleStyle)
        sheet.write(row + 1, col + 1, str(self.etlTime)[:7], style)
        sheet.write(row, col + 2, "OBI / DWH Version: ", titleStyle)
        if (self.getDWHver!=self.getOBIver):
            sheet.write(row + 1, col + 2, self.getOBIver + "/" + self.getDWHver, style)
        else:
            sheet.write(row + 1, col + 2, self.getOBIver, style)

        sheet.write(row, col + 3, "OBI Fix: ", titleStyle)
        sheet.write(row + 1, col + 3, self.getOBIFix, style)
        sheet.write(row, col + 4, "DWH Fix: ", titleStyle)
        sheet.write(row + 1, col + 4, self.getDWHFix, style)

        col+=5
        """
        sheet.write(row, col + 5, "ECILobieeInst1: ", titleStyle)
        sheet.write(row + 1, col + 5, self.getObieeInst1, style)
        sheet.write(row, col + 6, "ECILobieeInst2: ", titleStyle)
        sheet.write(row + 1, col + 6, self.getObieeInst2, style)
        col +=2
        """


        #flags
        sheet.write(row, col, "Incremental mode : ", titleStyle)
        sheet.write(row + 1, col, self.getInc, style)
        sheet.write(row, col + 1, "Failure Management: ", titleStyle)
        sheet.write(row + 1, col + 1, self.getFailMng, style)
        sheet.write(row, col + 2, "Skip Mode: ", titleStyle)
        sheet.write(row + 1, col + 2, self.getSkip, style)
        sheet.write(row, col + 3, "Recovery Mode: ", titleStyle)
        sheet.write(row + 1, col + 3, self.getRecovery, style)
        sheet.write(row, col + 4, "Transfer Only: ", titleStyle)
        sheet.write(row + 1, col + 4, self.getTO, style)
        sheet.write(row, col + 5 , "mirror DB params: ", titleStyle)
        sheet.write(row+1, col + 5, self.getParams, style)





        row+=3
        sheet.write(row, 1, "Job Name: ", titleStyle)
        sheet.write(row, 2, "Status: ", titleStyle)
        sheet.write(row, 3, "Massages: ", titleStyle)
        sheet.write(row, 4, "Job total time: ", titleStyle)
        for r in self.sortETL():
            sheet.write(row+1, 1, r[1][0], style) #job name
            sheet.write(row+1, 2, r[1][1], style) # status
            sheet.write(row+1, 3, r[1][2], style) # msg

            if r[1][1] == "SUCCESS":
                for s in self.sortETL():
                    if s[1][0] == r[1][0]:
                        sheet.write(row+1, 4, str(calcTime(s[0], r[0])), style)  # total time
                        break
            row+=1


        for i in range(20):
            sheet.col(i).width = (15) * 367
