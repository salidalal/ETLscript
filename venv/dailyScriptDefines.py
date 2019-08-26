#!/usr/bin/env python

import xlwt
import datetime
from datetime import timedelta

style = xlwt.easyxf('align: horiz center')
succ = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                   'font: colour white, bold True;'
                   'align: horiz center')
warr = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;'
                   'font: colour black, bold True;'
                   'align: horiz center')
fail = xlwt.easyxf('pattern: pattern solid, fore_colour red;'
                   'font: colour white, bold True;'
                   'align: horiz center')
titleStyle = xlwt.easyxf("font: bold True; align: horiz center")
format = '%Y-%m-%d %H:%M:%S.%f'

mirror = {"-q": "quiet mode",
          "-skip_pm": "skip PM collection",
          "-skip_cfm": "skip CFM PM collection",
          "-skip_service": "skip service tables",
          "-skip_pm_service": "skip PM service collection",
          "-skip_15_pm_service": "skip 15Min PM service collection",
          "-skip_15_pm": "skip 15Min PM collection",
          "-skip_sdh": "skip sdh tables",
          "-skip_optics": "skip optics tables",
          "-skip_otn_performance": "skip optics PM collection",
          "-skip_l3vpn_inventory": "skip l3vpn network inventory",
          "-skip_configuration": "skip configuration validation",
          "-skip_health": "skip health model",
          "-clean_duplicate_pm": "clean duplicate PM records",
          "-pm_csv_path n": "full path to a given pm_source_data.csv file",
          "-user_ports": "upload user ports",
          "-to": "transfer only",
          "-db": "mirror tables using database only",
          "-na": "no alarms import",
          "-h": "this message"
          }


def calcTime(start, end):


    e=(end-start)

    #e.hour += e.days*24
    return e


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
        params=[]
        for line in self.envfile:
            if "mirror_db_param" in line:
                params= str(line.split('=')[1]).split()
                break
        dbparams=""

        for param in params:
                if (param in mirror):
                    dbparams+=str(mirror[param])
                    dbparams+=" | "
        return dbparams



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
                return self.envfile[i + 2:]

    def sortETL(self):
        from collections import OrderedDict

        etl = {}
        ret = []
        for line in self.getETL:
            etl.update({datetime.datetime.strptime(line.split(',')[0], format): line.split(',')[2:]})
        for key in (OrderedDict(sorted(etl.items(), key=lambda x: x))):
            ret += [[key, etl[key]]]

            cou = 0

        return ret

    @property
    def etlTime(self):
        return calcTime(self.getETLStart, self.getETLEnd)

    @property
    def isPass(self):
        return self.sortETL()[-1][1][1] == "SUCCESS" and self.sortETL()[0][1][0] == "MAIN JOB"

    @property
    def isNotDone(self):
        return self.sortETL()[-1][1][0] != "MAIN JOB"

    @property
    def getNMS(self):
        for i in range(len(self.envfile)):
            if "NMS_IP" in self.envfile[i]:
                return self.envfile[i][6:].split("/")[0][1:]
    @property
    def getNE(self):
        for i in range(len(self.envfile)):
            if "NMS_IP" in self.envfile[i]:
                return self.envfile[i][6:].split("/")[1].split()[1]

    @property
    def getPM(self):
        for i in range(len(self.envfile)):
            if "pm_source_data.csv size" in self.envfile[i]:
                return self.envfile[i].split()[-1]


    def warning(self):
        for row in self.sortETL():
            if row[1][1] == "WARNING":
                return True
        return False

    def printToFile(self, outputLog):
        try:
            sheet = outputLog.add_sheet(self.getOBIhost)
        except:
            sheet = outputLog.add_sheet(self.getOBIhost+"2")

        row = 0
        col = 0

        sheet.write(row, col, "OBI Host Name: ", titleStyle)
        sheet.write(row + 1, col, self.getOBIhost, style)
        sheet.write(row, col + 1, "DWH Host Name: ", titleStyle)
        sheet.write(row + 1, col + 1, self.getDWHhost, style)
        sheet.write(row, col + 2, "Last ETL status: ", titleStyle)

        if self.isNotDone:
            sheet.write(row + 1, col + 2, "ETL is not done!", warr)
        elif (self.warning() and self.isPass):
            sheet.write(row + 1, col + 2, "SUCCESS with WARNING", warr)
        elif (self.isPass):
            sheet.write(row + 1, col + 2, "ETL SUCCESS", succ)
        else:
            sheet.write(row + 1, col + 2, "ETL Failed", fail)

        col += 3
        sheet.write(row, col, "Starting time: ", titleStyle)
        sheet.write(row + 1, col, str(self.getETLStart)[:19], style)
        sheet.write(row, col + 1, "Total time: ", titleStyle)
        sheet.write(row + 1, col + 1, str(self.etlTime), style)
        sheet.write(row, col + 2, "OBI / DWH Version: ", titleStyle)

        if (self.getDWHver != self.getOBIver):
            sheet.write(row + 1, col + 2, self.getOBIver + "/" + self.getDWHver, titleStyle)
        else:
            sheet.write(row + 1, col + 2, self.getOBIver, style)

        sheet.write(row, col + 3, "OBI Fix: ", titleStyle)
        sheet.write(row + 1, col + 3, self.getOBIFix, style)
        sheet.write(row, col + 4, "DWH Fix: ", titleStyle)
        sheet.write(row + 1, col + 4, self.getDWHFix, style)

        col += 5
        """
        sheet.write(row, col + 5, "ECILobieeInst1: ", titleStyle)
        sheet.write(row + 1, col + 5, self.getObieeInst1, style)
        sheet.write(row, col + 6, "ECILobieeInst2: ", titleStyle)
        sheet.write(row + 1, col + 6, self.getObieeInst2, style)
        col +=2
        """

        # flags
        sheet.write(row, col, "Incremental mode : ", titleStyle)
        sheet.write(row + 1, col, self.getInc, style)
        sheet.write(row, col + 1, "Failure Management: ", titleStyle)
        sheet.write(row + 1, col + 1, self.getFailMng, style)
        sheet.write(row, col + 2, "Skip Mode: ", titleStyle)
        sheet.write(row + 1, col + 2, self.getSkip, style)
        sheet.write(row, col + 3, "Recovery Mode: ", titleStyle)
        sheet.write(row + 1, col + 3, self.getRecovery, style)
        sheet.write(row, col + 4, "mirror DB params: ", titleStyle)
        sheet.write(row + 1, col + 4, self.getParams, style)
        sheet.write(row, col + 5, "NMS IP: ", titleStyle)
        sheet.write(row + 1, col + 5, self.getNMS, style)
        sheet.write(row, col + 6, "NE count: ", titleStyle)
        sheet.write(row + 1, col + 6, self.getNE, style)
        sheet.write(row, col + 7, "PM size: ", titleStyle)
        sheet.write(row + 1, col + 7, self.getPM, style)

        row += 3
        sheet.write(row, 1, "Job Name: ", titleStyle)
        sheet.write(row, 2, "Status: ", titleStyle)
        sheet.write(row, 3, "Massages: ", titleStyle)
        sheet.write(row, 4, "Job total time: ", titleStyle)
        for r in self.sortETL():
            sheet.write(row + 1, 1, r[1][0], style)  # job name
            sheet.write(row + 1, 2, r[1][1], style)  # status
            sheet.write(row + 1, 3, r[1][2], style)  # msg



            if r[1][1] == "SUCCESS":
                for s in self.sortETL():
                    if s[1][0] == r[1][0]:
                        sheet.write(row + 1, 4, str(calcTime(s[0], r[0])), style)  # total time
                        break
            row += 1

        self.isNotDone
        #settings cols size
        for i in range(20):
            sheet.col(i).width = (15) * 367
