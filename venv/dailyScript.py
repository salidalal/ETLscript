#!/usr/bin/env python3
# coding: utf-8
from datetime import datetime
import xlwt
import functools
from dailyScriptDefines import *
from sendMail import *
from monthlyScript import *
import os






#./collectLIDetails.sh
environments=[]




dirc = "//netstore2/sdh_tst_pub/TamarS/LIreport"


for filename in os.listdir(dirc):
    if filename.endswith("txt"):
        with open (os.path.join(dirc, filename)) as f:
            environments+=[Environment(f.read().splitlines())]



outputLog = xlwt.Workbook(encoding="utf-8")

for env in environments:
    env.printToFile(outputLog)



now = datetime.now().strftime("%d.%m - %b.xls")

print("C:/Users/sdalal/OneDrive - ECI Telecom LTD/PycharmProjects/untitled/venv/logs/"+now)
outputLog.save("C:/Users/sdalal/OneDrive - ECI Telecom LTD/PycharmProjects/untitled/venv/logs/"+now)


sendMail()
monthly()