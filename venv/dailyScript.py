#!/usr/bin/env python

from datetime import datetime
import xlwt
import functools
from dailyScriptDefines import *
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



now = datetime.datetime.now().strftime("%d.%m") + "-Etl analyse.xls"
outputLog.save(now)
