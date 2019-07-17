#!/usr/bin/env python3
# coding: utf-8
from datetime import datetime
from xlutils.copy import copy
import xlrd
from weeklyScriptsDefines import *
import os




environmentsLogs = {}

logsDirc = "C:/Users/sdalal/OneDrive - ECI Telecom LTD/PycharmProjects/untitled/venv/logs"
envDirc = "//netstore2/sdh_tst_pub/TamarS/LIreport"


# getting enviroments into a dictionary where the obi name is the key and empty list is the value
for filename in os.listdir(envDirc):
    if filename.endswith("txt"):
        with open (os.path.join(envDirc, filename)) as f:
            envName = filename.split(".")[0]
            if envName not in environmentsLogs.keys():
                environmentsLogs[envName] = []




cur_month = datetime.now()
cur_month = cur_month.strftime("%b")

#looking for the daily report in the current month and put the data in the dictionary list for each enviro.
for filename in os.listdir(logsDirc):
    if filename.endswith(cur_month+".xls"):
        addToList(os.path.join(logsDirc, filename), environmentsLogs)


"""
#adding sheet for each env.
for envName in environmentsLogs.keys():
    outputLog = xlwt.Workbook(encoding="utf-8")
    try:
        temp = xlrd.open_workbook("C:/Users/sdalal/OneDrive - ECI Telecom LTD/PycharmProjects/untitled/venv/logs/" + envName)
        outputLog = copy(temp)
        temp.sheet_by_name()
    except:
        outputLog = xlwt.Workbook()

"""
outputLog = xlwt.Workbook(encoding="utf-8")

#calling print to file for each sheet ( each shhet for each env)
for env in environmentsLogs:
    print(env)
    printToFile(outputLog, env, environmentsLogs, cur_month)

outputLog.save("C:/Users/sdalal/OneDrive - ECI Telecom LTD/PycharmProjects/untitled/venv/logs/" + cur_month+".xls")
