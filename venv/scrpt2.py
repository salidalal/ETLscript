






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
        print(etls[i].err)
        printETL(row, i)

    sheet2.write(1, 1, "Average total time", titleStyle)
    sheet2.write(1, 2, str(calcAVG()), titleStyle)

    log.save("log.xls")


printToFile()


