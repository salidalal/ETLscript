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









def calcAVG():
    total = functools.reduce((lambda x, y: x + y), map(lambda x: x.totalTime(), etls))
    return total / len(etls)













