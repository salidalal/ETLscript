from datetime import datetime
format = '%Y-%m-%d %H:%M:%S'
ErrorFormat = '%d/%m/%Y %H:%M:%S'


def calcTime(start, end):
    return end - start


def calcAVG():
    total = functools.reduce((lambda x, y: x + y), map(lambda x: x.totalTime(), etls))
    return total / len(etls)













