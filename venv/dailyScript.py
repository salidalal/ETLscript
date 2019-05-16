class environment:
    def __init__(self, envfile):
        self.envfile = envfile

    ######### DWH #########
    @property
    def getDWHhost(self):
        return self.envfile[0].split()[2]

    @property
    def getDWHFix(self):
        return self.envfile[1].split()[0]

    @property
    def getDWHver(self):
        return self.envfile[2].split('=')[1]

    @property
    def getParams(self):
        return self.envfile[3].split('=')[1]

    @property
    def getInc(self):
        return self.envfile[4].split('=')[1]

    @property
    def getRecovery(self):
        return self.envfile[6].split('=')[1]

    @property
    def getFailMng(self):
        return self.envfile[5].split('=')[1]

    @property
    def getSkip(self):
        return self.envfile[7].split('=')[1]

    @property
    def getTO(self):
        return self.envfile[5].split('=')[1]


    ######### OBI #########
    @property
    def getOBIhost(self):
        return self.envfile[11].split()[2]

    @property
    def getOBIFix(self):
        return self.envfile[15].split()[0]

    @property
    def getOBIver(self):
        return self.envfile[12].split('=')[1]

    @property
    def getObieeInst1(self):
        return self.envfile[13].split('=')[1]

    @property
    def getObieeInst2(self):
        return self.envfile[14].split('=')[1]




    ######### OBI #########
    @property
    def getETLStrart(self):
        return self.envfile[14].split('=')[1]





