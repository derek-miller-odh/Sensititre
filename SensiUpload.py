import pandas as pd
import re
import os
import datetime 

reString = r"\.xlsx$"
DFs = []
fPath = 'L:/Micro/AR-TB/AR/ABI7500 import - export files/Sensititre Imports/'

idDict = {"Escherichia coli" : "E.COL^Escherichia coli^E.coli",
          "Klebsiella pneumoniae" : "K.PNEU^Klebsiella pneumoniae^Kleb.pneumo",
          "Enterobacter hormaechei" : "E.HOR^Enterobacter hormaechei^Ent.hormaech",
          "Pseudomonas aeruginosa" : "P.AER^Pseudomonas aeruginosa^P.aeruginosa",
          "Klebsiella oxytoca" : "K.OXY^Klebsiella oxytoca^K.oxytoca",
          "Enterobacter kobei" : "E.KOB^Enterobacter kobei^E.kobei",
          "Acinetobacter baumannii" : "A.BAU^Acinetobacter baumannii^A.baumannii",
          "Enterobacter cloacae complex" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Proteus mirabilis" : "P.MIR^Proteus mirabilis^P.mirabilis",
          "Raoultella ornithinolytica" : "R.ORN^Raoultella ornithinolytica^R.ornithinol",
          "Enterobacter absuriae" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Enterobacter bugandensis" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Enterobacter roggenkampii" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Klebsiella aerogenes" : "K.AER^Klebsiella aerogenes^K.aerogenes",
          "Citrobacter freundii" : "C.FRE^Citrobacter freundii^C.freundii",
          "Providencia stuartii" : "P.STU^Providencia stuartii^P.stuartii",
          "Raoultella planticola" : "R.PLA^Raoultella planticola^R.plant",
          "Serratia marcescens" : "S.MAR^Serratia marcescens^S.marcescens",
          "Providencia rettgeri" : "P>RET^Providencia rettgeri^P.rettgeri"
}

class ASTM:
    def __init__(self):
        tmpDate = datetime.datetime.now()
        self.date = tmpDate.strftime("%Y%m%d%H%M%S")
        self.orders = []
        self.makeHeader(self.date)
    def makeHeader(self, _date):
        """
        Take today's date and create header for ASTM File
        """
        self.head = "H|\\^&|||ODH Laboratory|||||||P|ASTM E1394-91|" + _date
    def addOrder(self, _spec, _id, _i):
        self.orders.append(Order(_spec, _id, _i))
    def open(self):
        self.eFileName = fPath + "transfer.txt"
        self.eFile = open(self.eFileName,"w")
        self.eFile.write("H|\\^&|||ODH Laboratory|||||||P|ASTM E1394-91|"+ self.date + "\n")
    def write(self,_line):
        self.eFile.write(_line)

class Order:
    def __init__(self, _spec:str, _id:str, _i:int):
        self.date = datetime.datetime.now()
        self.date = self.date.strftime("%Y%m%d%H%M%S")
        self.specimen = _spec
        self.id = idDict[_id]
        self.i = _i
        self.oline = "O|"+ str(self.i) + "|" + self.specimen + "^A|||R||||||N|||" + self.date + "|I^Isolate|^||Specimen Comment|||||||C||^|N\n"
        self.c1 = "C|1|L|Spec Coded Note 1:^^\n"
        self.c2 = "C|2|L|Spec Coded Note 2:^^\n"
        self.c3 = "C|3|L|Spec Coded Note 3:^^\n"
        self.c4 = "C|4|L|NOTE 1^\n"
        self.c5 = "C|5|L|NOTE 2^\n"
        self.c6 = "C|6|L|NOTE 3^\n"
        self.c7 = "C|7|L|PRELIMID^" + self.id + "\n"
        self.c8 = "C|8|L|PRE-DATE^" + self.date + "\n"
        self.c9 = "C|9|L|FINALID^" + self.id + "\n"
    def write(self):
        newASTM.write(self.oline)
        newASTM.write(self.c1)
        newASTM.write(self.c2)
        newASTM.write(self.c3)
        newASTM.write(self.c4)
        newASTM.write(self.c5)
        newASTM.write(self.c6)
        newASTM.write(self.c7)
        newASTM.write(self.c8)
        newASTM.write(self.c9)
def getDataset(_fname):
    tmpdf = pd.read_excel(fPath+_fname)
    return tmpdf

newASTM = ASTM()



mFiles = os.listdir(fPath)
for fname in mFiles:
    fname = str(fname)
    if re.search(reString, fname):
        DFs.append(getDataset(str(fname)))

eNumbers = DFs[0]['Specimen']
print(eNumbers)
for i,n in enumerate(eNumbers):
    tmpid = str(DFs[0].loc[DFs[0]['Specimen'] == n, 'Interpretation'].item())
    print(tmpid)
    newASTM.addOrder(n, tmpid, i+1)

newASTM.open()

for order in newASTM.orders:
    order.write()
today = datetime.datetime.now()
today = today.strftime("%Y%m%d%H%M%S")

