import pandas as pd
import re
import os
import datetime 

#File path for Sensititre imports
fPath = 'L:/Micro/AR-TB/AR/ABI7500 import - export files/Sensititre Imports/'
fname = 'elims.xlsx'
reString = r"\.xlsx$"

for file in os.listdir(fPath):
    if re.search(reString, file):
        fname = file
        print(f"Found file: {fname}")
#ID Dictionary for ASTM formatting
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
          "Enterobacter asburiae" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Enterobacter bugandensis" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Enterobacter roggenkampii" : "E.CLC^Enterobacter cloacae complex^E.clocomplex",
          "Klebsiella aerogenes" : "K.AER^Klebsiella aerogenes^K.aerogenes",
          "Citrobacter freundii" : "C.FRE^Citrobacter freundii^C.freundii",
          "Providencia stuartii" : "P.STU^Providencia stuartii^P.stuartii",
          "Raoultella planticola" : "R.PLA^Raoultella planticola^R.plant",
          "Serratia marcescens" : "S.MAR^Serratia marcescens^S.marcescens",
          "Providencia rettgeri" : "P>RET^Providencia rettgeri^P.rettgeri"
}

class DataFrameProcessor:
    # Initialize with file path and name, read Excel into dataframe
    def __init__(self, _filePath:str, _fileName:str):
        self.filePath = _filePath
        self.fileName = _fileName
        self.df = pd.read_excel(self.filePath + self.fileName)
    def makeMCIM(self):
        #select only rows with ProfileName 'mCIM'
        mcimMask = (self.df['ProfileName'] == 'mCIM') 
        self.mcimDF = self.df[mcimMask]
        #select only rows with NaN Final Interpretation
        naMask = self.mcimDF['Final Interpretation'].isna()
        self.mcimDF = self.mcimDF[naMask]
        #create list of enumbers from Collectiondate column as specimens needing IDs
        self.enumbers = self.mcimDF['Collectiondate'].tolist()
    def makeID(self):
        #select only rows with ProfileName 'CRE - Identification'
        idMask = self.df['ProfileName'] == 'CRE - Identification'
        self.idDF = self.df[idMask]
        #remove nan
        self.idDF = self.idDF.dropna(subset=['Final Interpretation'])

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
    def write(self, _ASTM):
        _ASTM.write(self.oline)
        _ASTM.write(self.c1)
        _ASTM.write(self.c2)
        _ASTM.write(self.c3)
        _ASTM.write(self.c4)
        _ASTM.write(self.c5)
        _ASTM.write(self.c6)
        _ASTM.write(self.c7)
        _ASTM.write(self.c8)
        _ASTM.write(self.c9)

# Main Program

# Read in Excel file
#Create Masks and dataframes
df = DataFrameProcessor(fPath, fname)
df.makeMCIM()
df.makeID()

#Initialize ASTM file
newASTM = ASTM()
newASTM.open()
#Iterate through enumbers and create ASTM orders
for i, n in enumerate(df.enumbers):
    if df.idDF.loc[df.idDF['Collectiondate'] == n, 'Final Interpretation'].empty:
        print(f"No ID found for {n}")
    elif df.idDF.loc[df.idDF['Collectiondate'] == n, 'Final Interpretation'].item() in idDict:
        tmpid = str(df.idDF.loc[df.idDF['Collectiondate'] == n, 'Final Interpretation'].item())
        newASTM.addOrder(n, tmpid, i+1)
        print(f"ID for {n}: {df.idDF.loc[df.idDF['Collectiondate'] == n, 'Final Interpretation'].item()}")   
    elif df.idDF.loc[df.idDF['Collectiondate'] == n, 'Final Interpretation'].item() not in idDict:
        print(f"ID for {n}: {df.idDF.loc[df.idDF['Collectiondate'] == n, 'Final Interpretation'].item()} not in ID dictionary. Skipping.")
#Write orders to ASTM file
for order in newASTM.orders:
    order.write(newASTM)
newASTM.eFile.close()
os.rename(fPath + fname, fPath+'Processed/'+fname)