import openpyxl as opxl
from progress.bar import IncrementalBar
import datetime
import database

class Logbook:

    def __init__(self):
        # try to open 'APN'.xls
        try:
            self.wb1225 = opxl.open("1225.xlsx")
            self.wb1328 = opxl.open("1328.xlsx")
            self.wb1498 = opxl.open("1498.xlsx")
        except:
            print("Не могу открыть какую-то таблицу")
            exit()

        # load worksheets
        self.ws1225 = self.wb1225.active
        self.ws1328 = self.wb1328.active
        self.ws1498 = self.wb1498.active
        self.wblb = opxl.Workbook("Logbook.xlsx")
        self.wslb = self.wblb.active

        # get nessesary columns` number 
        self.tow = self.__getInfo(self.ws1225, "ToW")
        self.cmproject = self.__getInfo(self.ws1225, "CM Project")
        self.sbt = self.__getInfo(self.ws1225, "User Sign")
        self.date = self.__getInfo(self.ws1225, "End Date")
        self.description = self.__getInfo(self.ws1328, "Description")
        self.pn = self.__getInfo(self.ws1328, "Part No.")
        self.cmproject1328 = self.__getInfo(self.ws1328, "CM Project No.")
        self.sn = self.__getInfo(self.ws1328, "Serial No. / Batch No.")
        self.timeDuration = self.__getInfo(self.ws1225, "Used MHR")
        self.ata = self.__getInfo(self.ws1328, "ATA Chapter")

    def __getInfo(self, ws, info): # find column`s number in worksheet by info
        for i in range(ws.max_column):
            if ws[1][i].value == info:
                return i


    def clear(self, ws):    # clear worksheet from empty rows by 'A' column
        clearing = IncrementalBar("Clearing...", max = ws.max_row-1)
        for i in range(ws.max_row,1,-1):
            if ws[i][0].value == None:
                ws.delete_rows(i)
            clearing.next()
        clearing.finish()
        return ws


    def formateRow(self, row):      # formate logbook row like a array of string
        row1225 = self.ws1225[row]
        row1328 = self.getRow1328(row1225[self.cmproject])
        usedTime = self.getTime(row1225[self.timeDuration])
        task = self.getTaskType()
        data = [
                row1225[self.date].value.strftime('%d.%m.%y'),    # date
                "ACCMS",    # location "ACCMS"
                row1328[self.description].value,    # Component type (description)
                row1328[self.sn].value,    # Serial number
                "C6 Equipment",    # Type of maintenance(raiting)
                "trainee",    # Privelege
                task["fot"],    # task type: FOT 
                task["sgh"],    # task type: SGH
                task["r/i"],    # task type: R/I
                task["ts"],    # task type: TS
                task["mod"],    # task type: MOD
                task["rep"],    # task type: REP
                task["insp"],    # task type: INSP
                    # type of activity: Training
                    # type of activity: Perform
                    # type of activity: Suppervise
                    # type of activity: CRS
                row1328[self.ata].value,    # ATA
                    # Operation performed
                usedTime,    # Time duration
                row1225[self.cmproject].value,    # Maintenance record ref.(CM number)
                "ToW is \"{}\"".format(row1225[self.tow].value)    # Remarks
        ]
        return data


 
    def writeRow(self, sheet, data):    # write row in the sheet with SBT name
        if sheet not in self.wblb.sheetnames:
            ws = self.wblb.create_sheet(sheet)
            ws.title = sheet
        else:
            ws = self.wblb[sheet]
        print("sheet is {}, data is {}".format(sheet,data))
        ws.append(data)


    def getRow1328(self, cm):       # get row from 1328 apn by CM project
        for row in self.ws1328:
            if cm.value in row[self.cmproject1328].value:
                return row


    def getTime(self, time):        # get time by delta seconds
        hours, remainder = divmod(time.value.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours}:{minutes}"


    def getTaskType(self):      # get task type from apn1225, 1328 and 1498
        return {
            "fot":"",
            "sgh":"",
            "r/i":"",
            "ts":"",
            "mod":"",
            "rep":"",
            "insp":"",
        }


    def save(self):     # save logbook 
        self.wblb.save("Logbook.xlsx")


