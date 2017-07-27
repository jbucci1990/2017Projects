import sys
import os
from PyQt5.QtWidgets import (QListWidget, QComboBox, QLineEdit, QTextEdit, QAction, QFileDialog, QApplication, QPushButton, QWidget, QLabel)

import xlrd

import xlsxwriter
from collections import Counter
import itertools



global uDates

uDates = []


global name_List
name_List = ["Cortese", "Lallo", "Dixon", "Busch", "Narain", "Conetta", "Hurd", "Malave", "Friedland",
             "Turrentine", "Carman", "Sarno", "Greenfield", "Plaut", "Dwyer", "Ferraiolo", "Ma", "Duffy",
             "Nardella", "O'Brien"]
name_List.sort()



global admin_List
admin_List = ["Bitetto", "DeFonzo", "Feibusch", "McClain", "Messina", "Palgon", "Patel", "Sidoti", "Tampellini",
              "Yorke"]

global downstairs_List
downstairs_List = ["Maldonado", "De La Rosa", "Stein", "Amadeo", "Bradley", "Presser", "Rosen"]
downstairs_List.sort()


global Trainer_list
Trainer_list = ["Espey", "Akins", "Fried", "Dellatorri"]
Trainer_list.sort()

global fulltime_List
fulltime_List = ["Butto", "Carey", "Courtright", "Dolega", "James", "Maple", "Richman", "Sensale", "Spellman"]


class Example(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.button = QPushButton('Pick File', self)
        self.button.clicked.connect(self.showDialog)
        self.button.move(20,20)
        self.label = QLabel('______', self)
        self.label.setFixedWidth(250)

        self.label.move(120,27)

        self.runProgram_Btn = QPushButton('Run', self)
        self.runProgram_Btn.move(20, 70)
        self.runProgram_Btn.clicked.connect(self.runProgram)


        self.employee = QLineEdit(self)
        self.employee.move(430, 27)

        self.employeeRank = QComboBox(self)
        self.employeeRank.addItem("Admin")
        self.employeeRank.addItem("5Q")
        self.employeeRank.addItem("Operator")
        self.employeeRank.addItem("Vets")
        self.employeeRank.addItem("Fulltime")
        self.employeeRank.move(580,24)

        self.addEmployee = QPushButton('Add', self)
        self.addEmployee.move(705, 24)
        self.addEmployee.clicked.connect(self.addE)

        # self.removeEmployee = QPushButton('Remove', self)
        # self.removeEmployee.move(780, 24)
        self.hoursLabel = QLabel('Total Hours', self)
        self.hoursLabel.move(180, 65)
        self.hoursList = QListWidget(self)
        self.hoursList.setFixedWidth(150)
        self.hoursList.move(150,85)


        self.adminLabel = QLabel("Admin", self)
        self.adminLabel.move(445, 65)
        self.adminList = QListWidget(self)
        self.adminList.setFixedWidth(100)
        self.adminList.move (430, 85)

        self.operatorLabel = QLabel("Operator", self)
        self.operatorLabel.move(550, 65)
        self.operatorList = QListWidget(self)
        self.operatorList.setFixedWidth(100)
        self.operatorList.move(535, 85)
        self.operatorList.itemClicked.connect(self.printSelection)

        self.fiveQLabel = QLabel("5Q", self)
        self.fiveQLabel.move(790, 65)
        self.fiveQList = QListWidget(self)
        self.fiveQList.setFixedWidth(100)
        self.fiveQList.move(745, 85)

        self.vetLabel = QLabel("Vets", self)
        self.vetLabel.move(670, 65)
        self.vetList = QListWidget(self)
        self.vetList.setFixedWidth(100)
        self.vetList.move(640, 85)

        self.fullTimeLabel = QLabel("FullTime", self)
        self.fullTimeLabel.move(880, 65)
        self.fullTimeList = QListWidget(self)
        self.fullTimeList.setFixedWidth(100)
        self.fullTimeList.move(850, 85)


        self.unavailableLabel = QLabel("Pick Date", self)
        self.unavailableLabel.move(425, 285)

        self.dateUnavailable_Btn = QComboBox(self)
        self.dateUnavailable_Btn.addItem("(Show All)")
        self.dateUnavailable_Btn.addItem("Monday")
        self.dateUnavailable_Btn.addItem("Tuesday")
        self.dateUnavailable_Btn.addItem("Wednesday")
        self.dateUnavailable_Btn.addItem("Thursday")
        self.dateUnavailable_Btn.addItem("Friday")
        self.dateUnavailable_Btn.addItem("Saturday")
        self.dateUnavailable_Btn.addItem("Sunday")
        self.dateUnavailable_Btn.move(490, 280)

        self.uRefresh_Btn = QPushButton('Refresh', self)
        self.uRefresh_Btn.move(640, 279)
        self.uRefresh_Btn.clicked.connect(self.filterUDates)

        self.warningLabel = QLabel('                                           ', self)
        self.warningLabel.move(750, 279)

        self.doubleBookedLabel = QLabel('Double Booked', self)
        self.doubleBookedLabel.move(275, 320)
        self.doubleBookedList = QListWidget(self)
        self.doubleBookedList.move(150,340)
        self.doubleBookedList.setFixedHeight(100)
        self.doubleBookedList.setFixedWidth(400)

        self.unavailableBookedLabel = QLabel('Unavailable that you Booked', self)
        self.unavailableBookedLabel.move(680, 320)
        self.unavailableBooked = QListWidget(self)
        self.unavailableBooked.move(580, 340)
        self.unavailableBooked.setFixedHeight(100)
        self.unavailableBooked.setFixedWidth(375)

        self.shortTurnAroundLabel = QLabel('Short Turnaround', self)
        self.shortTurnAroundLabel.move(275, 440)
        self.shortTurnAround = QListWidget(self)
        self.shortTurnAround.move(150, 460)
        self.shortTurnAround.setFixedWidth(400)
        self.shortTurnAround.setFixedHeight(100)

        self.operatorList.clear()
        for item in name_List:
            self.operatorList.addItem(item)

        self.adminList.clear()
        for item in admin_List:
            self.adminList.addItem(item)

        self.fiveQList.clear()
        for item in downstairs_List:
            self.fiveQList.addItem(item)

        self.fullTimeList.clear()
        for item in fulltime_List:
            self.fullTimeList.addItem(item)

        self.vetList.clear()
        for item in Trainer_list:
            self.vetList.addItem(item)

        self.setGeometry(1020, 650, 1020, 650)
        self.setWindowTitle('OnDeck Scheduling')




        self.show()

    def allEmployees(self):
        self.operatorList.clear()
        self.adminList.clear()
        self.fiveQList.clear()
        self.fullTimeList.clear()
        self.vetList.clear()



    def filterUDates(self):

        try:
            if str(self.dateUnavailable_Btn.currentText()) == "(Show All)":
                self.operatorList.clear()
                for item in name_List:
                    self.operatorList.addItem(item)

                self.adminList.clear()
                for item in admin_List:
                    self.adminList.addItem(item)

                self.fiveQList.clear()
                for item in downstairs_List:
                    self.fiveQList.addItem(item)

                self.fullTimeList.clear()
                for item in fulltime_List:
                    self.fullTimeList.addItem(item)

                self.vetList.clear()
                for item in Trainer_list:
                    self.vetList.addItem(item)


            if str(self.dateUnavailable_Btn.currentText()) == "Monday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[0]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)



                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)

            if str(self.dateUnavailable_Btn.currentText()) == "Tuesday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[1]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)

                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)

            if str(self.dateUnavailable_Btn.currentText()) == "Wednesday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[2]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)

                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)

            if str(self.dateUnavailable_Btn.currentText()) == "Thursday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[3]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)

                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)

            if str(self.dateUnavailable_Btn.currentText()) == "Friday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[4]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)

                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)

            if str(self.dateUnavailable_Btn.currentText()) == "Saturday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[5]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)

                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)

            if str(self.dateUnavailable_Btn.currentText()) == "Sunday":

                self.adminList.clear()
                self.fiveQList.clear()
                self.operatorList.clear()
                self.fullTimeList.clear()
                self.vetList.clear()

                adminU = admin_List[:]
                operatorU = name_List[:]
                fulltimeU = fulltime_List[:]
                downstairsU = downstairs_List[:]
                vetU = Trainer_list[:]

                namestoDelete = []
                mondayU = uDates[6]
                print(mondayU)
                for employee in mondayU:
                    namestoDelete.append(employee[1])
                    print(employee[1])

                for name in namestoDelete:
                    if name in adminU:
                        adminU.remove(name)
                    if name in operatorU:
                        operatorU.remove(name)
                    if name in fulltimeU:
                        fulltimeU.remove(name)
                    if name in downstairsU:
                        downstairsU.remove(name)
                    if name in vetU:
                        vetU.remove(name)

                for item in adminU:
                    self.adminList.addItem(item)
                for item in operatorU:
                    self.operatorList.addItem(item)
                for item in downstairsU:
                    self.fiveQList.addItem(item)
                for item in fulltimeU:
                    self.fullTimeList.addItem(item)
                for item in vetU:
                    self.vetList.addItem(item)
        except Exception as e:
            print("forgot to run")
            self.warningLabel.setText('Must Run Program First!')
    def addE(self):

        if str(self.employeeRank.currentText()) == "Admin":
            admin = (str(self.employee.text()))
            admin_List.append(admin)
            print(admin_List)
            self.adminList.clear()
            admin_List.sort()
            for item in admin_List:
                self.adminList.addItem(item)

        if str(self.employeeRank.currentText()) == "5Q":
            downstairs = (str(self.employee.text()))
            downstairs_List.append(downstairs)
            self.fiveQList.clear()
            downstairs_List.sort()
            for item in downstairs_List:
                self.fiveQList.addItem(item)


        if str(self.employeeRank.currentText()) == "Fulltime":
            fulltime = (str(self.employee.text()))
            fulltime_List.append(fulltime)
            self.fullTimeList.clear()
            fulltime_List.sort()
            for item in fulltime_List:
                self.fullTimeList.addItem(item)

        if str(self.employeeRank.currentText()) == "Operator":
            operator = (str(self.employee.text()))
            name_List.append(operator)
            self.operatorList.clear()
            name_List.sort()
            for item in name_List:
                self.operatorList.addItem(item)

        if str(self.employeeRank.currentText()) == "Vets":
            operator = (str(self.employee.text()))
            Trainer_list.append(operator)
            self.vetList.clear()
            Trainer_list.sort()
            for item in Trainer_list:
                self.vetList.addItem(item)



    def printSelection(self):
            print(self.operatorList.currentItem().text())
    def showDialog(self):
        global filename
        fname = QFileDialog.getOpenFileName(self, 'Open File Now', '/home')
        self.filename = fname[0]
        digit = self.filename.rfind('/')
        self.shortFile = self.filename[digit:]
        self.label.setText(self.shortFile)


        print(self.filename, digit, self.shortFile)
    def convertDates(self):
        worksheet.write(4, 1, '=TEXT(B113, "ddd - MM/DD")')
        worksheet.write(4, 2, '=TEXT(C113, "ddd - MM/DD")')
        worksheet.write(4, 3, '=TEXT(D113, "ddd - MM/DD")')
        worksheet.write(4, 4, '=TEXT(E113, "ddd - MM/DD")')
        worksheet.write(4, 5, '=TEXT(F113, "ddd - MM/DD")')
        worksheet.write(4, 6, '=TEXT(G113, "ddd - MM/DD")')
        worksheet.write(4, 7, '=TEXT(H113, "ddd - MM/DD")')

        worksheet2.write(4, 1, '=TEXT(B113, "ddd - MM/DD")')
        worksheet2.write(4, 2, '=TEXT(C113, "ddd - MM/DD")')
        worksheet2.write(4, 3, '=TEXT(D113, "ddd - MM/DD")')
        worksheet2.write(4, 4, '=TEXT(E113, "ddd - MM/DD")')
        worksheet2.write(4, 5, '=TEXT(F113, "ddd - MM/DD")')
        worksheet2.write(4, 6, '=TEXT(G113, "ddd - MM/DD")')
        worksheet2.write(4, 7, '=TEXT(H113, "ddd - MM/DD")')
    def figureDates(self):
        # workbook, worksheet, worksheet2, sheet, sheet2, format, format0, format1, format2, format3, format4, format5, format6, formatgreen, formatred, formatU = setFile()
        for row in range(sheet.nrows):
            if isinstance(sheet.cell_value(row, 1), float):
                date_List.append(sheet.cell_value(row, 1))
        dateListComplete.extend(set(date_List))
        dateListComplete.sort()
        print("datelist is ", dateListComplete)

    def unavailableDates(self):
        # workbook, worksheet, worksheet2, sheet, sheet2, format, format0, format1, format2, format3, format4, format5, format6, formatgreen, formatred, formatU = setFile()
        dates = []
        for col in range(sheet2.ncols):
            datesWithNames = []
            date = sheet2.cell_value(4, col)
            dates.append(date)

            for row in range(sheet2.nrows):
                if row > 4:
                    if sheet2.cell_value(row, col) != "":
                        datesWithNames.append([date, sheet2.cell_value(row, col)])

            print(datesWithNames)
            uDates.append(datesWithNames)

        print(dates)

    def schedule(self, x, y, yy, yyy, z, zz):
        # workbook, worksheet, worksheet2, sheet, sheet2, format, format0, format1, format2, format3, format4, format5, format6, formatgreen, formatred, formatU = setFile()
        rowz = 5
        rowzz = 5
        colz = 1
        colzz = 8

        # global sheet
        # global worksheet
        for date in x:
            notAvailable = []
            dlist = []
            duplicateCheck = []

            for row in range(sheet.nrows):
                if sheet.cell_value(row, 1) == date:
                    positionList = []
                    techName = sheet.cell_value(row, 8)
                    shiftHours = sheet.cell_value(row, 9).strip()
                    positionList = [sheet.cell_value(row, 9).strip(), str(sheet.cell_value(row, 2))]
                    techTime = {sheet.cell_value(row, 8).strip(): positionList}
                    namePerDay.update(techTime)
                    if "." in str(sheet.cell_value(row, 2)):

                        position = "Operator"
                    else:
                        position = str(sheet.cell_value(row, 2))
                    duplicateCheck.append([position, sheet.cell_value(row, 8).strip(), str(sheet.cell_value(row, 10))])


                    # 				print positionList

                    # 		namePerDayComplete = [dict(tupleized) for tupleized in set(tuple(item.items()) for item in namePerDay)]
                    # 		namePerDayComplete.sort()

            print(namePerDay, "THIS IS NAMEPERDAY")

            # 		namePerDayComplete.append(set(namePerDay))


            print(duplicateCheck, "duplicateCheck")
            dailyPositions = list(duplicateCheck for duplicateCheck, _ in itertools.groupby(duplicateCheck))
            print(dailyPositions, "DAILY POSITIONS")

            for item in dailyPositions:
                if item[1] not in dlist and item[1] != "":
                    dlist.append(item[1])
                else:
                    if item[1] != "":
                        dateConverted = xlrd.xldate.xldate_as_datetime(date, book.datemode)
                        double = "%s is double booked on %s" % (item[1], dateConverted)
                        print(double)

                        doubleBooked.append(double)

            print(uDates, "UDATES")

            for dateItem in uDates:
                print(dateItem, "DATEITEM")

                for object in dateItem:
                    print(object, "OBJECT")

                    if object[0] == date:
                        print(date, "Operator unavailable:", object[1])

                        notAvailable.append(object[1])



                        # 	print mondayConverted, "DATES"

            for name in z:

                rowformath = rowz + 1
                rowformath2 = rowzz + 1

                worksheet.write(rowz, 16, "=SUM(I%d:O%d)" % (rowformath, rowformath))
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '>',
                                                                    'value': 40,
                                                                    'format': formatred})
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '<=',
                                                                    'value': 40,
                                                                    'format': formatgreen})

                worksheet2.write(rowzz, 16, "=SUM(I%d:O%d)" % (rowformath2, rowformath2))
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '>',
                                                                      'value': 40,
                                                                      'format': formatred})
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '<=',
                                                                      'value': 40,
                                                                      'format': formatgreen})
                if name in namePerDay:
                    math = []

                    dailyHours = {}

                    if name in notAvailable:
                        print(name, date, "MISTAKE")

                        convertedDate = xlrd.xldate.xldate_as_datetime(date, book.datemode)
                        availableString = "%s is unavailable on: %s" % (name, convertedDate)
                        weekNotAvailable.append(availableString)

                    worksheet.write(rowz, colz, namePerDay[name][0], format1)
                    worksheet2.write(rowzz, colz, namePerDay[name][0], format1)

                    worksheet.write(rowz, 0, name)
                    worksheet2.write(rowzz, 0, name)

                    worksheet.set_row(rowz + 1, 4)

                    times = []
                    times.extend(namePerDay[name][0].split(" - "))
                    for time in times:
                        print(time)

                        if time in timeComparison:
                            print("TRUE ")

                            print(timeComparison[time])


                            math.append(timeComparison[time])
                        else:
                            print("FALSE")

                    timeTotal = math[1] - math[0]
                    print(timeTotal, "is today's hours")

                    dailyHours = {name: timeTotal}
                    adminTotalHours.append(dailyHours)

                    print(dailyHours)

                    worksheet.write(rowz, colzz, timeTotal)
                    worksheet2.write(rowzz, colzz, timeTotal)

                    print(times)



                else:
                    # 				print "Off", name
                    if name in notAvailable:
                        print("Unavailable", name)

                        worksheet.write(rowz, colz, "Unavailable", formatU)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "Unavailable", formatU)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                    else:

                        print("Off", name)

                        worksheet.write(rowz, colz, "OFF", format)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "OFF", format)
                        worksheet2.write(rowzz, 0, name)

                    worksheet.set_row(rowz + 1, 4)

                rowz += 2
                rowzz += 1

            # print " ---------------------------- \n"
            rowz += 1
            rowzz += 1

            for name in y:

                rowformath = rowz + 1
                rowformath2 = rowzz + 1

                worksheet2.write(rowzz, 16, "=SUM(I%d:O%d)" % (rowformath2, rowformath2))
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '>',
                                                                      'value': 40,
                                                                      'format': formatred})
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '<=',
                                                                      'value': 40,
                                                                      'format': formatgreen})

                worksheet.write(rowz, 16, "=SUM(I%d:O%d)" % (rowformath, rowformath))
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '>',
                                                                    'value': 40,
                                                                    'format': formatred})
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '<=',
                                                                    'value': 40,
                                                                    'format': formatgreen})

                math = []
                dailyHours = {}
                if name in namePerDay:

                    if name in notAvailable:
                        print(name, date, "MISTAKE")

                        convertedDate = xlrd.xldate.xldate_as_datetime(date, book.datemode)
                        availableString = "%s is unavailable on: %s" % (name, convertedDate)
                        weekNotAvailable.append(availableString)

                    if "Marker" in namePerDay[name][1] or "Cube" in namePerDay[name][1] or "Spooler" in namePerDay[name][1] or "Float" in namePerDay[name][1] or "Skipper" in namePerDay[name][
                        1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format3)
                        worksheet2.write(rowzz, colz, namePerDay[name][0], format3)

                        worksheet.write(rowz, 0, name)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                        print(name, namePerDay[name][0])

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:
                            print(time)


                            if time in timeComparison:
                                print("TRUE ")

                                print(timeComparison[time])


                                math.append(timeComparison[time])
                                print("This is math:", math)

                            else:
                                print("FALSE ")


                        timeTotal = math[1] - math[0]
                        print(timeTotal, "is today's hours")

                        dailyHours = {name: timeTotal}
                        print(dailyHours)

                        techTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                        print(times)



                    elif "EIC" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format2)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format2)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                        print(name, namePerDay[name][0])


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:
                            print(time)

                            if time in timeComparison:
                                print("TRUE ")

                                print(timeComparison[time])


                                math.append(timeComparison[time])
                                print("This is math:", math)

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]
                        print(timeTotal, "is today's hours")

                        dailyHours = {name: timeTotal}
                        print(dailyHours)

                        techTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                        print(times)




                    elif "DA" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format4)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format4)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                        print(name, namePerDay[name][0])


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:
                            print(time)

                            if time in timeComparison:
                                print("TRUE ")

                                print(timeComparison[time])


                                math.append(timeComparison[time])
                                print("This is math:", math)

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]
                        print(timeTotal, "is today's hours")

                        dailyHours = {name: timeTotal}
                        print(dailyHours)

                        techTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                        print(times)

                    elif "Utility" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format5)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format5)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format5)
                        worksheet2.write(rowzz, 0, name)

                        print(name, namePerDay[name][0])

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:
                            print(time)

                            if time in timeComparison:
                                print("TRUE ")

                                print(timeComparison[time])


                                math.append(timeComparison[time])
                                print("This is math:", math)

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]
                        print(timeTotal, "is today's hours")

                        dailyHours = {name: timeTotal}
                        print(dailyHours)

                        techTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                        print(times)


                    elif "Pre-Game Assist" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format6)
                        worksheet.write(rowz, 0, name)
                        worksheet.set_row(rowz + 1, 4)
                        print(name, namePerDay[name][0])


                        worksheet2.write(rowzz, colz, namePerDay[name][0], format6)
                        worksheet2.write(rowzz, 0, name)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:
                            print( time)

                            if time in timeComparison:
                                print("TRUE ")

                                print(timeComparison[time])


                                math.append(timeComparison[time])
                                print("This is math:", math)

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]
                        print(timeTotal, "is today's hours")

                        dailyHours = {name: timeTotal}
                        print(dailyHours)

                        techTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                        print(times)




                    else:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format0)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format0)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("False")
                        timeTotal = math[1] - math[0]


                        dailyHours = {name: timeTotal}

                        techTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)




                else:
                    if name in notAvailable:

                        worksheet.write(rowz, colz, "Unavailable", formatU)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "Unavailable", formatU)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                    else:


                        worksheet.write(rowz, colz, "OFF", format)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "OFF", format)
                        worksheet2.write(rowzz, 0, name)

                    worksheet.set_row(rowz + 1, 4)

                rowz += 2
                rowzz += 1

            # print " ---------------------------- \n"
            rowz += 1
            rowzz += 1

            for name in yy:

                rowformath = rowz + 1
                rowformath2 = rowzz + 1

                worksheet2.write(rowzz, 16, "=SUM(I%d:O%d)" % (rowformath2, rowformath2))
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '>',
                                                                      'value': 40,
                                                                      'format': formatred})
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '<=',
                                                                      'value': 40,
                                                                      'format': formatgreen})

                worksheet.write(rowz, 16, "=SUM(I%d:O%d)" % (rowformath, rowformath))
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '>',
                                                                    'value': 40,
                                                                    'format': formatred})
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '<=',
                                                                    'value': 40,
                                                                    'format': formatgreen})

                math = []
                dailyHours = {}
                if name in namePerDay:

                    if name in notAvailable:

                        convertedDate = xlrd.xldate.xldate_as_datetime(date, book.datemode)
                        availableString = "%s is unavailable on: %s" % (name, convertedDate)
                        weekNotAvailable.append(availableString)

                    if "Marker" in namePerDay[name][1] or "Cube" in namePerDay[name][1] or "Spooler" in namePerDay[name][1] or "Float" in namePerDay[name][1] or "Skipper" in namePerDay[name][
                        1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format3)
                        worksheet2.write(rowzz, colz, namePerDay[name][0], format3)

                        worksheet.write(rowz, 0, name)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:


                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")


                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        backWallTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)




                    elif "EIC" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format2)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format2)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        backWallTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)





                    elif "DA" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format4)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format4)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        backWallTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)


                    elif "Utility" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format5)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format5)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format5)
                        worksheet2.write(rowzz, 0, name)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        backWallTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)



                    elif "Pre-Game Assist" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format6)
                        worksheet.write(rowz, 0, name)
                        worksheet.set_row(rowz + 1, 4)


                        worksheet2.write(rowzz, colz, namePerDay[name][0], format6)
                        worksheet2.write(rowzz, 0, name)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        backWallTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)





                    else:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format0)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format0)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}


                        backWallTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)





                else:
                    if name in notAvailable:

                        worksheet.write(rowz, colz, "Unavailable", formatU)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "Unavailable", formatU)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                    else:


                        worksheet.write(rowz, colz, "OFF", format)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "OFF", format)
                        worksheet2.write(rowzz, 0, name)

                    worksheet.set_row(rowz + 1, 4)

                rowz += 2
                rowzz += 1

            # print " ---------------------------- \n"
            rowz += 1
            rowzz += 1

            for name in yyy:

                rowformath = rowz + 1
                rowformath2 = rowzz + 1

                worksheet2.write(rowzz, 16, "=SUM(I%d:O%d)" % (rowformath2, rowformath2))
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '>',
                                                                      'value': 40,
                                                                      'format': formatred})
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '<=',
                                                                      'value': 40,
                                                                      'format': formatgreen})

                worksheet.write(rowz, 16, "=SUM(I%d:O%d)" % (rowformath, rowformath))
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '>',
                                                                    'value': 40,
                                                                    'format': formatred})
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '<=',
                                                                    'value': 40,
                                                                    'format': formatgreen})

                math = []
                dailyHours = {}
                if name in namePerDay:

                    if name in notAvailable:

                        convertedDate = xlrd.xldate.xldate_as_datetime(date, book.datemode)
                        availableString = "%s is unavailable on: %s" % (name, convertedDate)
                        weekNotAvailable.append(availableString)

                    if "Marker" in namePerDay[name][1] or "Cube" in namePerDay[name][1] or "Spooler" in namePerDay[name][1] or "Float" in namePerDay[name][1] or "Skipper" in namePerDay[name][
                        1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format3)
                        worksheet2.write(rowzz, colz, namePerDay[name][0], format3)

                        worksheet.write(rowz, 0, name)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:


                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")


                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}


                        fullTimeTotalHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)



                    elif "EIC" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format2)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format2)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        fullTimeTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)





                    elif "DA" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format4)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format4)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        fullTimeTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)


                    elif "Utility" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format5)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format5)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format5)
                        worksheet2.write(rowzz, 0, name)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        fullTimeTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)



                    elif "Pre-Game Assist" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format6)
                        worksheet.write(rowz, 0, name)
                        worksheet.set_row(rowz + 1, 4)


                        worksheet2.write(rowzz, colz, namePerDay[name][0], format6)
                        worksheet2.write(rowzz, 0, name)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        fullTimeTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)





                    else:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format0)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format0)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        fullTimeTotalHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)





                else:
                    if name in notAvailable:

                        worksheet.write(rowz, colz, "Unavailable", formatU)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "Unavailable", formatU)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                    else:


                        worksheet.write(rowz, colz, "OFF", format)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "OFF", format)
                        worksheet2.write(rowzz, 0, name)

                    worksheet.set_row(rowz + 1, 4)

                rowz += 2
                rowzz += 1

            # print " ---------------------------- \n"
            rowz += 1
            rowzz += 1
            for name in zz:
                print(name, "This is the name")
                rowformath = rowz + 1
                rowformath2 = rowzz + 1

                worksheet2.write(rowzz, 16, "=SUM(I%d:O%d)" % (rowformath2, rowformath2))
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '>',
                                                                      'value': 40,
                                                                      'format': formatred})
                worksheet2.conditional_format("Q%d" % (rowformath2), {'type': 'cell',
                                                                      'criteria': '<=',
                                                                      'value': 40,
                                                                      'format': formatgreen})

                worksheet.write(rowz, 16, "=SUM(I%d:O%d)" % (rowformath, rowformath))
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '>',
                                                                    'value': 40,
                                                                    'format': formatred})
                worksheet.conditional_format("Q%d" % (rowformath), {'type': 'cell',
                                                                    'criteria': '<=',
                                                                    'value': 40,
                                                                    'format': formatgreen})

                math = []
                dailyHours = {}

                if name in namePerDay:
                    print(name, namePerDay[name][0], namePerDay[name][1], "look here")
                    if name in notAvailable:

                        convertedDate = xlrd.xldate.xldate_as_datetime(date, book.datemode)
                        availableString = "%s is unavailable on: %s" % (name, convertedDate)
                        weekNotAvailable.append(availableString)
                    if "EIC" in namePerDay[name][1]:
                        # 	print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format2)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format2)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)


                    elif "Marker" in namePerDay[name][1] or "Cube" in namePerDay[name][1] or "Spooler" in namePerDay[name][1] or "Float" in namePerDay[name][1] or "Skipper" in namePerDay[name][
                        1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format3)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format3)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)

                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)




                    elif "DA" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format4)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format4)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                    elif "INFRA" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], formatInfra)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], formatInfra)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)

                    elif "MOD" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format7)
                        worksheet.write(rowz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format7)
                        worksheet2.write(rowzz, 0, name)
                        #

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)



                    elif "Pre-Game Assist" in namePerDay[name][1]:
                        # print sheet.cell_value(row,3)
                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format6)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format6)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)


                        # worksheet2.write(rowzz, colz, namePerDay[name][0], format6)
                        # 					worksheet2.write(rowzz, 0, name)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)



                    else:


                        # 					print namePerDay[name], name
                        worksheet.write(rowz, colz, namePerDay[name][0], format0)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, namePerDay[name][0], format0)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)

                        times = []
                        times.extend(namePerDay[name][0].split(" - "))
                        for time in times:

                            if time in timeComparison:


                                math.append(timeComparison[time])

                            else:
                                print("FALSE ")

                        timeTotal = math[1] - math[0]

                        dailyHours = {name: timeTotal}

                        realFullTimeHours.append(dailyHours)
                        worksheet.write(rowz, colzz, timeTotal)
                        worksheet2.write(rowzz, colzz, timeTotal)




                else:
                    if name in notAvailable:

                        worksheet.write(rowz, colz, "Unavailable", formatU)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "Unavailable", formatU)
                        worksheet2.write(rowzz, 0, name)

                        worksheet.set_row(rowz + 1, 4)
                    else:


                        worksheet.write(rowz, colz, "OFF", format)
                        worksheet.write(rowz, 0, name)

                        worksheet2.write(rowzz, colz, "OFF", format)
                        worksheet2.write(rowzz, 0, name)

                    worksheet.set_row(rowz + 1, 4)

                rowz += 2
                rowzz += 1


            # print " ---------------------------- \n"
            # 		del namePerDayComplete [ : ]
            worksheet.write(4, colz, date)
            worksheet2.write(4, colz, date)

            worksheet.write(112, colz, date)
            worksheet2.write(112, colz, date)
            colz += 1
            rowz = 5
            rowzz = 5

            colzz += 1
            namePerDay.clear()




    def runProgram(self):
        print(self.shortFile)

        ##SetCREW




        ## SETDATA

        global adminTotalHours
        adminTotalHours = []
        global techTotalHours
        techTotalHours = []
        global backWallTotalHours
        backWallTotalHours = []
        global fullTimeTotalHours
        fullTimeTotalHours = []
        global realFullTimeHours
        realFullTimeHours = []

        global namePerDay
        namePerDay = {}
        global namePerDayComplete
        namePerDayComplete = []
        global timePerDay
        timePerDay = []
        global date_List
        date_List = []
        global dateListComplete
        dateListComplete = []
        global doubleBooked
        doubleBooked = []
        global weeklyHours
        weeklyHours = {}
        global weekNotAvailable
        weekNotAvailable = []
        global timeComparison
        timeComparison = {"1:00a": 25,
                          "1:30a": 25.5,
                          "2:00a": 26,
                          "2:30a": 26.5,
                          "3:00a": 27,
                          "3:30a": 27.5,
                          "4:00a": 4,
                          "5:00a": 5,
                          "6:00a": 6,
                          "7:00a": 7,
                          "7:30a": 7.5,
                          "8:00a": 8,
                          "8:30a": 8.5,
                          "9:00a": 9,
                          "9:30a": 9.5,
                          "10:00a": 10,
                          "10:30a": 10.5,
                          "11:00a": 11,
                          "11:30a": 11.5,
                          "12:00p": 12,
                          "12:30p": 12.5,
                          "1:00p": 13,
                          "1:30p": 13.5,
                          "2:00p": 14,
                          "2:30p": 14.5,
                          "3:00p": 15,
                          "3:30p": 15.5,
                          "4:00p": 16,
                          "4:30p": 16.5,
                          "5:00p": 17,
                          "5:30p": 17.5,
                          "6:00p": 18,
                          "6:30p": 18.5,
                          "7:00p": 19,
                          "7:30p": 19.5,
                          "8:00p": 20,
                          "8:30p": 20.5,
                          "9:00p": 21,
                          "9:30p": 21.5,
                          "10:00p": 22,
                          "10:30p": 22.5,
                          "11:00p": 23,
                          "11:30p": 23.5,
                          "12:00a": 24,
                          "12:30a": 24.5}

        splitFile = self.filename.split("/")

        # print splitFile

        fileTitle = splitFile[-1].split(".")
        # print fileTitle

        insertTitle = fileTitle[0]


        global book
        book = xlrd.open_workbook(self.filename)

        savePath = os.path.expanduser("~/Desktop/" + insertTitle + "_b.xlsx")
        workbook = xlsxwriter.Workbook(savePath)

        print(workbook, book)

        global worksheet
        global worksheet2

        worksheet = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet.set_column(0, 7, 14)
        worksheet2.set_column(0, 7, 14)
        worksheet.set_row(112, None, None, {'hidden': True})
        worksheet2.set_row(112, None, None, {'hidden': True})

        global sheet
        global sheet2
        sheet = book.sheet_by_index(0)
        sheet2 = book.sheet_by_index(1)

        global format
        global format0
        global format1
        global format2
        global format3
        global format4
        global format5
        global format6
        global format7
        global formatgreen
        global formatred
        global formatU
        global formatInfra

        format = workbook.add_format()
        format0 = workbook.add_format()
        format1 = workbook.add_format()
        format2 = workbook.add_format()
        format3 = workbook.add_format()
        format4 = workbook.add_format()
        format5 = workbook.add_format()
        format6 = workbook.add_format()
        format7 = workbook.add_format()


        formatgreen = workbook.add_format()
        formatred = workbook.add_format()
        formatU = workbook.add_format()
        formatInfra = workbook.add_format()

        format.set_bg_color('#D8DCDC')
        format.set_border(1)
        format.set_align('center')

        format0.set_border(1)
        format0.set_align('center')

        format1.set_bg_color('#FCE4D6')
        format1.set_border(1)
        format1.set_align('center')

        format2.set_bg_color('#C5E0B3')
        format2.set_border(1)
        format2.set_align('center')

        format3.set_bg_color('#FF6B7B')
        format3.set_border(1)
        format3.set_align('center')

        # Trainnee
        format4.set_bg_color('#B4C6E7')
        format4.set_border(1)
        format4.set_align('center')

        # Marker  - Utility
        format5.set_bg_color('#F4B084')
        format5.set_border(1)
        format5.set_align('center')

        format6.set_bg_color('#A6A6A6')
        format6.set_border(1)
        format6.set_align('center')

        format7.set_bg_color('#A9D08E')
        format7.set_border(1)
        format7.set_align('center')

        # # Marker
        # format8.set_bg_color('#DA9694')
        # format8.set_border(1)
        # format8.set_align('center')






        formatred.set_bg_color('#ff0000')
        formatred.set_border(1)
        formatred.set_align('center')

        formatgreen.set_bg_color('#BCED91')
        formatgreen.set_border(1)
        formatgreen.set_align('center')

        formatU.set_bg_color('#808080')
        formatU.set_border(1)
        formatU.set_align('center')

        formatInfra.set_bg_color("#FFC000")
        formatInfra.set_border(1)
        formatInfra.set_align('center')

        weekOf = sheet.cell_value(1, 0)
        actualWeek = sheet.cell_value(1, 1)
        departmentName = sheet.cell_value(2, 0)
        actualDepName = sheet.cell_value(2, 1)
        print(weekOf, actualWeek, departmentName, actualDepName)

        worksheet.write(1, 0, weekOf)
        worksheet.write(1, 1, actualWeek)
        worksheet.write(2, 0, departmentName)
        worksheet.write(2, 1, actualDepName)

        worksheet2.write(1, 0, weekOf)
        worksheet2.write(1, 1, actualWeek)
        worksheet2.write(2, 0, departmentName)
        worksheet2.write(2, 1, actualDepName)

        try:
            self.figureDates()
            self.unavailableDates()
            self.schedule(dateListComplete, name_List, downstairs_List, Trainer_list, admin_List, fulltime_List)
            self.convertDates()
            workbook.close()
        except Exception as e:
            self.label.setText("ERROR!")
            print(e)
            if e == "list index out of range":
                self.unavailableBooked.addItem("Check your Shift Times!!!")


        uString = ""
        doubleString = ""
        self.doubleBookedList.clear()
        self.unavailableBooked.clear()
        for item in doubleBooked:
            doubleString += item + "\n"
            self.doubleBookedList.addItem(item)

        for object in weekNotAvailable:
            uString += object + "\n"
            self.unavailableBooked.addItem(object)
        print(doubleBooked)
        print(weekNotAvailable)
        self.declareHours()
        self.turnAround(dateListComplete)





    def declareHours(self):
        employeeHours = []
        adminhours = Counter()
        operatorhours = Counter()
        backwallhours = Counter()
        fulltimehours = Counter()
        realtimefulltime = Counter()

        for object in adminTotalHours:
            adminhours.update(object)
        for object in techTotalHours:
            operatorhours.update(object)
        for object in backWallTotalHours:
            backwallhours.update(object)
        for object in fullTimeTotalHours:
            fulltimehours.update(object)
        for object in realFullTimeHours:
            realtimefulltime.update(object)
        print(adminhours)
        print(realFullTimeHours, "FULLTIMEHOURS")


        for object in adminhours:
            hours = adminhours.get(object)
            employeehourObject = str(object) + ' ' + str(hours)
            employeeHours.append(employeehourObject)
        for object in operatorhours:
            hours = operatorhours.get(object)
            employeehourObject = str(object) + ' ' + str(hours)
            employeeHours.append(employeehourObject)

        for object in backwallhours:
            hours = backwallhours.get(object)
            employeehourObject = str(object) + ' ' + str(hours)
            employeeHours.append(employeehourObject)

        for object in realtimefulltime:
            hours = realtimefulltime.get(object)
            employeehourObject = str(object) + ' ' + str(hours)
            employeeHours.append(employeehourObject)

        for object in fulltimehours:
            hours = fulltimehours.get(object)
            employeehourObject = str(object) + ' ' + str(hours)
            employeeHours.append(employeehourObject)

        employeeHours.sort()
        print(employeeHours, "EMPLOYEEHOURS")
        self.hoursList.clear()
        for item in employeeHours:
            self.hoursList.addItem(item)


    def turnAround(self, x):
        shortTurnAroundList = []

        timeForTurnAround = {"1:00a": 1,
                          "1:30a": 1.5,
                          "2:00a": 2,
                          "2:30a": 2.5,
                          "3:00a": 3,
                          "3:30a": 3.5,
                          "4:00a": 4,
                          "5:00a": 5,
                          "6:00a": 6,
                          "7:00a": 7,
                          "7:30a": 7.5,
                          "8:00a": 8,
                          "8:30a": 8.5,
                          "9:00a": 9,
                          "9:30a": 9.5,
                          "10:00a": 10,
                          "10:30a": 10.5,
                          "11:00a": 11,
                          "11:30a": 11.5,
                          "12:00p": 12,
                          "12:30p": 12.5,
                          "1:00p": 13,
                          "1:30p": 13.5,
                          "2:00p": 14,
                          "2:30p": 14.5,
                          "3:00p": 15,
                          "3:30p": 15.5,
                          "4:00p": 16,
                          "4:30p": 16.5,
                          "5:00p": 17,
                          "5:30p": 17.5,
                          "6:00p": 18,
                          "6:30p": 18.5,
                          "7:00p": 19,
                          "7:30p": 19.5,
                          "8:00p": 20,
                          "8:30p": 20.5,
                          "9:00p": 21,
                          "9:30p": 21.5,
                          "10:00p": 22,
                          "10:30p": 22.5,
                          "11:00p": 23,
                          "11:30p": 23.5,
                          "12:00a": 0,
                          "12:30a": 0.5}

        day = 0
        mondaySchedule = []
        tuesdaySchedule = []
        wednesdaySchedule = []
        thursdaySchedule = []
        fridaySchedule = []
        saturdaySchedule = []
        sundaySchedule = []

        mondayScheduleFinal = []
        tuesdayScheduleFinal = []
        wednesdayScheduleFinal = []
        thursdayScheduleFinal = []
        fridayScheduleFinal = []
        saturdayScheduleFinal = []
        sundayScheduleFinal = []
        for date in x:
            day += 1
            for row in range(sheet.nrows):
                operator = sheet.cell_value(row, 8)
                shiftTime = sheet.cell_value(row,9)
                if sheet.cell_value(row, 1) == date:
                    nameAndTime = {"operator": operator, "shiftTime" : shiftTime}
                    if day == 1:
                        mondaySchedule.append(nameAndTime)
                    if day == 2:
                        tuesdaySchedule.append(nameAndTime)
                    if day == 3:
                        wednesdaySchedule.append(nameAndTime)
                    if day == 4:
                        thursdaySchedule.append(nameAndTime)
                    if day == 5:
                        fridaySchedule.append(nameAndTime)
                    if day == 6:
                        saturdaySchedule.append(nameAndTime)
                    if day == 7:
                        sundaySchedule.append(nameAndTime)




            for i in mondaySchedule:
                if i not in mondayScheduleFinal:
                    mondayScheduleFinal.append(i)

            for i in tuesdaySchedule:
                if i not in tuesdayScheduleFinal:
                    tuesdayScheduleFinal.append(i)

            for i in wednesdaySchedule:
                if i not in wednesdayScheduleFinal:
                    wednesdayScheduleFinal.append(i)

            for i in thursdaySchedule:
                if i not in thursdayScheduleFinal:
                    thursdayScheduleFinal.append(i)

            for i in fridaySchedule:
                if i not in fridayScheduleFinal:
                    fridayScheduleFinal.append(i)

            for i in saturdaySchedule:
                if i not in saturdayScheduleFinal:
                    saturdayScheduleFinal.append(i)

            for i in sundaySchedule:
                if i not in sundayScheduleFinal:
                    sundayScheduleFinal.append(i)


        print(mondayScheduleFinal)
        print(tuesdayScheduleFinal)
        print(wednesdayScheduleFinal)
        print(thursdayScheduleFinal)
        print(fridayScheduleFinal)

        for employee in mondayScheduleFinal:
            checkEmployee = employee['operator']
            for object in tuesdayScheduleFinal:
                if object['operator'] == checkEmployee:
                    print(checkEmployee, 'worked 2 days in a row')
                    outTime = employee['shiftTime'].split(' - ')[1]
                    inTime = object['shiftTime'].split(' - ')[0]
                    outTime = outTime.strip()
                    inTime = inTime.strip()
                    print(outTime, inTime)

                    if timeForTurnAround[outTime] > 12:
                        turnAroundTime = 24 - timeForTurnAround[outTime] + timeForTurnAround[inTime]

                    else:
                        turnAroundTime = timeForTurnAround[inTime] - timeForTurnAround[outTime]


                    print(checkEmployee, turnAroundTime)

                    if turnAroundTime < 12:
                        shortTurnAround = checkEmployee + " has a turn around time of " + str(turnAroundTime) + " on Monday night."
                        shortTurnAroundList.append(shortTurnAround)



        for employee in tuesdayScheduleFinal:
            checkEmployee = employee['operator']
            for object in wednesdayScheduleFinal:
                if object['operator'] == checkEmployee:
                    print(checkEmployee, 'worked 2 days in a row')
                    outTime = employee['shiftTime'].split(' - ')[1]
                    inTime = object['shiftTime'].split(' - ')[0]
                    outTime = outTime.strip()
                    inTime = inTime.strip()
                    print(outTime, inTime)

                    if timeForTurnAround[outTime] > 12:
                        turnAroundTime = 24 - timeForTurnAround[outTime] + timeForTurnAround[inTime]

                    else:
                        turnAroundTime = timeForTurnAround[inTime] - timeForTurnAround[outTime]


                    print(checkEmployee, turnAroundTime)

                    if turnAroundTime < 12:
                        shortTurnAround = checkEmployee + " has a turn around time of " + str(turnAroundTime) + " on Tuesday night."
                        shortTurnAroundList.append(shortTurnAround)

        for employee in wednesdayScheduleFinal:
            checkEmployee = employee['operator']
            for object in thursdayScheduleFinal:
                if object['operator'] == checkEmployee:
                    print(checkEmployee, 'worked 2 days in a row')
                    outTime = employee['shiftTime'].split(' - ')[1]
                    inTime = object['shiftTime'].split(' - ')[0]
                    outTime = outTime.strip()
                    inTime = inTime.strip()
                    print(outTime, inTime)

                    if timeForTurnAround[outTime] > 12:
                        turnAroundTime = 24 - timeForTurnAround[outTime] + timeForTurnAround[inTime]

                    else:
                        turnAroundTime = timeForTurnAround[inTime] - timeForTurnAround[outTime]


                    print(checkEmployee, turnAroundTime)

                    if turnAroundTime < 12:
                        shortTurnAround = checkEmployee + " has a turn around time of " + str(turnAroundTime) + " on Wednesday night."
                        shortTurnAroundList.append(shortTurnAround)

        for employee in thursdayScheduleFinal:
            checkEmployee = employee['operator']
            for object in fridayScheduleFinal:
                if object['operator'] == checkEmployee:
                    print(checkEmployee, 'worked 2 days in a row')
                    outTime = employee['shiftTime'].split(' - ')[1]
                    inTime = object['shiftTime'].split(' - ')[0]
                    outTime = outTime.strip()
                    inTime = inTime.strip()
                    print(outTime, inTime)

                    if timeForTurnAround[outTime] > 12:
                        turnAroundTime = 24 - timeForTurnAround[outTime] + timeForTurnAround[inTime]

                    else:
                        turnAroundTime = timeForTurnAround[inTime] - timeForTurnAround[outTime]


                    print(checkEmployee, turnAroundTime)

                    if turnAroundTime < 12:
                        shortTurnAround = checkEmployee + " has a turn around time of " + str(turnAroundTime) + " on Thursday night."
                        shortTurnAroundList.append(shortTurnAround)

        for employee in fridayScheduleFinal:
            checkEmployee = employee['operator']
            for object in saturdayScheduleFinal:
                if object['operator'] == checkEmployee:
                    print(checkEmployee, 'worked 2 days in a row')
                    outTime = employee['shiftTime'].split(' - ')[1]
                    inTime = object['shiftTime'].split(' - ')[0]
                    outTime = outTime.strip()
                    inTime = inTime.strip()
                    print(outTime, inTime)

                    if timeForTurnAround[outTime] > 12:
                        turnAroundTime = 24 - timeForTurnAround[outTime] + timeForTurnAround[inTime]

                    else:
                        turnAroundTime = timeForTurnAround[inTime] - timeForTurnAround[outTime]


                    print(checkEmployee, turnAroundTime)

                    if turnAroundTime < 12:
                        shortTurnAround = checkEmployee + " has a turn around time of " + str(turnAroundTime) + " on Friday night."
                        shortTurnAroundList.append(shortTurnAround)

        for employee in saturdayScheduleFinal:
            checkEmployee = employee['operator']
            for object in sundayScheduleFinal:
                if object['operator'] == checkEmployee:
                    print(checkEmployee, 'worked 2 days in a row')
                    outTime = employee['shiftTime'].split(' - ')[1]
                    inTime = object['shiftTime'].split(' - ')[0]
                    outTime = outTime.strip()
                    inTime = inTime.strip()
                    print(outTime, inTime)

                    if timeForTurnAround[outTime] > 12:
                        turnAroundTime = 24 - timeForTurnAround[outTime] + timeForTurnAround[inTime]

                    else:
                        turnAroundTime = timeForTurnAround[inTime] - timeForTurnAround[outTime]


                    print(checkEmployee, turnAroundTime)

                    if turnAroundTime < 12:
                        shortTurnAround = checkEmployee + " has a turn around time of " + str(turnAroundTime) + " on Saturday night."
                        shortTurnAroundList.append(shortTurnAround)

        print(shortTurnAroundList)
        self.shortTurnAround.clear()
        for item in shortTurnAroundList:
            self.shortTurnAround.addItem(item)

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())