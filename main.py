import os,sys,csv,datetime
if hasattr(sys,'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ';' + os.environ['PATH']
from Email import Ui_Email
from GUI import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *



class mainForm(QMainWindow,Ui_MainWindow):
    def __init__(self, *args, **kwargs):
        QMainWindow.__init__(self, *args, **kwargs)
        self.setupUi(self)
        self.initUI()

    def initUI(self):
        self.openFile.clicked.connect(self.singleBrowse)
        self.openFile_2.clicked.connect(self.singleBrowse)
        self.openFile_3.clicked.connect(self.singleBrowse)
        self.openFile_4.clicked.connect(self.singleBrowse)
        self.openFile_5.clicked.connect(self.singleBrowse)
        self.openFile_6.clicked.connect(self.singleBrowse)
        self.openFile_7.clicked.connect(self.singleBrowse)

        self.createNewLoan.clicked.connect(self.createNewRow)
        self.createNewGift.clicked.connect(self.createNewRow)
        self.createNewLoan_2.clicked.connect(self.createNewRow)
        self.createNewLoan_3.clicked.connect(self.createNewRow)
        self.createNewLoan_4.clicked.connect(self.createNewRow)
        self.createNewGift_2.clicked.connect(self.createNewRow)
        self.createNewInst.clicked.connect(self.createNewRow)

        self.createNewColumn_1.clicked.connect(self.createNewColumn)
        self.createNewColumn_2.clicked.connect(self.createNewColumn)
        self.createNewColumn_3.clicked.connect(self.createNewColumn)
        self.createNewColumn_4.clicked.connect(self.createNewColumn)
        self.createNewColumn_5.clicked.connect(self.createNewColumn)
        self.createNewColumn_6.clicked.connect(self.createNewColumn)
        self.createNewColumn_7.clicked.connect(self.createNewColumn)


        self.exit.clicked.connect(self.close)
        self.exit_2.clicked.connect(self.close)
        self.exit_3.clicked.connect(self.close)
        self.exit_4.clicked.connect(self.close)
        self.exit_5.clicked.connect(self.close)
        self.exit_6.clicked.connect(self.close)
        self.exit_7.clicked.connect(self.close)

        self.inLoanTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)
        self.inGiftTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)
        self.inLoanReqTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)
        self.outLoanTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)
        self.outLoanReqTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)
        self.outGiftTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)
        self.instTable.horizontalHeader().sectionDoubleClicked.connect(self.changeHeader)

        self.search.clicked.connect(self.findNmber)
        self.search_2.clicked.connect(self.findNmber)
        self.search_3.clicked.connect(self.findNmber)
        self.search_4.clicked.connect(self.findNmber)
        self.search_5.clicked.connect(self.findNmber)
        self.search_6.clicked.connect(self.findNmber)
        self.search_7.clicked.connect(self.findNmber)

        self.generateNum.clicked.connect(self.createNum)
        self.generateNum_2.clicked.connect(self.createNum)
        self.generateNum_3.clicked.connect(self.createNum)
        self.generateNum_4.clicked.connect(self.createNum)
        self.generateNum_5.clicked.connect(self.createNum)
        self.generateNum_6.clicked.connect(self.createNum)

        self.autoFillInLoan.clicked.connect(self.autoFill)
        self.autoFillInGift.clicked.connect(self.autoFill)
        #self.autoFillInLoanR.clicked.connect(self.autoFill)
        self.autoFillOutGift.clicked.connect(self.autoFill)
        self.autoFillOutLoan.clicked.connect(self.autoFill)
        #self.autoFillOutLoanR.clicked.connect(self.autoFill)

        self.save.clicked.connect(self.saveFile)
        self.save_2.clicked.connect(self.saveFile)
        self.save_3.clicked.connect(self.saveFile)
        self.save_4.clicked.connect(self.saveFile)
        self.save_5.clicked.connect(self.saveFile)
        self.save_6.clicked.connect(self.saveFile)
        self.save_7.clicked.connect(self.saveFile)

        self.clearC_1.clicked.connect(self.clearColour)
        self.clearC_2.clicked.connect(self.clearColour)
        self.clearC_3.clicked.connect(self.clearColour)
        self.clearC_4.clicked.connect(self.clearColour)
        self.clearC_5.clicked.connect(self.clearColour)
        self.clearC_6.clicked.connect(self.clearColour)
        self.clearC_7.clicked.connect(self.clearColour)

        self.sendMailInLoan.clicked.connect(self.openEmial)
        self.sendEmailOutLoans.clicked.connect(self.openEmial)


    def singleBrowse(self):    # This is the function to open a CSV file and browse it

        if self.tabWidget.currentIndex() is 0:
            self.singleBrowseCore(self.inLoanTable)

        elif self.tabWidget.currentIndex() is 1:
            self.singleBrowseCore(self.inGiftTable)

        elif self.tabWidget.currentIndex() is 2:
            self.singleBrowseCore(self.inLoanReqTable)

        elif self.tabWidget.currentIndex() is 3:
            self.singleBrowseCore(self.outLoanTable)

        elif self.tabWidget.currentIndex() is 4:
            self.singleBrowseCore(self.outLoanReqTable)

        elif self.tabWidget.currentIndex() is 5:
            self.singleBrowseCore(self.outGiftTable)

        elif self.tabWidget.currentIndex() is 6:
            self.singleBrowseCore(self.instTable)


    def singleBrowseCore(self,file:QTableWidget):
        self.filePath = QFileDialog.getOpenFileName(self,'Single File','','*.csv')
        if self.filePath[0] != '':
           with open(self.filePath[0],"r",encoding ='mac_roman') as data:
               reader = csv.reader(data, delimiter=',')
               self.tem = list(reader)
           file.setRowCount(0)
           file.setColumnCount(len(self.tem[0]))
           colHeaders = self.tem[0]
           file.setHorizontalHeaderLabels(colHeaders)
           for rowData in self.tem[1:]:  # populating data from local CSV file
                row = file.rowCount()
                file.insertRow(row)
                for column, stuff in enumerate(rowData):
                    item = QTableWidgetItem(stuff)
                    file.setItem(row,column,item)

                header = file.horizontalHeader()
                file.horizontalHeader().setStretchLastSection(True)
                file.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
                for i in range(len(self.tem[0])):
                    header.setSectionResizeMode(int(i),QtWidgets.QHeaderView.ResizeToContents)
                file.resizeRowsToContents()


    def createNewRow(self):   # This is the function to create a new row inside a table
        if self.tabWidget.currentIndex() is 0 and self.inLoanTable.columnCount() != 0:
            rowPosition = self.inLoanTable.rowCount()
            self.inLoanTable.setRowCount(rowPosition + 1)
        elif self.tabWidget.currentIndex() is 1 and self.inGiftTable.columnCount() != 0:
            rowPosition = self.inGiftTable.rowCount()
            self.inGiftTable.setRowCount(rowPosition + 1)
        elif self.tabWidget.currentIndex() is 2 and self.inLoanReqTable.columnCount() != 0:
            rowPosition = self.inLoanReqTable.rowCount()
            self.inLoanReqTable.setRowCount(rowPosition + 1)
        elif self.tabWidget.currentIndex() is 3 and self.outLoanTable.columnCount() != 0:
            rowPosition = self.outLoanTable.rowCount()
            self.outLoanTable.setRowCount(rowPosition + 1)
        elif self.tabWidget.currentIndex() is 4 and self.outLoanReqTable.columnCount() != 0:
            rowPosition = self.outLoanReqTable.rowCount()
            self.outLoanReqTable.setRowCount(rowPosition + 1)
        elif self.tabWidget.currentIndex() is 5 and self.outGiftTable.columnCount() != 0:
             rowPosition = self.outGiftTable.rowCount()
             self.outGiftTable.setRowCount(rowPosition + 1)
        elif self.tabWidget.currentIndex() is 6 and self.instTable.columnCount() != 0:
             rowPosition = self.instTable.rowCount()
             self.instTable.setRowCount(rowPosition + 1)
        else:
            pass


    def createNewColumn(self):  # This is the function to create a new row inside a table
        if self.tabWidget.currentIndex() is 0 and self.inLoanTable.rowCount() != 0:
            colPosition = self.inLoanTable.columnCount()
            self.inLoanTable.insertColumn(colPosition)
            self.inLoanTable.resizeColumnsToContents()
        elif self.tabWidget.currentIndex() is 1 and self.inGiftTable.rowCount() != 0:
            colPosition = self.inGiftTable.columnCount()
            self.inGiftTable.insertColumn(colPosition)
            self.inGiftTable.resizeColumnsToContents()
        elif self.tabWidget.currentIndex() is 2 and self.inLoanReqTable.rowCount() != 0:
            colPosition = self.inLoanReqTable.columnCount()
            self.inLoanReqTable.insertColumn(colPosition)
            self.inLoanReqTable.resizeColumnsToContents()
        elif self.tabWidget.currentIndex() is 3 and self.outLoanTable.rowCount() != 0:
            colPosition = self.outLoanTable.columnCount()
            self.outLoanTable.insertColumn(colPosition)
            self.outLoanTable.resizeColumnsToContents()
        elif self.tabWidget.currentIndex() is 4 and self.outLoanReqTable.rowCount() != 0:
            colPosition = self.outLoanReqTable.columnCount()
            self.outLoanReqTable.insertColumn(colPosition)
            self.outLoanReqTable.resizeColumnsToContents()
        elif self.tabWidget.currentIndex() is 5 and self.outGiftTable.rowCount() != 0:
             colPosition = self.outGiftTable.columnCount()
             self.outGiftTable.insertColumn(colPosition)
             self.outGiftTable.resizeColumnsToContents()
        elif self.tabWidget.currentIndex() is 6 and self.instTable.rowCount() != 0:
             colPosition = self.instTable.columnCount()
             self.instTable.insertColumn(colPosition)
             self.instTable.resizeColumnsToContents()
        else:
            pass


    def keyPressEvent(self, event):  # Single click to remove the whole row/column(only one row/colum at a time)
        if event.key() == Qt.Key_Backspace and len(self.inLoanTable.selectionModel().selectedRows()) == 1 \
                or event.key() == Qt.Key_Backspace and len(self.inGiftTable.selectionModel().selectedRows()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.inLoanReqTable.selectionModel().selectedRows()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.outLoanTable.selectionModel().selectedRows()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.outLoanReqTable.selectionModel().selectedRows()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.outGiftTable.selectionModel().selectedRows()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.instTable.selectionModel().selectedRows()) == 1:
            self.removeRow()

        elif event.key() == Qt.Key_Backspace and len(self.inLoanTable.selectionModel().selectedColumns()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.inGiftTable.selectionModel().selectedColumns()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.inLoanReqTable.selectionModel().selectedColumns()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.outLoanTable.selectionModel().selectedColumns()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.outLoanReqTable.selectionModel().selectedColumns()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.outGiftTable.selectionModel().selectedColumns()) == 1\
                or event.key() == Qt.Key_Backspace and len(self.instTable.selectionModel().selectedColumns()) == 1:
            self.removeColum()


    def findNmber(self):  # find the desire number and highlighten the cell
        if self.tabWidget.currentIndex() is 0:
            itemFound = False
            item = self.searchLoanNum.text()
            for rowIndex in range(self.inLoanTable.rowCount()):
                tbItem = self.inLoanTable.item(rowIndex,1)
                if tbItem.text() == item:
                   itemFound = True
                   self.inLoanTable.item(rowIndex,1).setBackground(QtGui.QColor(255,124,124))
                   QApplication.processEvents()
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)


        elif self.tabWidget.currentIndex() is 1:
            itemFound = False
            item = self.giftNumSText.text()
            for rowIndex in range(self.inGiftTable.rowCount()):
                tbItem = self.inGiftTable.item(rowIndex,0)
                if tbItem.text() == item:
                   itemFound = True
                   self.inGiftTable.item(rowIndex,0).setBackground(QtGui.QColor(255,124,124))
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)

        elif self.tabWidget.currentIndex() is 2:
            itemFound = False
            item = self.searchLoanNum_2.text()
            for rowIndex in range(self.inLoanReqTable.rowCount()):
                tbItem = self.inLoanReqTable.item(rowIndex,0)
                if tbItem.text() == item:
                   itemFound = True
                   self.inLoanReqTable.item(rowIndex,0).setBackground(QtGui.QColor(255,124,124))
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)


        elif self.tabWidget.currentIndex() is 3:
            itemFound = False
            item = self.searchLoanNum_3.text()
            for rowIndex in range(self.outLoanTable.rowCount()):
                tbItem = self.outLoanTable.item(rowIndex,1)
                if tbItem.text() == item:
                   itemFound = True
                   self.outLoanTable.item(rowIndex,1).setBackground(QtGui.QColor(255,124,124))
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)

        elif self.tabWidget.currentIndex() is 4:
            itemFound = False
            item = self.searchLoanNum_4.text()
            for rowIndex in range(self.outLoanReqTable.rowCount()):
                tbItem = self.outLoanReqTable.item(rowIndex,0)
                if tbItem.text() == item:
                   itemFound = True
                   self.outLoanReqTable.item(rowIndex,0).setBackground(QtGui.QColor(255,124,124))
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)

        elif self.tabWidget.currentIndex() is 5:
            itemFound = False
            item = self.giftNumSText_2.text()
            for rowIndex in range(self.outGiftTable.rowCount()):
                tbItem = self.outGiftTable.item(rowIndex,0)
                if tbItem.text() == item:
                   itemFound = True
                   self.outGiftTable.item(rowIndex,0).setBackground(QtGui.QColor(255,124,124))
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)

        elif self.tabWidget.currentIndex() is 6:
            itemFound = False
            item = self.acrSText.text()
            for rowIndex in range(self.instTable.rowCount()):
                tbItem = self.instTable.item(rowIndex,0)
                if tbItem.text() == item:
                   itemFound = True
                   self.instTable.item(rowIndex,0).setBackground(QtGui.QColor(255,124,124))
            if itemFound == False:
                QMessageBox.warning(self,"Error","No Match",QMessageBox.Ok)


    def removeRow(self):
        if self.tabWidget.currentIndex() is 0 and self.inLoanTable.rowCount() > 0:
            self.inLoanTable.selectionModel().selectedRows()
            row = self.inLoanTable.selectionModel().selectedIndexes()[0].row()
            self.inLoanTable.removeRow(int(row))
        elif self.tabWidget.currentIndex() is 1 and self.inGiftTable.rowCount() > 0:
            self.inGiftTable.selectionModel().hasSelection()
            row = self.inGiftTable.selectionModel().selectedIndexes()[0].row()
            self.inGiftTable.removeRow(int(row))
        elif self.tabWidget.currentIndex() is 2 and self.inLoanReqTable.rowCount() > 0:
            self.inLoanReqTable.selectionModel().hasSelection()
            row = self.inLoanReqTable.selectionModel().selectedIndexes()[0].row()
            self.inLoanReqTable.removeRow(int(row))
        elif self.tabWidget.currentIndex() is 3 and self.outLoanTable.rowCount() > 0:
            self.outLoanTable.selectionModel().hasSelection()
            row = self.outLoanTable.selectionModel().selectedIndexes()[0].row()
            self.outLoanTable.removeRow(int(row))
        elif self.tabWidget.currentIndex() is 4 and self.outLoanReqTable.rowCount() > 0:
            self.outLoanReqTable.selectionModel().hasSelection()
            row = self.outLoanReqTable.selectionModel().selectedIndexes()[0].row()
            self.outLoanReqTable.removeRow(int(row))
        elif self.tabWidget.currentIndex() is 5 and self.outGiftTable.rowCount() > 0:
            self.outGiftTable.selectionModel().hasSelection()
            row = self.outGiftTable.selectionModel().selectedIndexes()[0].row()
            self.outGiftTable.removeRow(int(row))
        elif self.tabWidget.currentIndex() is 6 and self.instTable.rowCount() > 0:
            self.instTable.selectionModel().hasSelection()
            row = self.instTable.selectionModel().selectedIndexes()[0].row()
            self.instTable.removeRow(int(row))


    def removeColum(self):
        if self.tabWidget.currentIndex() is 0 and self.inLoanTable.columnCount() > 0:
            self.inLoanTable.selectionModel().selectedColumns()
            col = self.inLoanTable.selectionModel().selectedIndexes()[0].column()
            self.inLoanTable.removeColumn(int(col))
        elif self.tabWidget.currentIndex() is 1 and self.inGiftTable.columnCount() > 0:
            self.inGiftTable.selectionModel().selectedColumns()
            col = self.inGiftTable.selectionModel().selectedIndexes()[0].column()
            self.inGiftTable.removeColumn(int(col))
        elif self.tabWidget.currentIndex() is 2 and self.inLoanReqTable.columnCount() > 0:
            self.inLoanReqTable.selectionModel().selectedColumns()
            col = self.inLoanReqTable.selectionModel().selectedIndexes()[0].column()
            self.inLoanReqTable.removeColumn(int(col))
        elif self.tabWidget.currentIndex() is 3 and self.outLoanTable.columnCount() > 0:
            self.outLoanTable.selectionModel().selectedColumns()
            col = self.outLoanTable.selectionModel().selectedIndexes()[0].column()
            self.outLoanTable.removeColumn(int(col))
        elif self.tabWidget.currentIndex() is 4 and self.outLoanReqTable.columnCount() > 0:
            self.outLoanReqTable.selectionModel().selectedColumns()
            col = self.outLoanReqTable.selectionModel().selectedIndexes()[0].column()
            self.outLoanReqTable.removeColumn(int(col))
        elif self.tabWidget.currentIndex() is 5 and self.outGiftTable.columnCount() > 0:
            self.outGiftTable.selectionModel().selectedColumns()
            col = self.outGiftTable.selectionModel().selectedIndexes()[0].column()
            self.outGiftTable.removeColumn(int(col))
        elif self.tabWidget.currentIndex() is 6 and self.instTable.columnCount() > 0:
            self.instTable.selectionModel().selectedColumns()
            col = self.instTable.selectionModel().selectedIndexes()[0].column()
            self.instTable.removeColumn(int(col))


    def createNum(self):  # This is the function of auto generating Loan/Gift number
        if self.tabWidget.currentIndex() is 0 and self.inLoanTable.rowCount() > 0:
            currentRow = self.inLoanTable.currentRow()
            if currentRow != -1:
                currentAcrR = self.inLoanTable.item(currentRow,2)
                selectedYearR = self.inLoanTable.item(currentRow,6)
                if currentAcrR is not None and selectedYearR is not None:
                    currentAcr = currentAcrR.text()
                    selectedYear = selectedYearR.text().split('-')
                    if len(selectedYear) == 3:
                      if int(selectedYear[2]) < 50:
                        currentYear = "20" + selectedYear[2]
                      elif int(selectedYear[2]) > 50:
                        currentYear = "19" + selectedYear[2]
                    elif len(selectedYear) == 2:
                      if int(selectedYear[1]) < 50:
                        currentYear = "20" + selectedYear[1]
                      elif int(selectedYear[1]) > 50:
                        currentYear = "19" + selectedYear[1]
                    elif len(selectedYear) == 1:
                        currentYear = selectedYear[0]
                    loanNum = "IN-"+currentYear+"-"+currentAcr
                    baseLoanNum = "IN-"+currentYear+"-"+currentAcr
                    repeatTimes = 0
                    duplicta = False
                    for rowIndex in range(self.inLoanTable.rowCount()):
                        tbItem = self.inLoanTable.item(rowIndex,1)
                        if tbItem is not None and tbItem.text() == loanNum and rowIndex != currentRow:
                            duplicta = True
                            repeatTimes +=1
                            loanNum = baseLoanNum +"-"+str(repeatTimes)
                    if duplicta:
                        item = QTableWidgetItem(baseLoanNum+"-"+str(repeatTimes))
                        self.inLoanTable.setItem(currentRow,1,item)
                    else:
                        item = QTableWidgetItem(baseLoanNum)
                        self.inLoanTable.setItem(currentRow,1,item)

        elif self.tabWidget.currentIndex() is 1 and self.inGiftTable.rowCount() > 0:
            currentRow = self.inGiftTable.currentRow()
            currentAcrR = self.inGiftTable.item(currentRow ,1)
            selectedYearR = self.inGiftTable.item(currentRow,3)
            if currentAcrR is not None and selectedYearR is not None:
                currentAcr = currentAcrR.text()
                selectedYear = selectedYearR.text().split('-')
                if len(selectedYear) == 3:
                    if int(selectedYear[2]) < 50:
                       currentYear = "20" + selectedYear[2]
                    elif int(selectedYear[2]) > 50:
                        currentYear = "19" + selectedYear[2]
                elif len(selectedYear) == 2:
                    if int(selectedYear[1]) < 50:
                        currentYear = "20" + selectedYear[1]
                    elif int(selectedYear[1]) > 50:
                        currentYear = "19" + selectedYear[1]
                elif len(selectedYear) == 1:
                    currentYear = selectedYear[0]
                loanNum = "G-"+currentYear+"-"+currentAcr
                baseLoanNum = "G-"+currentYear+"-"+currentAcr
                repeatTimes = 0
                duplicta = False
                for rowIndex in range(self.inGiftTable.rowCount()):
                    tbItem = self.inGiftTable.item(rowIndex,0)
                    if tbItem is not None and tbItem.text() == loanNum and rowIndex != currentRow:
                        duplicta = True
                        repeatTimes +=1
                        loanNum = baseLoanNum +"-"+str(repeatTimes)
                if duplicta:
                    item = QTableWidgetItem(baseLoanNum+"-"+str(repeatTimes))
                    self.inGiftTable.setItem(currentRow,0,item)
                else:
                    item = QTableWidgetItem(baseLoanNum)
                    self.inGiftTable.setItem(currentRow,0,item)

        elif self.tabWidget.currentIndex() is 2 and self.inLoanReqTable.rowCount() > 0:
            currentRow = self.inLoanReqTable.currentRow()
            currentAcrR = self.inLoanReqTable.item(currentRow,2)
            selectedYearR = self.inLoanReqTable.item(currentRow,1)
            if currentAcrR is not None and selectedYearR is not None:
                currentAcr = currentAcrR.text().split(';')
                selectedYear = selectedYearR.text().split('-')
                if len(selectedYear) == 3:
                    if int(selectedYear[2]) < 50:
                       currentYear = "20" + selectedYear[2]
                    elif int(selectedYear[2]) > 50:
                        currentYear = "19" + selectedYear[2]
                elif len(selectedYear) == 2:
                    if int(selectedYear[1]) < 50:
                       currentYear = "20" + selectedYear[1]
                    elif int(selectedYear[1]) > 50:
                        currentYear = "19" + selectedYear[1]
                elif len(selectedYear) == 1:
                    currentYear = selectedYear[0]
                loanNum = "IN-"+currentYear+"-"+currentAcr[0]
                baseLoanNum = "IN-"+currentYear+"-"+currentAcr[0]
                repeatTimes = 0
                duplicta = False
                for rowIndex in range(self.inLoanReqTable.rowCount()):
                    tbItem = self.inLoanReqTable.item(rowIndex,0)
                    if tbItem is not None and tbItem.text() == loanNum and rowIndex != currentRow:
                        duplicta = True
                        repeatTimes +=1
                        loanNum = baseLoanNum +"-"+str(repeatTimes)
                if duplicta:
                    item = QTableWidgetItem(baseLoanNum+"-"+str(repeatTimes))
                    self.inLoanReqTable.setItem(currentRow,0,item)
                else:
                    item = QTableWidgetItem(baseLoanNum)
                    self.inLoanReqTable.setItem(currentRow,0,item)

        elif  self.tabWidget.currentIndex() is 3 and self.outLoanTable.rowCount() > 0:
            currentRow = self.outLoanTable.currentRow()
            currentAcrR = self.outLoanTable.item(currentRow ,2)
            selectedYearR = self.outLoanTable.item(currentRow,6)
            if currentAcrR is not None and selectedYearR is not None:
                currentAcr = currentAcrR.text()
                selectedYear = selectedYearR.text().split('-')
                if len(selectedYear) == 3:
                    if int(selectedYear[2]) < 50:
                       currentYear = "20" + selectedYear[2]
                    elif int(selectedYear[2]) > 50:
                        currentYear = "19" + selectedYear[2]
                elif len(selectedYear) == 2:
                    if int(selectedYear[1]) < 50:
                       currentYear = "20" + selectedYear[1]
                    elif int(selectedYear[1]) > 50:
                        currentYear = "19" + selectedYear[1]
                elif len(selectedYear) == 1:
                    currentYear = selectedYear[0]
                baseLoanNum = "OUT-"+currentYear+"-"+currentAcr
                loanNum =  "OUT-"+currentYear+"-"+currentAcr
                duplicta = False
                repeatTimes = 0
                for rowIndex in range(self.outLoanTable.rowCount()):
                    tbItem = self.outLoanTable.item(rowIndex,1)
                    if tbItem is not None and tbItem.text() == loanNum and rowIndex != currentRow:
                        duplicta = True
                        repeatTimes +=1
                        loanNum = baseLoanNum +"-"+str(repeatTimes)
                if duplicta:
                    item = QTableWidgetItem(baseLoanNum+"-"+str(repeatTimes))
                    self.outLoanTable.setItem(currentRow,1,item)
                else:
                    item = QTableWidgetItem(baseLoanNum)
                    self.outLoanTable.setItem(currentRow,1,item)

        elif self.tabWidget.currentIndex() is 4 and self.outLoanReqTable.rowCount() > 0:
            currentRow = self.outLoanReqTable.currentRow()
            currentAcrR = self.outLoanReqTable.item(currentRow ,2)
            selectedYearR = self.outLoanReqTable.item(currentRow,1)
            if currentAcrR is not None and selectedYearR is not None:
                currentAcr = currentAcrR.text().split(';')
                selectedYear = selectedYearR.text().split('-')
                if len(selectedYear) == 3:
                    if int(selectedYear[2]) < 50:
                       currentYear = "20" + selectedYear[2]
                    elif int(selectedYear[2]) > 50:
                        currentYear = "19" + selectedYear[2]
                elif len(selectedYear) == 2:
                    if int(selectedYear[1]) < 50:
                       currentYear = "20" + selectedYear[1]
                    elif int(selectedYear[1]) > 50:
                        currentYear = "19" + selectedYear[1]
                elif len(selectedYear) == 1:
                    currentYear = selectedYear[0]
                baseLoanNum = "OUT-"+currentYear+"-"+currentAcr[0]
                loanNum =  "OUT-"+currentYear+"-"+currentAcr[0]
                duplicta = False
                repeatTimes = 0
                for rowIndex in range(self.outLoanReqTable.rowCount()):
                    tbItem = self.outLoanReqTable.item(rowIndex,0)
                    if tbItem is not None and tbItem.text() == loanNum and rowIndex != currentRow:
                        duplicta = True
                        repeatTimes +=1
                        loanNum = baseLoanNum +"-"+str(repeatTimes)
                if duplicta:
                    item = QTableWidgetItem(baseLoanNum+"-"+str(repeatTimes))
                    self.outLoanReqTable.setItem(currentRow,0,item)
                else:
                    item = QTableWidgetItem(baseLoanNum)
                    self.outLoanReqTable.setItem(currentRow,0,item)

        elif self.tabWidget.currentIndex() is 5 and self.outGiftTable.rowCount() > 0:
            currentRow = self.outGiftTable.currentRow()
            currentAcrR = self.outGiftTable.item(currentRow ,1)
            selectedYearR = self.outGiftTable.item(currentRow,5)
            if currentAcrR is not None and selectedYearR is not None:
                currentAcr = currentAcrR.text()
                selectedYear = selectedYearR.text().split('-')
                if len(selectedYear) == 3:
                    if int(selectedYear[2]) < 50:
                       currentYear = "20" + selectedYear[2]
                    elif int(selectedYear[2]) > 50:
                        currentYear = "19" + selectedYear[2]
                elif len(selectedYear) == 2:
                    if int(selectedYear[1]) < 50:
                       currentYear = "20" + selectedYear[1]
                    elif int(selectedYear[1]) > 50:
                        currentYear = "19" + selectedYear[1]
                elif len(selectedYear) == 1:
                    currentYear = selectedYear[0]
                baseLoanNum = "OUT-"+currentYear+"-"+currentAcr
                loanNum =  "OUT-"+currentYear+"-"+currentAcr
                duplicta = False
                repeatTimes = 0
                for rowIndex in range(self.outGiftTable.rowCount()):
                    tbItem = self.outGiftTable.item(rowIndex,0)
                    if tbItem is not None and tbItem.text() == loanNum and rowIndex != currentRow:
                        duplicta = True
                        repeatTimes +=1
                        loanNum = baseLoanNum +"-"+str(repeatTimes)
                if duplicta:
                    item = QTableWidgetItem(baseLoanNum+"-"+str(repeatTimes))
                    self.outGiftTable.setItem(currentRow,0,item)
                else:
                    item = QTableWidgetItem(baseLoanNum)
                    self.outGiftTable.setItem(currentRow,0,item)


    def changeHeader(self,index): # Double click to change the header label
        if self.tabWidget.currentIndex() is 0:
           it = self.inLoanTable.horizontalHeaderItem(index)
           if it is None:
              val = self.inLoanTable.model().headerData(index,QtCore.Qt.Horizontal)
              it = QtWidgets.QTableWidgetItem(str(val))
              self.inLoanTable.setHorizontalHeaderItem(index,it)
              oldHeader = it.text()
              newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
           else:
              oldHeader = it.text()
              newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
           if ok:
              it.setText(newHeader)

        elif self.tabWidget.currentIndex() is 1:
            it = self.inGiftTable.horizontalHeaderItem(index)
            if it is None:
              val = self.inGiftTable.model().headerData(index,QtCore.Qt.Horizontal)
              it = QtWidgets.QTableWidgetItem(str(val))
              self.inGiftTable.setHorizontalHeaderItem(index,it)
              oldHeader = it.text()
              newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
            else:
              oldHeader = it.text()
              newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
            if ok:
              it.setText(newHeader)

        elif self.tabWidget.currentIndex() is 2:
             it = self.inLoanReqTable.horizontalHeaderItem(index)
             if it is None:
               val = self.inLoanReqTable.model().headerData(index,QtCore.Qt.Horizontal)
               it = QtWidgets.QTableWidgetItem(str(val))
               self.inLoanReqTable.setHorizontalHeaderItem(index,it)
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             else:
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             if ok:
               it.setText(newHeader)

        elif self.tabWidget.currentIndex() is 3:
             it = self.outLoanTable.horizontalHeaderItem(index)
             if it is None:
               val = self.outLoanTable.model().headerData(index,QtCore.Qt.Horizontal)
               it = QtWidgets.QTableWidgetItem(str(val))
               self.outLoanTable.setHorizontalHeaderItem(index,it)
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             else:
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             if ok:
               it.setText(newHeader)
        elif self.tabWidget.currentIndex() is 4:
             it = self.outLoanReqTable.horizontalHeaderItem(index)
             if it is None:
               val = self.outLoanReqTable.model().headerData(index,QtCore.Qt.Horizontal)
               it = QtWidgets.QTableWidgetItem(str(val))
               self.outLoanReqTable.setHorizontalHeaderItem(index,it)
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             else:
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             if ok:
               it.setText(newHeader)

        elif self.tabWidget.currentIndex() is 5:
             it = self.outGiftTable.horizontalHeaderItem(index)
             if it is None:
               val = self.outGiftTable.model().headerData(index,QtCore.Qt.Horizontal)
               it = QtWidgets.QTableWidgetItem(str(val))
               self.outGiftTable.setHorizontalHeaderItem(index,it)
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             else:
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             if ok:
               it.setText(newHeader)

        elif self.tabWidget.currentIndex() is 6:
             it = self.instTable.horizontalHeaderItem(index)
             if it is None:
               val = self.instTable.model().headerData(index,QtCore.Qt.Horizontal)
               it = QtWidgets.QTableWidgetItem(str(val))
               self.instTable.setHorizontalHeaderItem(index,it)
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             else:
               oldHeader = it.text()
               newHeader, ok = QInputDialog.getText(self,'Change header label  column %d' % index,'Header:',QtWidgets.QLineEdit.Normal,oldHeader)
             if ok:
               it.setText(newHeader)


    def autoFill(self):  # This is the function of auto fill contacts/institution names/address
        if self.tabWidget.currentIndex() is 0 and self.instTable.rowCount() > 0:
            currentRow = self.inLoanTable.currentRow()
            newAcr = self.inLoanTable.item(currentRow,2)
            for rowIndex in range(self.instTable.rowCount()):
                tbItem = self.instTable.item(rowIndex,0)
                if tbItem.text() == newAcr.text():
                    fillInsitution = self.instTable.item(rowIndex,1)
                    fillContact = self.instTable.item(rowIndex,2)
                    fillAdress = self.instTable.item(rowIndex,3)
                else:
                    pass
            institution = QTableWidgetItem(fillInsitution)
            contact = QTableWidgetItem(fillContact)
            adress = QTableWidgetItem(fillAdress)
            self.inLoanTable.setItem(currentRow,3,institution)
            self.inLoanTable.setItem(currentRow,4,contact)
            self.inLoanTable.setItem(currentRow,5,adress)
            self.inLoanTable.resizeRowsToContents()

        elif self.tabWidget.currentIndex() is 1 and self.instTable.rowCount() > 0:
            currentRow = self.inGiftTable.currentRow()
            newAcr = self.inGiftTable.item(currentRow,1)
            for rowIndex in range(self.instTable.rowCount()):
                tbItem = self.instTable.item(rowIndex,0)
                if tbItem.text() == newAcr.text():
                    fillInsitution = self.instTable.item(rowIndex,1)
                else:
                    pass
            institution = QTableWidgetItem(fillInsitution)
            self.inGiftTable.setItem(currentRow,2,institution)
            self.inGiftTable.resizeRowsToContents()

        elif self.tabWidget.currentIndex() is 2 and self.instTable.rowCount() > 0:
            pass

        elif self.tabWidget.currentIndex() is 3 and self.instTable.rowCount() > 0:
            currentRow = self.outLoanTable.currentRow()
            newAcr = self.outLoanTable.item(currentRow,2)
            for rowIndex in range(self.instTable.rowCount()):
                tbItem = self.instTable.item(rowIndex,0)
                if tbItem.text() == newAcr.text():
                    fillInsitution = self.instTable.item(rowIndex,1)
                    fillContact = self.instTable.item(rowIndex,2)
                    fillAdress = self.instTable.item(rowIndex,3)
                else:
                    pass
            institution = QTableWidgetItem(fillInsitution)
            contact = QTableWidgetItem(fillContact)
            adress = QTableWidgetItem(fillAdress)
            self.outLoanTable.setItem(currentRow,3,institution)
            self.outLoanTable.setItem(currentRow,4,contact)
            self.outLoanTable.setItem(currentRow,5,adress)
            self.outLoanTable.resizeRowsToContents()

        elif self.tabWidget.currentIndex() is 4 and self.instTable.rowCount() > 0:
            pass

        elif self.tabWidget.currentIndex() is 5 and self.instTable.rowCount() > 0:
            currentRow = self.outGiftTable.currentRow()
            newAcr = self.outGiftTable.item(currentRow,1)
            for rowIndex in range(self.instTable.rowCount()):
                tbItem = self.instTable.item(rowIndex,0)
                if tbItem.text() == newAcr.text():
                    fillInsitution = self.instTable.item(rowIndex,1)
                    fillContact = self.instTable.item(rowIndex,2)
                    fillAdress = self.instTable.item(rowIndex,3)
                else:
                    pass
            institution = QTableWidgetItem(fillInsitution)
            contact = QTableWidgetItem(fillContact)
            adress = QTableWidgetItem(fillAdress)
            self.outGiftTable.setItem(currentRow,2,institution)
            self.outGiftTable.setItem(currentRow,3,contact)
            self.outGiftTable.setItem(currentRow,4,adress)
            self.outGiftTable.resizeRowsToContents()


    def saveFile(self):  # save new CSV files
        if  self.tabWidget.currentIndex() is 0 and self.inLoanTable.rowCount() > 0:
            self.saveFileCore(self.inLoanTable)
        elif  self.tabWidget.currentIndex() is 1 and self.inGiftTable.rowCount() > 0:
            self.saveFileCore(self.inGiftTable)
        elif  self.tabWidget.currentIndex() is 2 and self.inLoanReqTable.rowCount() > 0:
            self.saveFileCore(self.inLoanReqTable)
        elif  self.tabWidget.currentIndex() is 3 and self.outLoanTable.rowCount() > 0:
            self.saveFileCore(self.outLoanTable)
        elif  self.tabWidget.currentIndex() is 4 and self.outLoanReqTable.rowCount() > 0:
            self.saveFileCore(self.outLoanReqTable)
        elif  self.tabWidget.currentIndex() is 5 and self.outGiftTable.rowCount() > 0:
            self.saveFileCore(self.outGiftTable)
        elif  self.tabWidget.currentIndex() is 6 and self.instTable.rowCount() > 0:
            self.saveFileCore(self.instTable)

    def saveFileCore(self,file:QTableWidget):
        path = QFileDialog.getSaveFileName(self,'Save CSV',os.getenv('HOME'),'CSV(*.csv)')
        if path[0] != '':
            with open(path[0],'w')as csvFile:
                writer = csv.writer(csvFile,delimiter=',')
                headerData = []
                for col in range(file.columnCount()):
                    header = file.horizontalHeaderItem(col).text()
                    headerData.append(header)
                writer.writerow(headerData)
                for row in range(file.rowCount()):
                    rowData = []
                    for column in range(file.columnCount()):
                        item = file.item(row,column)
                        if item is not None:
                            rowData.append(item.text())
                        else:
                            rowData.append('')
                    writer.writerow(rowData)

    def openEmial(self):
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_Email()
        self.ui.setupUi(self.window)
        self.window.show()

        self.emailCore()

    def emailCore(self):
        if self.tabWidget.currentIndex() is 0 and self.inLoanTable.rowCount() > 0:
            self.ui.loanInfo.setRowCount(100)
            row = self.ui.loanInfo.rowCount()
            self.ui.loanInfo.insertRow(row)
            self.ui.loanInfo.setColumnCount(6)
            column = self.ui.loanInfo.columnCount()
            self.ui.loanInfo.insertColumn(column)
            duedate = ""
            tem = []
            for row in range(self.inLoanTable.rowCount()):
                item = self.inLoanTable.item(row,7)
                if item is not None:
                    if item.text() != 'Not specified.' and item.text() != '':
                        dueDate = item.text().split('-')
                        if len(dueDate) == 3:
                           if dueDate[1] == "Jan" and len(dueDate) == 3:
                             month = "01"
                           elif dueDate[1] == "Feb" and len(dueDate) == 3:
                             month = "02"
                           elif dueDate[1] == "Mar" and len(dueDate) == 3:
                             month = "03"
                           elif dueDate[1] == "Apr" and len(dueDate) == 3:
                             month = "04"
                           elif dueDate[1] == "May" and len(dueDate) == 3:
                             month = "05"
                           elif dueDate[1] == "Jun" and len(dueDate) == 3:
                             month = "06"
                           elif dueDate[1] == "Jul" and len(dueDate) == 3:
                             month = "07"
                           elif dueDate[1] == "Aug" and len(dueDate) == 3:
                             month = "08"
                           elif dueDate[1] == "Sep" and len(dueDate) == 3:
                             month = "09"
                           elif dueDate[1] == "Oct" and len(dueDate) == 3:
                             month = "10"
                           elif dueDate[1] == "Nov" and len(dueDate) == 3:
                             month = "11"
                           elif dueDate[1] == "Dec" and len(dueDate) == 3:
                             month = "12"
                           indate = "20"+str(dueDate[2]) +'-'+ month +'-' + dueDate[0]
                           if (datetime.datetime.now() + datetime.timedelta(days = 365)).strftime("%Y-%m-%d") >= indate and (datetime.datetime.now()).strftime("%Y-%m-%d") < indate :
                              loanNum = self.inLoanTable.item(row,1)
                              acr = self.inLoanTable.item(row,2)
                              institution = self.inLoanTable.item(row,3)
                              contacts = self.inLoanTable.item(row,4)
                              loanDueDate = self.inLoanTable.item(row,7)
                              invoiceDate = self.inLoanTable.item(row,6)
                              if loanNum != None and acr != None and institution != None and contacts != None and loanDueDate != None and invoiceDate != None:
                                tem.append([loanNum.text(),acr.text(),institution.text(),contacts.text(),loanDueDate.text(),invoiceDate.text()])
            for index in range(len(tem)):
                self.ui.loanInfo.setItem(index,0,QTableWidgetItem(tem[index][0]))
                self.ui.loanInfo.setItem(index,1,QTableWidgetItem(tem[index][1]))
                self.ui.loanInfo.setItem(index,2,QTableWidgetItem(tem[index][2]))
                self.ui.loanInfo.setItem(index,3,QTableWidgetItem(tem[index][3]))
                self.ui.loanInfo.setItem(index,4,QTableWidgetItem(tem[index][4]))
                self.ui.loanInfo.setItem(index,5,QTableWidgetItem(tem[index][5]))
            self.ui.loanInfo.resizeRowsToContents()
            self.emailTemplate()

        elif self.tabWidget.currentIndex() is 3 and self.outLoanTable.rowCount() > 0:
            self.ui.loanInfo.setRowCount(100)
            row = self.ui.loanInfo.rowCount()
            self.ui.loanInfo.insertRow(row)
            self.ui.loanInfo.setColumnCount(6)
            column = self.ui.loanInfo.columnCount()
            self.ui.loanInfo.insertColumn(column)
            duedate = ""
            tem = []
            for row in range(self.outLoanTable.rowCount()):
                item = self.outLoanTable.item(row,8)
                if item is not None:
                    if item.text() != 'Not specified.' and item.text() != '':
                        dueDate = item.text().split('-')
                        if len(dueDate) == 3:
                           if dueDate[1] == "Jan" and len(dueDate) == 3:
                             month = "01"
                           elif dueDate[1] == "Feb" and len(dueDate) == 3:
                             month = "02"
                           elif dueDate[1] == "Mar" and len(dueDate) == 3:
                             month = "03"
                           elif dueDate[1] == "Apr" and len(dueDate) == 3:
                             month = "04"
                           elif dueDate[1] == "May" and len(dueDate) == 3:
                             month = "05"
                           elif dueDate[1] == "Jun" and len(dueDate) == 3:
                             month = "06"
                           elif dueDate[1] == "Jul" and len(dueDate) == 3:
                             month = "07"
                           elif dueDate[1] == "Aug" and len(dueDate) == 3:
                             month = "08"
                           elif dueDate[1] == "Sep" and len(dueDate) == 3:
                             month = "09"
                           elif dueDate[1] == "Oct" and len(dueDate) == 3:
                             month = "10"
                           elif dueDate[1] == "Nov" and len(dueDate) == 3:
                             month = "11"
                           elif dueDate[1] == "Dec" and len(dueDate) == 3:
                             month = "12"
                           indate = "20"+str(dueDate[2]) +'-'+ month +'-' + dueDate[0]
                           if (datetime.datetime.now() + datetime.timedelta(days = 365)).strftime("%Y-%m-%d") >= indate and (datetime.datetime.now()).strftime("%Y-%m-%d") < indate :
                              loanNum = self.outLoanTable.item(row,1)
                              acr = self.outLoanTable.item(row,2)
                              institution = self.outLoanTable.item(row,3)
                              contacts = self.outLoanTable.item(row,4)
                              loanDueDate = self.outLoanTable.item(row,7)
                              invoiceDate = self.outLoanTable.item(row,6)
                              if loanNum != None and acr != None and institution != None and contacts != None and loanDueDate != None and invoiceDate != None:
                                tem.append([loanNum.text(),acr.text(),institution.text(),contacts.text(),loanDueDate.text(),invoiceDate.text()])
            for index in range(len(tem)):
                self.ui.loanInfo.setItem(index,0,QTableWidgetItem(tem[index][0]))
                self.ui.loanInfo.setItem(index,1,QTableWidgetItem(tem[index][1]))
                self.ui.loanInfo.setItem(index,2,QTableWidgetItem(tem[index][2]))
                self.ui.loanInfo.setItem(index,3,QTableWidgetItem(tem[index][3]))
                self.ui.loanInfo.setItem(index,4,QTableWidgetItem(tem[index][4]))
                self.ui.loanInfo.setItem(index,5,QTableWidgetItem(tem[index][5]))
            self.ui.loanInfo.resizeRowsToContents()
            self.emailTemplate()

    def emailTemplate(self):
            currentRow = self.ui.loanInfo.currentRow()
            self.ui.template_2.setPlainText("Hello ,\nI am emailing to inform you that Loan-   "+"sent to UWO on-  " + "is being returned to you via Canada Post (tracking number: )."
                                        "\nAttached is a copy of our return documentation for you to sign and email back once the specimens have been processed.\n Thank you for loaning us this material!")

    def clearColour(self):
        if self.tabWidget.currentIndex() is 0:
            self.resetBackgroundsCore(self.inLoanTable)
        elif self.tabWidget.currentIndex() is 1:
            self.resetBackgroundsCore(self.inGiftTable)
        elif self.tabWidget.currentIndex() is 2:
            self.resetBackgroundsCore(self.inLoanReqTable)
        elif self.tabWidget.currentIndex() is 3:
            self.resetBackgroundsCore(self.outLoanTable)
        elif self.tabWidget.currentIndex() is 4:
            self.resetBackgroundsCore(self.outLoanReqTable)
        elif self.tabWidget.currentIndex() is 5:
            self.resetBackgroundsCore(self.outGiftTable)
        elif self.tabWidget.currentIndex() is 6:
            self.resetBackgroundsCore(self.instTable)

    def resetBackgroundsCore(self,table:QTableWidget):
        model = table.model()
        for row in range(table.rowCount()):
            for column in range(table.columnCount()):
                index = model.index(row,column)
                if index.data(QtCore.Qt.BackgroundColorRole):
                    model.setData(index,None,QtCore.Qt.BackgroundColorRole)
        table.clearSelection()



def window():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = mainForm()
    MainWindow.show()
    sys.exit(app.exec_())

window()
