# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Email.ui'
#
# Created by: PyQt5 UI code generator 5.15.0
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Email(object):
    def setupUi(self, Email):
        Email.setObjectName("Email")
        Email.resize(822, 615)
        self.centralwidget = QtWidgets.QWidget(Email)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.loanInfo = QtWidgets.QTableWidget(self.centralwidget)
        self.loanInfo.setObjectName("loanInfo")
        self.loanInfo.setColumnCount(0)
        self.loanInfo.setRowCount(0)
        self.gridLayout.addWidget(self.loanInfo, 0, 0, 1, 2)
        self.template_2 = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.template_2.setObjectName("template_2")
        self.gridLayout.addWidget(self.template_2, 1, 0, 1, 2)
        Email.setCentralWidget(self.centralwidget)

        self.retranslateUi(Email)
        QtCore.QMetaObject.connectSlotsByName(Email)

    def retranslateUi(self, Email):
        _translate = QtCore.QCoreApplication.translate
        Email.setWindowTitle(_translate("Email", "MainWindow"))
