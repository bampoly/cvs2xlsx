# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MainWindow.ui'
#
# Created: Thu Nov 12 23:44:12 2015
#      by: pyside-uic 0.2.15 running on PySide 1.2.4
#
# WARNING! All changes made in this file will be lost!

from PySide import QtCore, QtGui
import csv
import glob
import os
import xlsxwriter

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(472, 191)
        self.gridLayoutWidget = QtGui.QWidget(MainWindow)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(50, 40, 371, 109))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtGui.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.source_label = QtGui.QLabel(self.gridLayoutWidget)
        self.source_label.setObjectName("source_label")
        self.gridLayout.addWidget(self.source_label, 0, 0, 1, 1)
        self.sourceFolder_lineEdit = QtGui.QLineEdit(self.gridLayoutWidget)
        self.sourceFolder_lineEdit.setObjectName("sourceFolder_lineEdit")
        self.gridLayout.addWidget(self.sourceFolder_lineEdit, 0, 1, 1, 1)
        self.dist_label = QtGui.QLabel(self.gridLayoutWidget)
        self.dist_label.setObjectName("dist_label")
        self.gridLayout.addWidget(self.dist_label, 1, 0, 1, 1)
        self.destFolder_lineEdit = QtGui.QLineEdit(self.gridLayoutWidget)
        self.destFolder_lineEdit.setObjectName("destFolder_lineEdit")
        self.gridLayout.addWidget(self.destFolder_lineEdit, 1, 1, 1, 1)
        self.selectFolder_btn = QtGui.QPushButton(self.gridLayoutWidget)
        self.selectFolder_btn.setObjectName("selectFolder_btn")
        self.gridLayout.addWidget(self.selectFolder_btn, 0, 2, 1, 1)
        self.convert_btn = QtGui.QPushButton(self.gridLayoutWidget)
        self.convert_btn.setObjectName("convert_btn")
        self.gridLayout.addWidget(self.convert_btn, 2, 1, 1, 1)
        self.exit_btn = QtGui.QPushButton(self.gridLayoutWidget)
        self.exit_btn.setObjectName("exit_btn")
        self.gridLayout.addWidget(self.exit_btn, 2, 2, 1, 1)
        self.selectFolder_btn_2 = QtGui.QPushButton(self.gridLayoutWidget)
        self.selectFolder_btn_2.setObjectName("selectFolder_btn_2")
        self.gridLayout.addWidget(self.selectFolder_btn_2, 1, 2, 1, 1)

        self.retranslateUi(MainWindow)
        QtCore.QObject.connect(self.exit_btn, QtCore.SIGNAL("clicked()"), MainWindow.close)
        QtCore.QObject.connect(self.selectFolder_btn, QtCore.SIGNAL("clicked()"), MainWindow.openFolder)
        QtCore.QObject.connect(self.convert_btn, QtCore.SIGNAL("clicked()"), MainWindow.convert)
        QtCore.QObject.connect(self.selectFolder_btn_2, QtCore.SIGNAL("clicked()"), MainWindow.openFolder)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QtGui.QApplication.translate("MainWindow", "csv2xlsx Converter", None, QtGui.QApplication.UnicodeUTF8))
        self.source_label.setText(QtGui.QApplication.translate("MainWindow", "Source", None, QtGui.QApplication.UnicodeUTF8))
        self.sourceFolder_lineEdit.setText(QtGui.QApplication.translate("MainWindow", "C:\\", None, QtGui.QApplication.UnicodeUTF8))
        self.dist_label.setText(QtGui.QApplication.translate("MainWindow", "Destination", None, QtGui.QApplication.UnicodeUTF8))
        self.destFolder_lineEdit.setText(QtGui.QApplication.translate("MainWindow", "C:\\", None, QtGui.QApplication.UnicodeUTF8))
        self.selectFolder_btn.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))
        self.convert_btn.setText(QtGui.QApplication.translate("MainWindow", "Convert", None, QtGui.QApplication.UnicodeUTF8))
        self.exit_btn.setText(QtGui.QApplication.translate("MainWindow", "Exit", None, QtGui.QApplication.UnicodeUTF8))
        self.selectFolder_btn_2.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))

    def openFolder(self):
        ID_btn = self.sender() .objectName()# this line is used to know which button is clicked	
        self.dir = QtGui.QFileDialog.getExistingDirectory()
        self.csvFiles = [file for file in  os.listdir(str(self.dir)) if file.endswith(".csv")]
        print ID_btn
        if ID_btn == "selectFolder_btn":
		    self.sourceFolder_lineEdit.setText(self.dir)
        elif ID_btn ==  "selectFolder_btn_2": 
            self.destFolder_lineEdit.setText(self.dir)
		
		
    def convert(self):
         for index, file in enumerate(self.csvFiles):
			fileName = file[0:file.find(".")]
			xlsxFile = xlsxwriter.Workbook(fileName + ".xlsx")
			worksheet = xlsxFile.add_worksheet()
			source_full_path = os.path.join(self.dir , file)
			with open(source_full_path, "rb") as f:
				 data = csv.reader(f)
				 for row,  dataInRow in enumerate(data):
					for col, dataInCell in enumerate(dataInRow):
						worksheet.write(row, col, dataInCell)
         xlsxFile.close()