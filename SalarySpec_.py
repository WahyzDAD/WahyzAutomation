# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'SalarySpec.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(640, 480)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(50, 40, 251, 351))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.widget_2 = QtWidgets.QWidget(self.verticalLayoutWidget)
        self.widget_2.setObjectName("widget_2")
        self.btn_SelectFile = QtWidgets.QPushButton(self.widget_2)
        self.btn_SelectFile.setGeometry(QtCore.QRect(0, 0, 75, 23))
        self.btn_SelectFile.setObjectName("btn_SelectFile")
        self.lineEdit_SelectedFile = QtWidgets.QLineEdit(self.widget_2)
        self.lineEdit_SelectedFile.setGeometry(QtCore.QRect(0, 20, 231, 20))
        self.lineEdit_SelectedFile.setObjectName("lineEdit_SelectedFile")
        self.btn_SaveDir = QtWidgets.QPushButton(self.widget_2)
        self.btn_SaveDir.setGeometry(QtCore.QRect(0, 50, 151, 23))
        self.btn_SaveDir.setObjectName("btn_SaveDir")
        self.lineEdit_SaveDir = QtWidgets.QLineEdit(self.widget_2)
        self.lineEdit_SaveDir.setGeometry(QtCore.QRect(0, 80, 231, 20))
        self.lineEdit_SaveDir.setObjectName("lineEdit_SaveDir")
        self.verticalLayout.addWidget(self.widget_2)
        self.widget = QtWidgets.QWidget(self.verticalLayoutWidget)
        self.widget.setObjectName("widget")
        self.lineEdit_From = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_From.setGeometry(QtCore.QRect(0, 20, 113, 20))
        self.lineEdit_From.setObjectName("lineEdit_From")
        self.label_From = QtWidgets.QLabel(self.widget)
        self.label_From.setGeometry(QtCore.QRect(0, 0, 56, 12))
        self.label_From.setObjectName("label_From")
        self.lineEdit_To = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_To.setGeometry(QtCore.QRect(120, 20, 113, 20))
        self.lineEdit_To.setObjectName("lineEdit_To")
        self.label_To = QtWidgets.QLabel(self.widget)
        self.label_To.setGeometry(QtCore.QRect(120, 0, 56, 12))
        self.label_To.setObjectName("label_To")
        self.verticalLayout.addWidget(self.widget)
        self.widget_3 = QtWidgets.QWidget(self.verticalLayoutWidget)
        self.widget_3.setObjectName("widget_3")
        self.btn_MakePDFs = QtWidgets.QPushButton(self.widget_3)
        self.btn_MakePDFs.setGeometry(QtCore.QRect(0, 10, 75, 23))
        self.btn_MakePDFs.setObjectName("btn_MakePDFs")
        self.textEdit = QtWidgets.QTextEdit(self.widget_3)
        self.textEdit.setGeometry(QtCore.QRect(0, 40, 251, 51))
        self.textEdit.setObjectName("textEdit")
        self.progressBar = QtWidgets.QProgressBar(self.widget_3)
        self.progressBar.setGeometry(QtCore.QRect(0, 90, 241, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout.addWidget(self.widget_3)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 640, 21))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionOpen_File = QtWidgets.QAction(MainWindow)
        self.actionOpen_File.setObjectName("actionOpen_File")
        self.menuFile.addAction(self.actionOpen_File)
        self.menubar.addAction(self.menuFile.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.btn_SelectFile.setText(_translate("MainWindow", "Select File"))
        self.btn_SaveDir.setText(_translate("MainWindow", "Select Saving Directory"))
        self.label_From.setText(_translate("MainWindow", "# From"))
        self.label_To.setText(_translate("MainWindow", "# To"))
        self.btn_MakePDFs.setText(_translate("MainWindow", "Make PDFs"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.actionOpen_File.setText(_translate("MainWindow", "Open File"))