import sys
import time
import typing

from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLineEdit,
    QVBoxLayout,
    QMainWindow,
    QWidget,
    QPushButton,
    QSpinBox,
    QGroupBox,
    QLabel,
    QPlainTextEdit,
    QFileDialog,
)
from PyQt5.QtCore import QThread, pyqtSignal, pyqtSlot,  QRunnable, QThreadPool, QObject
import pandas as pd

import utils

class WindowClass(QMainWindow):
    def __init__(self):
        super().__init__()
        
        
        # Buttons
        self.btn_SelectFile = QPushButton("Select File")
        self.btn_SelectFile.clicked.connect(self.select_file)
        
        self.lineEdit_SelectedFile = QLineEdit()
        
        self.lineEdit_From = QLineEdit()
        self.lineEdit_To = QLineEdit()
        
        self.btn_previewMailList = QPushButton("Preview Mail List")
        self.btn_previewMailList.clicked.connect(self.preview_mail_list)
        
        self.mailList = QLabel()
        self.mailList.setStyleSheet("border: 1px solid black;")
        self.mailList.setFixedHeight(100)
        
        self.btn_sendEmails = QPushButton("Send Emails")
        self.btn_sendEmails.clicked.connect(self.send_emails)
        
        
        self.setWindowTitle("WahyzAutomation")
        self.createForm()

    def createForm(self):
        # creating a form layout
        widget = QWidget()
        layout = QFormLayout(widget)
        
        # adding rows
        layout.addRow(QLabel("Select File"), self.btn_SelectFile)
        layout.addRow(QLabel("Selected File"), self.lineEdit_SelectedFile)
        layout.addRow(QLabel("Number From"), self.lineEdit_From)
        layout.addRow(QLabel("Number To"), self.lineEdit_To)
        layout.addRow(QLabel("Preview Mail List"), self.btn_previewMailList)
        layout.addRow(QLabel("mailList"), self.mailList)
        layout.addRow(QLabel("Send Emails"), self.btn_sendEmails)
        
        self.setCentralWidget(widget)
        
    def select_file(self):
        """
        Click "Select File",
        Display Selected File,
        Open Workbook as self.wb
        """
        try:
            filename = QFileDialog.getOpenFileName(self, 'Open file', './')
            self.lineEdit_SelectedFile.setText(filename[0]) 
            
            self.wb = utils.load_workbook(filename = filename[0])
        
        # 각 시트를 변수에 할당
        
        except Exception as e:
            print(e)

    def get_fromto(self):
        """
        시작 행 번호부터 마지막 행 번호를
        From #, To #에 각각 적고 그 값을 가져온다.
        """
        self.number_from = int(self.lineEdit_From.text())
        self.number_to = int(self.lineEdit_To.text())
        
        assert self.number_from <= self.number_to

    
    def preview_mail_list(self):
        try:
            self.get_fromto()
            self.mailList.setText(f"From {self.number_from} To {self.number_to}")
            # self.mailList.setText(df)
        except AssertionError as e:
            self.mailList.setText(f"The number for 'To' must be greater than or equal to 'From'")
        except:
            self.mailList.setText(f"Set From, To")
    
    def send_emails(self):
        pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WindowClass()
    window.show()
    sys.exit(app.exec_())