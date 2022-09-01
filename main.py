import os
import sys
import time
import typing
import traceback

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
    QScrollArea,
    QTableView,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QRunnable, QThreadPool, QObject, QSortFilterProxyModel
import pandas as pd

import openpyxl as xl
from openpyxl import Workbook, load_workbook
import pyperclip as clp
import pyautogui
import time

import models
from models import MyTableModel


class WindowClass(QMainWindow):
    def __init__(self):
        super().__init__()
        
        
        # Buttons
        self.btn_SelectFile = QPushButton("Select File")
        self.btn_SelectFile.clicked.connect(self.select_email_list_file)
        
        self.lineEdit_SelectedFile = QLineEdit()
        
        self.lineEdit_Year = QLineEdit()
        self.lineEdit_Month = QLineEdit()
        
        self.lineEdit_From = QLineEdit()
        self.lineEdit_To = QLineEdit()
        
        self.btn_previewMailList = QPushButton("Preview Mail List")
        self.btn_previewMailList.clicked.connect(self.preview_mail_list)
        
        self.mailListLabel = QLabel()
        self.mailListLabel.setStyleSheet("border: 1px solid black;")
        # self.mailListLabel.setFixedHeight(100)
        
        # Set up the view
        self.tableView = QTableView()
        
        self.messages = QLabel()
        self.messages.setStyleSheet("border: 1px solid black;")
        self.messages.setFixedHeight(100)
        self.message_list = []
        
        self.btn_SelectPDFFolder = QPushButton("Select Folder")
        self.btn_SelectPDFFolder.clicked.connect(self.select_pdf_folder)
        self.lineEdit_SelectedFolder = QLineEdit()
        
        self.ready_email_site = QLabel()
        self.ready_email_site.setStyleSheet("border: 1px solid black;")
        self.ready_email_site.setFixedHeight(30)
        self.ready_email_site.setText(f"이메일 사이트를 켜세요.")
        
        self.btn_sendEmails = QPushButton("Send Emails")
        self.btn_sendEmails.clicked.connect(self.send_emails)
        
        
        self.setWindowTitle("WahyzAutomation")
        self.createForm()

    def createForm(self):
        # creating a form layout
        widget = QWidget()
        layout = QFormLayout(widget)
        
        # adding rows
        layout.addRow(QLabel("Select Email List File"), self.btn_SelectFile)
        layout.addRow(QLabel("Selected Email List File"), self.lineEdit_SelectedFile)
        layout.addRow(QLabel("Year"), self.lineEdit_Year)
        layout.addRow(QLabel("Month"), self.lineEdit_Month)
        layout.addRow(QLabel("Number From"), self.lineEdit_From)
        layout.addRow(QLabel("Number To"), self.lineEdit_To)
        layout.addRow(QLabel("Preview Mail List"), self.btn_previewMailList)
    
        # self.scrollAreaforMailList.setWidget(layout.addRow(QLabel("Mail List"), self.mailListLabel))
        layout.addRow(QLabel("Mail List"), self.mailListLabel)
        layout.addRow(self.tableView)
        
        layout.addRow(QLabel("Messages"), self.messages)
        layout.addRow(QLabel("Select PDF Folder"), self.btn_SelectPDFFolder)
        layout.addRow(QLabel("Selected PDF Folder"), self.lineEdit_SelectedFolder)
        layout.addRow(QLabel("Get Ready Email Site"), self.ready_email_site)
        layout.addRow(QLabel("Send Emails"), self.btn_sendEmails)
        
        self.setCentralWidget(widget)
        
    def select_email_list_file(self):
        """
        Click "Select File",
        Display Selected File,
        Open Workbook as self.wb
        """
        try:
            filename = QFileDialog.getOpenFileName(self, 'Open file', './data')
            self.lineEdit_SelectedFile.setText(filename[0]) 
            
            self.df = pd.read_excel(filename[0], index_col=0)
            self.wb = xl.load_workbook(filename = filename[0])
            # 각 시트를 변수에 할당
            self.ws = self.wb.active
        
        except Exception as e:
            print(e)

    def get_year_month(self):
        """
        """
        if self.lineEdit_Year.text() and self.lineEdit_Month.text():
            self.year = int(self.lineEdit_Year.text())
            self.month = int(self.lineEdit_Month.text())
        else:
            self.message_list.append(f"Set Year, Month")
        
    def get_fromto(self):
        """
        시작 행 번호부터 마지막 행 번호를
        From #, To #에 각각 적고 그 값을 가져온다.
        """
        self.number_from = int(self.lineEdit_From.text())
        self.number_to = int(self.lineEdit_To.text())
        
        assert self.number_from <= self.number_to

    def set_model_to_view(self, df: pd.DataFrame):
        """
        Set Model to View.
        """
        # Set up the model
        # self.model = QSqlTableModel(self)
        self.model = MyTableModel(df)
        self.proxy_model = QSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.model)
        
        
        self.model.setTable("E-Mail List")
        self.model.setEditStrategy(models.QSqlTableModel.OnFieldChange)
        
        self.model.setHeaderData(-1, Qt.Horizontal, "No.")
        self.model.setHeaderData(0, Qt.Horizontal, "ID")
        self.model.setHeaderData(1, Qt.Horizontal, "Name")
        self.model.setHeaderData(2, Qt.Horizontal, "E-Mail")
        self.model.select()
        
        # Set up the view
        # self.tableView = QTableView()
        
        # self.tableView.setModel(self.model)
        self.tableView.setModel(self.proxy_model)
        self.tableView.setSortingEnabled(True)
        self.tableView.resizeColumnsToContents()
        # self.setCentralWidget(self.tableView)

    
    def preview_mail_list(self):
        self.email_list = []
        for r in self.ws.rows:
            print(f"r: {r}, ws.rows: {self.ws.rows}")
            if r[0].value is None:
                print(f"r[0]: {r[0]}, r[0].value: {r[0].value}")
                continue
            self.email_list.append([])
            for c in r:
                print(f"c: {c}, r: {r}")
                self.email_list[-1].append(c.value)
                print(f"self.email_list in for loop: {self.email_list}")
                print(f"self.email_list[-1] in for loop: {self.email_list[-1]}")
            print(f"self.email_list[-1]: {self.email_list[-1]}")
        print(f"self.email_list before pop: {self.email_list}")
        self.email_list.pop(0)
        print(f"self.email_list after pop: {self.email_list}")
        
        try:
            self.get_year_month()
            self.get_fromto()
            self.mailListLabel.setText(
                f"""
From {self.number_from} To {self.number_to}
{self.df.iloc[self.number_from-1:self.number_to]}
                """)
#             self.mailListLabel.setText(
#                 f"""
# From {self.number_from} To {self.number_to}
# {self.email_list}
#                 """)
            self.message_list.append(f"Worked well.")
        except AssertionError:
            self.message_list.append(f"The number for 'To' must be greater than or equal to 'From'")
            # self.messages.setText(f"The number for 'To' must be greater than or equal to 'From'")
        except:
            self.message_list.append(f"Set From, To")
            # self.messages.setText(f"Set From, To")
        
        self.messages.setText(str(self.message_list))
        self.message_list = []
        
        self.set_model_to_view(self.df)
    
    def select_pdf_folder(self):
        self.selected_folderpath = QFileDialog.getExistingDirectory(self, 'Select Folder')
        self.lineEdit_SelectedFolder.setText(self.selected_folderpath) 
    
    def send_emails(self):
        dirname = os.path.dirname(os.path.abspath(__file__))
        write_btn_rel_path = 'data\\write_mail.jpg'
        send_btn_rel_path = 'data\\btn_send.jpg'
        attach_btn_rel_path = 'data\\btn_attach.jpg'
        write_btn_path = os.path.join(dirname, write_btn_rel_path)
        send_btn_path = os.path.join(dirname, send_btn_rel_path)
        attach_btn_path = os.path.join(dirname, attach_btn_rel_path)
        print(write_btn_path)
        print(send_btn_path)
        
        for i in self.email_list:
            write_btn_location = pyautogui.locateOnScreen(write_btn_path, confidence = 0.8)
            print(f"write_btn_location: {write_btn_location}")
            write_btn_center = pyautogui.center(write_btn_location)
            pyautogui.click(write_btn_center[0], write_btn_center[1])
            time.sleep(3)
            clp.copy(i[-1]) # email address
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(3)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            clp.copy(f"{self.year}년 {self.month}월 급여명세서") # name
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.hotkey('tab')
            clp.copy(
            f'''
수고하셨습니다.
''')
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(3)
            
            attach_btn_location = pyautogui.locateOnScreen(attach_btn_path, confidence = 0.8)
            attach_btn_center = pyautogui.center(attach_btn_location)
            pyautogui.click(attach_btn_center[0], attach_btn_center[1])
            time.sleep(3)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
            time.sleep(3)
            clp.copy(self.selected_folderpath)
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.hotkey('enter')
            time.sleep(3)
            pyautogui.hotkey('tab')
            time.sleep(1)
            pyautogui.hotkey('tab')
            time.sleep(1)
            pyautogui.hotkey('tab')
            time.sleep(1)
            pyautogui.hotkey('tab')
            time.sleep(1)
            pyautogui.hotkey('tab')
            time.sleep(1)
            pyautogui.hotkey('tab')
            time.sleep(3)
            
            clp.copy(f"{i[1]}.{i[2]}.pdf") ##### id.name.pdf
            print(f"{i[1]}.{i[2]}.pdf copied")
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')     
            
            pyautogui.hotkey('enter')
            time.sleep(3)
            
            send_btn_location = pyautogui.locateOnScreen(send_btn_path, confidence = 0.8)
            print(f"send_btn_location: {send_btn_location}")
            send_btn_center = pyautogui.center(send_btn_location)
            pyautogui.click(send_btn_center[0], send_btn_center[1])
            time.sleep(10)
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WindowClass()
    window.show()
    sys.exit(app.exec_())