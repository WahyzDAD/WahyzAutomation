import sys
import time

from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import QtGui
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import  QWidget, QLabel, QApplication
from PyQt5.QtCore import QThread, Qt, pyqtSignal, pyqtSlot,  QRunnable, QThreadPool
from PyQt5.QtGui import QImage, QPixmap

import salarySpecDM as model

form_class = uic.loadUiType("SalarySpec.ui")[0]

class WindowClass(QMainWindow, form_class) :
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.excel = model.initExcel()
        self.excel.Visible = True
        
        # Buttons
        self.btn_SelectFile.clicked.connect(self.select_file)
        self.btn_SaveDir.clicked.connect(self.select_save_dir)
        self.btn_MakePDFs.clicked.connect(self.make_pdfs)
        
    def select_file(self):
        """
        Click "Select File",
        Display Selected File,
        Open Workbook as self.wb
        """
        filename = QFileDialog.getOpenFileName(self, 'Open file', './')
        self.lineEdit_SelectedFile.setText(filename[0]) 
        
        self.wb = self.excel.Workbooks.Open(filename[0])
        
        # 각 시트를 변수에 할당
        self.ws_basisTable = self.wb.Worksheets("1.기본정보TABLE")
        self.ws_ManagerSalary = self.wb.Worksheets("급여명세표_관리")
        
    def select_save_dir(self):
        dirname = QFileDialog.getExistingDirectory(self, 'Select Directory')
        self.lineEdit_SaveDir.setText(dirname)
        self.dirname = dirname

    def get_fromto(self):
        """
        엑셀 "1.기본정보TABLE"에서 급여명세표 pdf 변환을 할 사원 리스트의 
        시작 행 번호부터 마지막 행 번호를
        From #, To #에 각각 적고 그 값을 가져온다.
        """
        self.number_from = int(self.lineEdit_From.text())
        self.number_to = int(self.lineEdit_To.text())
        
    def make_pdfs(self):
        """
        pdf 변환을 시작한다. 선택된 사원들에 대한 각각의 pdf가 생성된다.
        """
        self.get_fromto()
        self.ws_ManagerSalary.Select()
        time.sleep(1)

        # # 자동 필터된 범위
        # rangeofInterest = ws_basisTable.Autofilter.Range.Address

        # if rangeofInterest.Item("구분") == "관리" or "연구소":
        #     pass

        # 1~4 출력번호 반복문
        for i in range(self.number_from, self.number_to+1):
            # 출력 번호 2행1열에 입력
            self.ws_ManagerSalary.Cells(3, 2).Value = self.ws_basisTable.Cells(i, 2).Value
            # 3행2열 값을 name 변수에 저장
            name = self.ws_ManagerSalary.Cells(3, 2).Value
            # 사번 저장
            idNumber = self.ws_ManagerSalary.Cells(3, 4).Value
            # pdf 저장경로, 파일명
            pdf_path = "{}\\{}.{}.pdf".format(self.dirname, int(idNumber), name)
            # pdf 저장
            self.wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
            time.sleep(1)
        
        self.wb.Close(False)  
        self.excel.Quit()

if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()
    
    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()