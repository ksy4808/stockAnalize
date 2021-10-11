import sys
import time
import threading
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QPalette

import win32com.client
import ctypes
import pandas as pd


g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')



#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("Deasin.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    
    def __init__(self) :
        super().__init__()
        

        #self.worker = Qworker()
        #self.worker.start()
        #self.worker.timeout.connect(self.timeout)   # 시그널 슬롯 등록

        self.setupUi(self)
        pal = QPalette()
        self.LogonCheck.clicked.connect(self.btn_login)


    def timeout(self, num):
        num+=1
    def writeLoginStatus(self, status):
        self.LoginStatus.setText(status)
    
    def btn_login(self):
        if (g_objCpStatus.IsConnect == 0):
            self.writeLoginStatus("Disconnected")
        else:
            self.writeLoginStatus("Connected")
            self.worker = GetItems()
            self.worker.start()
            self.worker.checked.connect(self.dispTotalItemsTable)   # 시그널 슬롯 등록

    def dispTotalItemsTable(Items):
        #self.totalItemsTable.setho
        i=1

class GetItems(QThread):
    checked = pyqtSignal(tuple)

    def __init__(self):
        super().__init__()

    def run(self):
        i=1
        g_codeList = g_objCodeMgr.GetStockListByMarket(1)
        self.checked.emit(g_codeList)     # 방출




class Qworker(QThread):
    timeout = pyqtSignal(int)
    def __init__(self):
        super().__init__()
        self.num = 0             # 초깃값 설정
    def run(self):
        i=0
        while(True):
            self.timeout.emit(self.num)     # 방출
            print(str(i))
            i += 1
            time.sleep(1)

if __name__ == "__main__" :
    global app
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 
    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 
    #프로그램 화면을 보여주는 코드
    myWindow.show()
    app.exec_()





