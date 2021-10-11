import sys
import os
import openpyxl
import sqlite3
import datetime
import threading
import time
from openpyxl import Workbook
from random import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QAxContainer import *
from time import sleep



form_class = uic.loadUiType("DBGetAll.ui")[0]  # ui 파일을 로드하여 form_class 생성

#sem0 = threading.Semaphore(1)


class superLoopFlag():
    global hasSomeThingTodo
    someThingTodo = 0
    global whatThingTodo
    whatThingTodo = 0#kospi 체결정보 read면 1 kosdaq체결정보 read면 2

class MyWindow(QMainWindow, form_class):  # MyWindow 클래스 QMainWindow, form_class 클래스를 상속 받아 생성됨

    global g_date
    g_date = "1"#1이면 당일 2이면 전일
    global g_code
    g_code = 0
    global rowOffset# 당일 체결정보를 전체 Item에 대하여 재귀적으로 진행할 경우 offset을 기억하기 위한 변수
    rowOffset = 0
    global getContinue# 당일 체결정보를 전체 Item에 대하여 재귀적으로 진행할 경우 set.
    getContinue = 0
    global maxLoop
    maxLoop = 0
    global KospiKosdaq#당일 체결정보를 읽는데 현재 kospi인지 kosdaq인지 구분 0이면 kospi, 1이면 kosdaq
    KospiKosdaq = 0
    global superLoopFlag
    global nextConclutionCnt
    nextConclutionCnt = 0
    global delay_gap_cnd
    delay_gap_cnd = 0


    def __init__(self):  # MyWindow 클래스의 초기화 함수(생성자)
        super().__init__()  # 부모클래스 QMainWindow 클래스의 초기화 함수(생성자)를 호출
        self.setupUi(self)  # ui 파일 화면 출력

        self.kiwoom = QAxWidget(
            "KHOPENAPI.KHOpenAPICtrl.1")  # 키움증권 Open API+의 ProgID를 사용하여 생성된 QAxWidget을 kiwoom 변수에 할당
        self.AllbuttonDisable()

        self.btnLogin.clicked.connect(self.btn_login)  # ui 파일을 생성할때 작성한 로그인 버튼의 objectName 으로 클릭 이벤트가 발생할 경우 btn_login 함수를 호출
        self.btnSearch2.clicked.connect(
            self.btn_search2)  # ui 파일을 생성할때 작성한 조회 버튼의 objectName 으로 클릭 이벤트가 발생할 경우 btn_search 함수를 호출
        self.btnSearch.clicked.connect(self.btn_search)  # ui 파일을 생성할때 작성한 조회 버튼의 objectName 으로 클릭 이벤트가 발생할 경우 btn_search 함수를 호출
        self.btnGetAllItems.clicked.connect(self.btn_GetAllItems) # 모든 종목 get함수 호출

        self.btnSearch_kospi.clicked.connect(self.btn_Search_kospi)  # 코스피 당일체결조회 버튼 클릭 시 btn_Search_kospi함수 호출


        self.kiwoom.OnEventConnect.connect(self.event_connect)  # 키움 서버 접속 관련 이벤트가 발생할 경우 event_connect 함수 호출
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)  # 키움 데이터 수신 관련 이벤트가 발생할 경우 receive_trdata 함수 호출

        #self.superLoop()
    def superLoop(self):
        global superLoopFlag
        while True:
            sleep(1)
            if superLoopFlag.someThingTodo == 1:
                superLoopFlag.someThingTodo = 0
                break



    def btn_GetAllItems(self): # 모든종목 get함수
        retKospi = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ["0"]) # 0이면 코스피, 10이면 코스닥
        retKosdaq = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ["10"])  # 0이면 코스피, 10이면 코스닥
        kospi_code_list = retKospi.split(';')
        kosdaq_code_list = retKosdaq.split(';')



        con = sqlite3.connect("Items.db")
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS KosdaqItems(num text PRIMARY KEY, name text)")
        cur.execute("CREATE TABLE IF NOT EXISTS KospiItems(num text PRIMARY KEY, name text)")

        kospi_code_name_list = []
        for code in kospi_code_list:
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [code])  # 맨뒤는 종목코드, 코드에 따른 종목명을 가져옴
            kospi_code_name_list.append(code + " : " + name)
            self.pteLog.appendPlainText(code + " " + ":" + name)
            sql = "insert or ignore into KospiItems(num,name) values(?, ?)"
            cur.execute(sql, (code.strip(), name.strip()))

        kosdaq_code_name_list = []
        for code in kosdaq_code_list:
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [code]) # 맨뒤는 종목코드, 코드에 따른 종목명을 가져옴
            kosdaq_code_name_list.append(code + " : " + name)
            self.pteLog.appendPlainText(code + " " + ":" + name)
            sql = "insert or ignore into KosdaqItems(num,name) values(?, ?)"
            cur.execute(sql, (code.strip(), name.strip()))
        con.commit()
        con.close()


    def btn_login(self):  # Login 버튼 클릭 시 실행되는 함수
        ret = self.kiwoom.dynamicCall("CommConnect()")  # 키움 로그인 윈도우를 실행

    def event_connect(self, err_code):  # 키움 서버 접속 관련 이벤트가 발생할 경우 실행되는 함수
        if err_code == 0:  # err_code가 0이면 로그인 성공 그외 실패
            self.AllbuttomEnable()
            self.pteLog.appendPlainText("로그인 성공")  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함

            account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"])  # 키움 dynamicCall 함수를 통해 GetLoginInfo 함수를 호출하여 계좌번호를 가져옴
            self.pteLog.appendPlainText("계좌번호: " + account_num.rstrip(';'))  # 키움은 전체 계좌를 반환하며 각 계좌 번호 끝에 세미콜론(;)이 붙어 있음으로 제거하여 plainTextEdit에 텍스트를 추가함
            self.btnLogin.setDisabled(True)  # 로그인 버튼을 비활성화 상태로 변경
        else:
            self.AllbuttonDisable()
            self.pteLog.appendPlainText("로그인 실패")  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
            self.btnLogin.setDisabled(False)  # 로그인 버튼을 활성화 상태로 변경

    def btn_search(self):  # 조회 버튼 클릭 시 실행되는 함수
        code = self.lineEditCode.text()  # ui 파일을 생성할때 작성한 종목코드 입력란의 objectName 으로 사용자가 입력한 종목코드의 텍스트를 가져옴
        self.pteLog.appendPlainText("종목코드: " + code)  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")  # 키움 dynamicCall 함수를 통해 CommRqData 함수를 호출하여 opt10001 API를 구분명 opt10001_req, 화면번호 0101으로 호출함
        self.AllbuttonDisable()

    def btn_search2(self):  # 조회 버튼 클릭 시 실행되는 함수
        global g_code
        global getContinue
        date = "1"# 1인경우 당일, 2이면 전일

        code = self.lineEditCode.text()  # ui 파일을 생성할때 작성한 종목코드 입력란의 objectName 으로 사용자가 입력한 종목코드의 텍스트를 가져옴
        if code == "":
            con = sqlite3.connect("Items.db")
            cur = con.cursor()
            cur.execute("select * from KosdaqItems;")
            #for row in cur:
            #    sem0.acquire()
            row = cur.fetchone()
            con.close()
            code = row[0]
            getContinue = 1
        g_code = code
        self.pteLog.appendPlainText("종목코드: " + code)  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
        self.opt10055_req(code, date, 0)
        """
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "당일전일", date)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10055_req", "opt10055", 0, "0101")  # 키움 dynamicCall 함수를 통해 CommRqData 함수를 호출하여 opt10055 API를 구분명 opt10055_req, 화면번호 0101으로 호출함
        """
        self.AllbuttonDisable()

    def btn_Search_kospi(self):  # 조회 버튼 클릭 시 실행되는 함수
        global g_code
        global getContinue
        global rowOffset
        global maxLoop
        global KospiKosdaq
        KospiKosdaq = 0

        rowOffset = 0
        getContinue = 1#연속적으로 읽을지 여부를 전달하는 메시지

        while True:
            con = sqlite3.connect("Items.db")
            cur = con.cursor()
            if KospiKosdaq == 1:
                cur.execute("select * from KosdaqItems;")
            elif KospiKosdaq == 0:
                cur.execute("select * from KospiItems;")
            rows = cur.fetchall()
            maxLoop = len(rows) - 1
            con.close()


            for row in rows:
                if (row[2] != None) and (row[2] != 0):
                    code = row[0]
                    g_code = code
                    if True == self.reqAftercExistionChk(rowOffset, code):  # 이미 gettering한 종목인지 확인하고 아니면 API요청후 return
                        return
                    else:  # 이미 gettering한 종목이면 다음 row를 확인하기 위해 계속 진행
                        rowOffset += 1
                        continue
                rowOffset += 1

                if rowOffset >= maxLoop and KospiKosdaq == 0:  # 최대 루프만큼 진행했으면 현재 kospi인지 확인해서 kospi이면 kosdaq종목으로 넘어간다.
                    getContinue = 1
                    KospiKosdaq = 1
                    rowOffset = 0  # kosdaq으로 옮겼으니 0번 옵셋부터 시작.
                    break  # 코스피로 옮겼으니 현재 반복문을 나와서 상위 while문에서 kosdaq DB를 읽어오도록 break
                elif rowOffset >= maxLoop and KospiKosdaq == 1:  # 최대 루프만큼 진행했고 현재 kosdaq이면 모두 진행 한 것임.
                    getContinue = 0
                    return


    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):  # 키움 데이터 수신 함수
        global g_code
        global getContinue
        global rowOffset
        global KospiKosdaq
        global superLoopFlag
        global nextConclutionCnt
        global g_date
        if rqname == "opt10001_req":  # 수신된 데이터 구분명이 opt10001_req 일 경우
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")  # 구분명 opt10001_req 의 종목명을 가져와서 name에 셋팅
            volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "거래량")  # 구분명 opt10001_req 의 거래량을 가져와서 volume에 셋팅

            self.pteLog.appendPlainText("종목명: " + name.strip())  # 종목명을 공백 제거해서 plainTextEdit에 텍스트를 추가함
            self.pteLog.appendPlainText("거래량: " + volume.strip())  # 거래량을 공백 제거해서 plainTextEdit에 텍스트를 추가함

            self.AllbuttomEnable()
        if rqname == "opt10055_req":  # 수신된 데이터 구분명이 opt10055_req 일 경우
            maxRepeatCnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
            self.pteLog.appendPlainText("Repeat Cnt: {}".format(maxRepeatCnt))
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", str(g_code))  # 맨뒤는 종목코드, 코드에 따른 종목명을 가져옴
            name = self.nameRefair(name)
            name = str(name)

            today = self.getToday()

            con = sqlite3.connect("../../conclusions/" + today + ".db")
            cur = con.cursor()
            cur.execute(
                "CREATE TABLE IF NOT EXISTS " + name + "(number integer PRIMARY KEY, conclusionTime text, price integer, quantity integer, conPrice integer)")

            for i in range(0,maxRepeatCnt):

                time = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "체결시간")  # 체결시간에 해당하는 string을 read
                concludePrice = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "체결가")  # 체결가격에 해당하는 string을 read
                concludeQuantity = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "체결량")  # 체결량에 해당하는 string을 read

                concludePrice = int(concludePrice.strip())
                concludeQuantity = int(concludeQuantity.strip())
                sql = "insert or ignore into  "+ name + "(number,conclusionTime,price,quantity,conPrice) values(?, ?, ?, ?, ?)"
                cur.execute(sql, (i+nextConclutionCnt, time.strip(), int(abs(concludePrice)), concludeQuantity, int(abs(concludePrice)*concludeQuantity)))

            con.commit()
            con.close()
            if(1000 == maxRepeatCnt):
                #superLoopFlag.someThingTodo = 1
                nextConclutionCnt += 1000
                date = g_date  # 1인경우 당일, 2이면 전일
                #code = self.lineEditCode.text()  # ui 파일을 생성할때 작성한 종목코드 입력란의 objectName 으로 사용자가 입력한 종목코드의 텍스트를 가져옴
                code = g_code
                self.opt10055_req(code, date, 2)

            else:
                nextConclutionCnt = 0
                if getContinue == 1:
                    rowOffset = rowOffset+1
                    self.ReadNextItemConclusion(rowOffset)
    def getToday(self):
        todayStr = self.todaySetting.text()
        if todayStr == "":
            today = datetime.date.today()
            date = today.weekday()

            if (5==date):
                today += datetime.timedelta(days=-1)
            elif (6==date):
                today += datetime.timedelta(days=-2)

            todayStr = str(today)
            todayStr = todayStr.replace('-', '_')

        return todayStr
    def nameRefair(self, name):
        name = name.replace('&', 'n')
        name = name.replace('-', '_')
        name = name.replace(' ', '_')
        name = name.replace('.', '_')
        name = name.replace('0', '_0')
        name = name.replace('1', '_1')
        name = name.replace('2', '_2')
        name = name.replace('3', '_3')
        name = name.replace('4', '_4')
        name = name.replace('5', '_5')
        name = name.replace('6', '_6')
        name = name.replace('7', '_7')
        name = name.replace('8', '_8')
        name = name.replace('9', '_9')
        name = name.replace('(', '_')
        name = name.replace(')', '_')
        name = name.replace('%', '_')
        name = name.replace('*', '_')
        name = name.replace('@', '_')
        name = name.replace('!', '_')
        name = name.replace('+', '_')
        name = name.replace('=', '_')
        name = name.replace('~', '_')
        name = name.replace('/', '_')
        return name
    def ReadNextItemConclusion(self, offset):
        global g_date
        global getContinue
        global g_code
        global maxLoop
        global rowOffset
        global KospiKosdaq
        if offset >= maxLoop and KospiKosdaq == 0:  # 최대 루프만큼 진행했으면 현재 kospi인지 확인해서 kospi이면 kosdaq종목으로 넘어간다.
            getContinue = 1
            KospiKosdaq = 1
            rowOffset = 0
        elif offset >= maxLoop and KospiKosdaq == 1: #최대 루프만큼 진행했고 현재 kosdaq이면 모두 진행 한 것임.
            getContinue = 0
            return

        while True:

            con = sqlite3.connect("Items.db")
            cur = con.cursor()
            if KospiKosdaq == 1:
                cur.execute("select * from KosdaqItems;")
            elif KospiKosdaq == 0:
                cur.execute("select * from KospiItems;")
            rows = cur.fetchall()
            maxLoop = len(rows) - 1
            con.close()
            while True:
                row = rows[offset]
                if (row[2] != None) and (row[2] != 0):#현재 row가 관심종목이면
                    code = row[0]
                    if True == self.reqAftercExistionChk(offset, code):#이미 gettering한 종목인지 확인하고 아니면 API요청후 return

                        return
                    else:#이미 gettering한 종목이면 다음 row를 확인하기 위해 계속 진행
                        offset += 1
                        continue
                offset += 1
                if offset >= maxLoop and KospiKosdaq == 0:  # 최대 루프만큼 진행했으면 현재 kospi인지 확인해서 kospi이면 kosdaq종목으로 넘어간다.
                    getContinue = 1
                    KospiKosdaq = 1
                    offset = 0# kosdaq으로 옮겼으니 0번 옵셋부터 시작.
                    break#코스피로 옮겼으니 현재 while을 나와서 상위 while문에서 kosdaq DB를 읽어오도록 break
                elif offset >= maxLoop and KospiKosdaq == 1:  # 최대 루프만큼 진행했고 현재 kosdaq이면 모두 진행 한 것임.
                    getContinue = 0
                    return
    def reqAftercExistionChk(self, offset, code):
        global g_code
        rowOffset = offset
        g_code = code

        name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", str(g_code))  # 맨뒤는 종목코드, 코드에 따른 종목명을 가져옴
        name = self.nameRefair(name)
        name = str(name)

        today = self.getToday()
        con = sqlite3.connect("../../conclusions/" + today + ".db")
        cur = con.cursor()
        cur.execute(" SELECT count(name) FROM sqlite_master WHERE type='table' AND name='%s' " % name)
        exsist  = cur.fetchone()[0]
        con.close()

        if exsist == 0:#1이면 table이 있음. 0이면 table이 없음 없으므로 API요청하고 return
            print("현재 offset: {}/{}, 종목 명: {}".format(offset, maxLoop, name))
            self.pteLog.appendPlainText("현재 offset: {}/{}, 종목 명: {}".format(offset, maxLoop, name))  #
            self.pteLog.appendPlainText("종목코드: " + code)  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
            self.opt10055_req(code, g_date, 0)
            self.AllbuttonDisable()

            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", str(g_code))
            return True
        else:# table이 있으면 0을 return
            print("현재 offset: {}/{}, 종목 명: {}, exist!!".format(offset, maxLoop, name))
            self.pteLog.appendPlainText("현재 offset: {}/{}, 종목 명: {}, exist!!".format(offset, maxLoop, name))  #
            return False
    def opt10055_req(self, code, date, repeat):
        global delay_gap_cnd
        delay_gap_cnd += 1
        if delay_gap_cnd > 50:
            sleep(20)
            delay_gap_cnd = 0
        self.pteLog.appendPlainText(
            "종목코드: " + code + "  " + str(delay_gap_cnd))  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
        sleep(3.6)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드",
                                code)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "당일전일",
                                date)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10055_req", "opt10055", repeat,
                                "0101")  # 키움 dynamicCall 함수를 통해 CommRqData 함수를 호출하여 opt10055 API를 구분명 opt10055_req, 화면번호 0101으로 호출함
    def AllbuttonDisable(self):
        self.lineEditCode.setDisabled(True)  # 종목코드 입력란을 비활성화 상태로 변경
        self.btnSearch2.setDisabled(True)  # 당일 체결정보 조회 버튼을 비활성화 상태로 변경
        self.btnSave.setDisabled(True)  # 종목정보 저장 버튼을 비활성화 상태로 변경
        self.btnSave2.setDisabled(True)  # 당일 체결정보 저장 버튼을 비활성화 상태로 변경
        self.btnSearch.setDisabled(True)  # 종목정보 조회 버튼을 비활성화 상태로 변경
        self.pteLog.setDisabled(True)  # plainTextEdit를 비활성화 상태로 변경
        self.btnGetAllItems.setDisabled(True) # 전체종목 get 버튼 비활성화
        self.btnSearch_kospi.setDisabled(True)  # 코스피 당일체결내역 조회 버튼비활성화
        self.todaySetting.setDisabled(True) # 날짜 지정 text line 비활성화
    def AllbuttomEnable(self):
        self.lineEditCode.setDisabled(False)  # 종목코드 입력란을 비활성화 상태로 변경
        self.btnSearch2.setDisabled(False)  # 당일 체결정보 조회 버튼을 비활성화 상태로 변경
        self.btnSave.setDisabled(False)  # 종목정보 저장 버튼을 비활성화 상태로 변경
        self.btnSave2.setDisabled(False)  # 당일 체결정보 저장 버튼을 비활성화 상태로 변경
        self.btnSearch.setDisabled(False)  # 종목정보 조회 버튼을 비활성화 상태로 변경
        self.pteLog.setDisabled(False)  # plainTextEdit를 비활성화 상태로 변경
        self.btnGetAllItems.setDisabled(False) # 전체종목 get 버튼 비활성화
        self.btnSearch_kospi.setDisabled(False)  # 코스피 당일체결내역 조회 버튼활성화
        self.todaySetting.setDisabled(False) # 날짜 지정 text line 활성화


# py 파일 실행시 제일 먼저 동작
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()  # MyWindow 클래스를 생성하여 myWondow 변수에 할당
    myWindow.show()  # MyWindow 클래스를 노출
    app.exec_()  # 메인 이벤트 루프에 진입 후 프로그램이 종료될 때까지 무한 루프 상태 대기