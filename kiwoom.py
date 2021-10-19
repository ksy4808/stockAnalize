import sys
import time
import threading
import pydevd
import sqlite3
import datetime as dt
from datetime import date
import matplotlib.patches as mpatches
import matplotlib.ticker as ticker
from matplotlib import pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import numpy as np
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QPalette
from PyQt5.QAxContainer import *

import pandas_datareader.data as web
from pandas import Series, DataFrame
import pandas as pd
import plotly.graph_objects as go

import constant as const
# 한글 폰트 사용을 위해서 세팅
from matplotlib import font_manager, rc
font_path = "C:/Windows/Fonts/NGULIM.TTF"
font = font_manager.FontProperties(fname=font_path).get_name()
rc('font', family=font)

const.codeLine = 0
const.nameLine = 1





import pandas as pd



#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("kiwoom.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    
    def __init__(self) :
        super().__init__()
        
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)  # 키움 데이터 수신 관련 이벤트가 발생할 경우 receive_trdata 함수 호출

        #self.worker = Qworker()
        #self.worker.start()
        #self.worker.timeout.connect(self.timeout)   # 시그널 슬롯 등록

        self.setupUi(self)
        pal = QPalette()

        self.plot = plt
        self.fig = self.plot.Figure()
        self.canvas = FigureCanvas(self.fig)
        self.layoutForPlot.addWidget(self.canvas)

        self.LogonCheck.clicked.connect(self.btn_login)
        self.kiwoom.OnEventConnect.connect(self.event_connect)  # 키움 서버 접속 관련 이벤트가 발생할 경우 event_connect 함수 호출
        self.Add.clicked.connect(self.btn_add)
        self.Sub.clicked.connect(self.btn_sub)
        self.up.clicked.connect(self.btn_up)
        self.down.clicked.connect(self.btn_down)
        self.comboBox.currentIndexChanged.connect(self.procFunc)
        self.ItemSearch.returnPressed.connect(self.ItemSearchEnter)
        self.intListTable.itemClicked.connect(self.procFunc)
        self.defaultUISetup()
    
    def ItemSearchEnter(self):
        for i in list(range(0, self.totalListTable.rowCount()-1, 1)):
            name = self.totalListTable.item(i, 1).text()
            if name == self.ItemSearch.text():
                self.totalListTable.setCurrentCell(i,1)
                break

    def defaultUISetup(self):
        self.ItemSearch.setStyleSheet("color: black;")#lineEdit 오브젝트의 글자 색상 변경
        self.Add.setDisabled(True)
        self.ItemSearch.setDisabled(True)
        self.totalListTable.setDisabled(True)
        self.Sub.setDisabled(True)
        self.intListTable.setDisabled(True)
        self.comboBox.setDisabled(True)
        self.OutputTable.setDisabled(True)

        ########## 시간 설정
        defaultToTime = QTime(15,30,0)
        defaultFromTime = QTime(9,0,0)
        self.toTime.setTime(defaultToTime)
        self.fromTime.setTime(defaultFromTime)
        ########## 날짜 설정
        defaultToDay = QDate(QDate.currentDate())
        defaultFromDay = QDate(defaultToDay.addDays(-90))
        self.toDay.setDate(defaultToDay)
        self.fromDay.setDate(defaultFromDay)
    
    def timeout(self, num):
        num+=1
    
    def writeLoginStatus(self, status):
        self.LoginStatus.setText(status)
    
    def btn_add(self):
        x = self.totalListTable.selectedIndexes()#선택된 셀의 행/열 번호가 반환된다.
        code = self.totalListTable.item(x[0].row(), 0).text()
        name = self.totalListTable.item(x[0].row(), 1).text()
        category = self.totalListTable.item(x[0].row(), 2).text()
        exist = False
        for i in list(range(0, self.intListTable.rowCount(), 1)):
            if code == self.intListTable.item(i,0).text():
                exist = True
                break
        if exist == False:
            intListTableRowCnt = self.intListTable.rowCount()
            self.intListTable.setRowCount(intListTableRowCnt+1)
            self.intListTable.setItem(intListTableRowCnt, 0, QTableWidgetItem(code))
            self.intListTable.setItem(intListTableRowCnt, 1, QTableWidgetItem(name))
            self.intListTable.setItem(intListTableRowCnt, 2, QTableWidgetItem(category))
            self.dbAdd("interrestLists.db", "list", "code", "name", "category", code, name, category)

    def btn_sub(self):
        x = self.intListTable.selectedIndexes()#선택된 셀의 행/열 번호가 반환된다.
        code = self.intListTable.item(x[0].row(), 0).text()
        name = self.intListTable.item(x[0].row(), 1).text()
        category = self.totalListTable.item(x[0].row(), 2).text()
        if len(x) != 0:
            row = x[0].row()
            self.intListTable.removeRow(row)
            self.dbSub("interrestLists.db", "list", "code", "name", "category", code, name, category)

    def dbAdd(self, dbName, tableName, field1, field2, field3, value1, value2, value3):
        con = sqlite3.connect(dbName)
        cur = con.cursor()
        sql = "insert or ignore into "+tableName+"("+field1+","+field2+", "+field3+") values(?,?,?)"
        cur.execute(sql, (value1, value2, value3))
        con.commit()
        con.close()

    def dbSub(self, dbName, tableName, field1, field2, field3, value1, value2, value3):
        con = sqlite3.connect(dbName)
        cur = con.cursor()
        sql = "delete from "+tableName+" where "+field1+" = '"+value1+"'"
        cur.execute(sql)
        con.commit()
        con.close()

    def dbUp(self, rowIndex):
        con = sqlite3.connect("interrestLists.db")
        cur = con.cursor()
        cur.execute("select * from list;")
        rows = cur.fetchall()
        con.close()
        code = rows[rowIndex][0]
        name = rows[rowIndex][1]
        category = rows[rowIndex][2]
        upperCode = rows[rowIndex-1][0]
        upperName = rows[rowIndex-1][1]
        upperCategory = rows[rowIndex-1][2]
        lstRows = list(rows)
        curLst = lstRows[rowIndex]
        upperLst = lstRows[rowIndex-1]
        lstRows[rowIndex] = upperLst
        lstRows[rowIndex-1] = curLst
        rows = tuple(lstRows)

        con = sqlite3.connect("interrestLists.db")
        cur = con.cursor()
        cur.execute("delete from 'list'")
        cur.execute("CREATE TABLE IF NOT EXISTS list(code text PRIMARY KEY, name text, category text)")
        cur.executemany("insert into 'list'(code, name, category) values(?,?,?)", rows)
        con.commit()
        con.close()

    def dbDown(self, rowIndex):
        con = sqlite3.connect("interrestLists.db")
        cur = con.cursor()
        cur.execute("select * from list;")
        rows = cur.fetchall()
        con.close()
        code = rows[rowIndex][0]
        name = rows[rowIndex][1]
        category = rows[rowIndex][2]
        upperCode = rows[rowIndex-1][0]
        upperName = rows[rowIndex-1][1]
        upperCategory = rows[rowIndex-1][2]
        lstRows = list(rows)
        curLst = lstRows[rowIndex]
        lowerLst = lstRows[rowIndex+1]
        lstRows[rowIndex] = lowerLst
        lstRows[rowIndex+1] = curLst
        rows = tuple(lstRows)

        con = sqlite3.connect("interrestLists.db")
        cur = con.cursor()
        cur.execute("delete from 'list'")
        cur.execute("CREATE TABLE IF NOT EXISTS list(code text PRIMARY KEY, name text, category text)")
        cur.executemany("insert into 'list'(code, name, category) values(?,?,?)", rows)
        con.commit()
        con.close()

    def btn_up(self):
        x = self.intListTable.selectedIndexes()#선택된 셀의 행/열 번호가 반환된다.
        code = self.intListTable.item(x[0].row(), 0).text()
        name = self.intListTable.item(x[0].row(), 1).text()
        category = self.intListTable.item(x[0].row(), 2).text()
        if len(x) != 0 and x[0].row() > 0:#선택된 셀이 있고 최상위 셀이 아닌 경우에만 실행
            upperRow = x[0].row()-1
            upperCode = self.intListTable.item(upperRow, 0).text()
            upperName = self.intListTable.item(upperRow, 1).text()
            upperCategory = self.intListTable.item(upperRow, 2).text()
            self.intListTable.setItem(upperRow, 0, QTableWidgetItem(code))
            self.intListTable.setItem(upperRow, 1, QTableWidgetItem(name))
            self.intListTable.setItem(upperRow, 2, QTableWidgetItem(category))
            self.intListTable.setItem(x[0].row(), 0, QTableWidgetItem(upperCode))
            self.intListTable.setItem(x[0].row(), 1, QTableWidgetItem(upperName))
            self.intListTable.setItem(x[0].row(), 2, QTableWidgetItem(upperCategory))
            self.intListTable.setCurrentCell(upperRow,1)
            self.dbUp(x[0].row())

    def btn_down(self):
        x = self.intListTable.selectedIndexes()#선택된 셀의 행/열 번호가 반환된다.
        code = self.intListTable.item(x[0].row(), 0).text()
        name = self.intListTable.item(x[0].row(), 1).text()
        category = self.intListTable.item(x[0].row(), 2).text()
        if len(x) != 0 and x[0].row() < self.intListTable.rowCount():#선택된 셀이 있고 최하위 셀이 아닌 경우에만 실행
            lowerRow = x[0].row()+1
            lowerCode = self.intListTable.item(lowerRow, 0).text()
            lowerName = self.intListTable.item(lowerRow, 1).text()
            lowerCategory = self.intListTable.item(lowerRow, 2).text()
            self.intListTable.setItem(lowerRow, 0, QTableWidgetItem(code))
            self.intListTable.setItem(lowerRow, 1, QTableWidgetItem(name))
            self.intListTable.setItem(lowerRow, 2, QTableWidgetItem(category))
            self.intListTable.setItem(x[0].row(), 0, QTableWidgetItem(lowerCode))
            self.intListTable.setItem(x[0].row(), 1, QTableWidgetItem(lowerName))
            self.intListTable.setItem(x[0].row(), 2, QTableWidgetItem(lowerCategory))
            self.intListTable.setCurrentCell(lowerRow,1)
            self.dbDown(x[0].row())

    def btn_login(self):
        ret = self.kiwoom.dynamicCall("CommConnect()")
    
    def event_connect(self, err_code):  # 키움 서버 접속 관련 이벤트가 발생할 경우 실행되는 함수
        if err_code == 0:  # err_code가 0이면 로그인 성공 그외 실패
            self.writeLoginStatus("Connected")
            self.LogonCheck.setDisabled(True)  # 로그인 버튼을 비활성화 상태로 변경
            self.Add.setDisabled(False)
            self.ItemSearch.setDisabled(False)
            self.totalListTable.setDisabled(False)
            self.Sub.setDisabled(False)
            self.intListTable.setDisabled(False)
            self.comboBox.setDisabled(False)
            self.OutputTable.setDisabled(False)


            self.getCode = GetItems(self.kiwoom)
            self.getCode.procComplete.connect(self.dispTotalItemsTable)   # 시그널 슬롯 등록
            self.getCode.start()  

            self.confIntListTable() 
            self.confFuncList()     
        else:
            #self.AllbuttonDisable()
            self.writeLoginStatus("Not connected")  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
            self.LogonCheck.setDisabled(False)  # 로그인 버튼을 활성화 상태로 변경
    
    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):  # 키움 데이터 수신 함수
        #pydevd.connected = True
        #pydevd.settrace(suspend=False)
        if rqname == "opt10015_req":
            i=0
        elif rqname == "opt10059_req":
            maxRepeatCnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
            #maxRepeatCnt = 10
            listDate = []
            intListVol = []
            intListIndividual = []
            intListForeigner = []
            intListAgency = []
            intListCorporation = []
            intListOtherForeigner = []
            #dfBuyer = DataFrame()# 비어있는 DataFrame을 선언함.
            dfBuyer = DataFrame(columns=['vol', 'individual', 'foreigner', 'agency', 'corporation', 'otherForeigner'])

            for i in range(0, maxRepeatCnt):
                listBuyer = []
                strDate = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "일자").strip()
                strYear = strDate[2:4]
                strMonth = strDate[4:6]
                strDay = strDate[6:8]
                #listDate.append(date(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "일자").strip()))
                #dtDate = (date(int(strYear), int(strMonth), int(strDay)))
                dtDate = strYear+"."+strMonth+"."+strDay
                #dtDate = strYear + "-" + strMonth + "-" + strDay
                listBuyer.append(int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "누적거래대금").strip()))
                listBuyer.append(int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "개인투자자").strip()))
                listBuyer.append(int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "외국인투자자").strip()))
                listBuyer.append(int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "기관계").strip()))
                listBuyer.append(int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "기타법인").strip()))
                listBuyer.append(int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, i, "내외국인").strip()))

                #dfBuyer = dfBuyer.append(Series(listBuyer, index = dfBuyer.columns, name=dtDate))
                #tmpDfBuyer = DataFrame(columns = dfBuyer.columns)
                #tmpDfBuyer = tmpDfBuyer.append(Series(listBuyer, index = dfBuyer.columns, name=dtDate))
                dfBuyer = dfBuyer.append(Series(listBuyer, index = dfBuyer.columns, name=dtDate))

            #dfBuyer = self.calcBuyerAccumulation(dfBuyer,maxRepeatCnt)
            self.plotBuyer(dfBuyer,maxRepeatCnt)
    
    def calcBuyerAccumulation(self, dfBuyer, maxRepeatCnt):
        acc = 0
        #dfBuyer = dfBuyer.tail(maxRepeatCnt)
        individualAcc = 0
        foreignerAcc = 0
        agencyAcc = 0
        corporationAcc = 0
        otherForeignerAcc = 0
        for i in range(0, maxRepeatCnt):
            aDfBuyer = dfBuyer.iloc[maxRepeatCnt-1-i]
            aDfBuyer.loc['individual'] = individualAcc + aDfBuyer.loc['individual']
            individualAcc = aDfBuyer.loc['individual']
            aDfBuyer.loc['foreigner'] = foreignerAcc + aDfBuyer.loc['foreigner']
            foreignerAcc = aDfBuyer.loc['foreigner']
            aDfBuyer.loc['agency'] = agencyAcc + aDfBuyer.loc['agency']
            agencyAcc = aDfBuyer.loc['agency']
            aDfBuyer.loc['corporation'] = corporationAcc + aDfBuyer.loc['corporation']
            corporationAcc = aDfBuyer.loc['corporation']
            aDfBuyer.loc['otherForeigner'] = otherForeignerAcc + aDfBuyer.loc['otherForeigner']
            otherForeignerAcc = aDfBuyer.loc['otherForeigner']
            dfBuyer.iloc[maxRepeatCnt-1-i] = aDfBuyer
        return dfBuyer


    def plotBuyer(self, dfBuyer, maxRepeatCnt):
        dfBuyer = self.calcBuyerAccumulation(dfBuyer,maxRepeatCnt)
        #bar = dfBuyer.plot.line(grid=True)

        self.fig.clf(111)
        #self.fig.add_trace(go.Scatter(x=dfBuyer.index, y=dfBuyer['vol']))  #go로 생성한 경우 trace생성하는 법
        ax = self.fig.add_subplot(111)
        self.plot.show()
        ax2 = ax.twinx()

        ax2.bar(dfBuyer.index, dfBuyer.loc[:,'vol'], label="거래량", color='g')
        ax.plot(dfBuyer.index, dfBuyer.loc[:, 'individual'], label="개인순매수(누적)", color='y')
        ax.plot(dfBuyer.index, dfBuyer.loc[:,'foreigner'], label="외인순매수(누적)", color='b')
        ax.plot(dfBuyer.index, dfBuyer.loc[:, 'agency'], label="기관순매수(누적)", color='r')

        #ax.set_xticklabels(rows[dateIndex], rotation = 30)
        ax.set_xlabel("x_axis")
        ax.set_ylabel("y_axis")
        ax.set_title("my graph")
        #ax.invert_xaxis()
        #ax2.invert_xaxis()
        ax.grid()

        #ax.xaxis.set_major_locator(ticker.MultipleLocator(3))
        #ax.xaxis.set_minor_locator(ticker.MultipleLocator(3))
        ax.legend(loc = 'upper right')
        ax2.legend(loc = 'upper left')

        self.canvas.draw()
        
        i=1

    def confFuncList(self):
        con = sqlite3.connect("functionLists.db")
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS funcList(No text PRIMARY KEY, funcCode text, funcName text)")
        cur.execute("select * from funcList;")
        rows = cur.fetchall()
        con.close()
        for i, row in enumerate(rows):
            self.comboBox.addItem(row[2])

    def confIntListTable(self):
        column_headers = ['코드', '종목명', '업종코드']
        self.intListTable.setColumnCount(3)
        self.intListTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.intListTable.verticalHeader().hide()
        self.intListTable.showGrid()
        self.intListTable.setHorizontalHeaderLabels(column_headers)
        self.intListTable.setColumnHidden(2, True)
        con = sqlite3.connect("interrestLists.db")
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS list(code text PRIMARY KEY, name text, category text)")
        cur.execute("select * from list;")
        rows = cur.fetchall()
        con.close()
        for i, row in enumerate(rows):
            self.intListTable.setRowCount(self.intListTable.rowCount()+1)
            self.intListTable.setItem(i, 0, QTableWidgetItem(row[0]))
            self.intListTable.setItem(i, 1, QTableWidgetItem(row[1]))
            self.intListTable.setItem(i, 2, QTableWidgetItem(row[2]))

    def dispTotalItemsTable(self, kospiList, kosdaqList):
        #self.totalItemsTable.setho

        self.totalListTable.setRowCount(len(kospiList)-1+len(kosdaqList)-1)
        self.totalListTable.setColumnCount(3)
        self.totalListTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.totalListTable.verticalHeader().hide()
        self.totalListTable.showGrid()
        self.totalListTable.setColumnHidden(2, True)
        column_headers = ['코드', '종목명', '업종코드']
        self.totalListTable.setHorizontalHeaderLabels(column_headers)
        self.ItemListModel = []
       
        for i, kospiItem in enumerate(kospiList, 0):
            self.totalListTable.setItem(i, 0, QTableWidgetItem(kospiItem[0]))
            self.totalListTable.setItem(i, 1, QTableWidgetItem(kospiItem[1]))
            self.totalListTable.setItem(i, 2, QTableWidgetItem("001"))
            self.ItemListModel.append(kospiItem[1])
        for i, kosdaqItem in enumerate(kosdaqList, i):
            self.totalListTable.setItem(i, 0, QTableWidgetItem(kosdaqItem[0]))
            self.totalListTable.setItem(i, 1, QTableWidgetItem(kosdaqItem[1]))
            self.totalListTable.setItem(i, 2, QTableWidgetItem("101"))
            self.ItemListModel.append(kosdaqItem[1])
        #
        #ItemSearch LineEdit의 자동완성 기능을 부여함.
        #자동완성기능의 Model을 지정하기 위해 모든 종목 리스트를 얻은 시점인 여기서 자동완성 기능을 부여한다.
        #
        #model = QStringListModel()
        #model.setStringList(self.ItemListModel)
        completer = QCompleter(self.ItemListModel)
        #completer.setModel(model)
        #completer.setCompletionMode(1)
        #completer.setModelSorting(1)#2: The model is sorted case insensitively.
        completer.setCaseSensitivity(2)#대소문자 구분 없이 자동완성 기능 구현
        #completer.setCompletionRole(4)#0: 한글 자동완성이 한템포느림, 1: 자동완성 동작안함, 2: 한글 자동완성이 한템포느림
        self.ItemSearch.setCompleter(completer)
        #print(str(completer.completionRole()))
        self.getCode.stop()

    def procFunc(self):
        functions = {'일자별 수급': self.funcBuyer,
                     '일자별 거래량': self.funcVolumn}
        func = functions[self.comboBox.currentText()]
        func()

    def funcBuyer(self):
        x = self.intListTable.selectedIndexes()#선택된 셀의 행/열 번호가 반환된다.
        if len(x) != 0:
            strCode = str(self.intListTable.item(x[0].row(), const.codeLine).text())
            strName = str(self.intListTable.item(x[0].row(), const.nameLine).text())
            strToday = self.getTodayStr()
            strQuantity = str(1)#1이면 금액, 2면 수량
            strTrade = str(0)#0이면 순매수, 1이면 매수, 2면 매도
            strUnit = str(1)#1000이면 천주, 1이면 단주
            self.opt10059Req(strToday, strCode, strQuantity, strTrade, strUnit)
            #self.worker = Buyer(self.kiwoom)
            #self.worker.procComplete.connect(self.procComplete)   # 시그널 슬롯 등록
            #self.worker.start()  

    def opt10015Req(self, code, startDate):
        strDate = str(startDate)#YYYYMMDD 형태로 정리해야함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드",
                                code)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "시작일자",
                                strDate)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("CommRqData(QString, QString, QString, QString)", "opt10015_req", "opt10015", "0",
                                "0101")  # 키움 dynamicCall 함수를 통해 CommRqData 함수를 호출하여 opt10015 API를 구분명 opt10015_req, 화면번호 0101으로 호출함
    
    def opt10059Req(self, strToday, strCode, strQuantity, strTrade, strUnit):
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "일자",
                                strToday)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드",
                                strCode)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "금액수량구분",
                                strQuantity)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분",
                                strTrade)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "단위구분",
                                strUnit)                                                                     
        self.kiwoom.dynamicCall("CommRqData(QString, QString, QString, QString)", "opt10059_req", "opt10059", "0",
                                "0101")

    def funcVolumn(self):
        x = self.intListTable.selectedIndexes()#선택된 셀의 행/열 번호가 반환된다.
        if len(x) != 0:
            self.worker = Volumn(self.kiwoom)
            self.worker.procComplete.connect(self.procComplete)   # 시그널 슬롯 등록
            self.worker.start()  

    def procComplete(self, list1, list2):
        list1 = list2
        self.worker.terminate()

    def getTodayStr(self):
        x = dt.datetime.now()
        strYear = str(x.year)
        if int(x.month) < 10:
            strMonth = "0"+str(x.month)
        else:
            strMonth = str(x.month)
        if int(x.day) < 10:
            strDay = "0"+str(x.day)
        else:
            strDay = str(x.day)
        return strYear + strMonth + strDay



class GetItems(QThread):
    procComplete = pyqtSignal(list, list)

    def __init__(self, kiwoom):
        super().__init__()
        self.kiwoom = kiwoom
        self.working = True
    def run(self):
        #pydevd.connected = True
        #pydevd.settrace(suspend=False)
        kospi = self.GetCodeListByMarket(0)
        kosdaq = self.GetCodeListByMarket(10)
        kospi = list(kospi.split(";"))#kospi는 str타입이므로 ';'구분자를 통해 리스트 로 변경한다.
        kosdaq = list(kosdaq.split(";"))#kosdaq은 str타입이므로 ';'구분자를 통해 리스트 로 변경한다.
        kospiList = []
        kosdaqList = []
        for i, kospiCode in enumerate(kospi, 0):
            code = kospiCode
            name = self.GetNameByCode(code)
            kospiList.insert(i, [code, name])
        for i, kosdaqCode in enumerate(kosdaq, 0):
            code = kosdaqCode
            name = self.GetNameByCode(code)
            kosdaqList.insert(i, [code, name])
        self.procComplete.emit(kospiList, kosdaqList)
    def GetCodeListByMarket(self, market):
        codelist = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", market)
        return codelist
    def GetNameByCode(self, code):
        name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
        return name
    def stop(self):
        self.working = False
        self.quit()
        self.wait(100) #100ms

class Buyer(QThread):
    procComplete = pyqtSignal(list, list)

    def __init__(self, kiwoom):
        super().__init__()
        self.kiwoom = kiwoom
    
    def run(self):
        #pydevd.connected = True
        #pydevd.settrace(suspend=False)
        i=1
        a = []
        b = []
        self.procComplete.emit(a, b)
        self.opt10015Req("000020", "20211007")

    def opt10015Req(self, code, startDate):
        strDate = str(startDate)#YYYYMMDD 형태로 정리해야함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드",
                                code)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "시작일자",
                                strDate)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("CommRqData(QString, QString, QString, QString)", "opt10015_req", "opt10015", "0",
                                "0101")  # 키움 dynamicCall 함수를 통해 CommRqData 함수를 호출하여 opt10015 API를 구분명 opt10015_req, 화면번호 0101으로 호출함
    
class Volumn(QThread):
    procComplete = pyqtSignal(list, list)

    def __init__(self, kiwoom):
        super().__init__()
        self.kiwoom = kiwoom
    def run(self):
        #pydevd.connected = True
        #pydevd.settrace(suspend=False)
        i=1
        a = []
        b = []
        self.procComplete.emit(a, b)

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





