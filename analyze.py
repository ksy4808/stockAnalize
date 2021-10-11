import os
import sys
import datetime
import sqlite3
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from mpl_toolkits.mplot3d import Axes3D


from pandas import Series, DataFrame

import multiprocessing as mp
from PyQt5.QtWidgets import *
from PyQt5 import QtGui

from PyQt5 import uic
from PyQt5.QtCore import QDate
from PyQt5.QtCore import pyqtSlot, Qt

import datetime as dt
from PyQt5.QtCore import QStringListModel
from multiprocessing import Pool
from os import getpid
import ItemSort
import threading
import time


form_class = uic.loadUiType('analyze.ui')[0]

Interrestset = 0


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.fig = plt.Figure()
        self.canvas = FigureCanvas(self.fig)
        self.ConcHistoryGraph.addWidget(self.canvas)

        self.dateEdit.setDate(datetime.date.today())
        self.dateEdit.setMinimumDate(QDate(2021, 1, 1))
        self.dateEdit.setMaximumDate(QDate(2100, 12, 31))

        self.dateEdit_2.setDate(datetime.date.today())
        self.dateEdit_2.setMinimumDate(QDate(2021, 1, 1))
        self.dateEdit_2.setMaximumDate(QDate(2100, 12, 31))

        self.dateEdit_3.setDate(datetime.date.today())
        self.dateEdit_3.setMinimumDate(QDate(2021, 1, 1))
        self.dateEdit_3.setMaximumDate(QDate(2100, 12, 31))


        #콤보box에 모든 item을 나열한다.
        tables = self.getAllItemList()
        for index in range(len(tables)):
            item = self.getPureText(tables[index])
            self.item_comboBox.addItem(item)

        #콤보박스에서 종목을 변경하면 테이블을 갱신한다.
        self.item_comboBox.currentIndexChanged.connect(self.dispConcludeTable)
        #날짜 위제에서 날짜를 변경하면 테이블을 갱신한다.
        self.dateEdit.dateTimeChanged.connect(self.dispConcludeTable)

        #스위치 이벤트를 mathod에 연결한다.
        self.run_btn.clicked.connect(self.runBtn)
        self.concludeByItematDate.clicked.connect(self.ConcludeBy_ItematDate)
        self.reArrange.clicked.connect(self.re_arrange)
        self.concluionDetail.clicked.connect(self.ConcludeBy_custom)



        #체결정보 테이블 그리기 최초 프로그램실행시 설정되어있는 날짜와 종목의 체결정보를 표시한다.
        self.dispConcludeTable()

        #메뉴바에서 관심종목 항목 선택시 이벤트 연결
        self.call_Itemsort.triggered.connect(run_subProcItemSort)

        self.lineEdit_second.setValidator(QtGui.QIntValidator(1,9999,self))#0초 부터 9999초까지 설정가능 범위로 한다.

        self.plotConcHistoryGraph()
    def plotConcHistoryGraph(self):
        x = np.arange(0, 100, 1)
        y = np.sin(x)

        ax = self.fig.add_subplot(111)
        ax.plot(x, y, label="label")
        ax.set_xlabel("x_axis")
        ax.set_xlabel("y_axis")

        ax.set_title("my graph")
        ax.legend()
        self.canvas.draw()

    def ConcludeBy_custom(self):
        return
    def re_arrange(self):
        datestr = self.getDateForFileName()
        item = self.item_comboBox.currentText()
        con = sqlite3.connect("../../conclusions/" + datestr + ".db")
        cur = con.cursor()
        cur.execute("select * from %s;" % item)
        rows = cur.fetchall()
        con.close()
        startTimeWindowStr = "090000"
        endTimeWindowStr = "153000"

        dateFormatter = "%Y_%m_%d %H%M%S"
        x = self.getDateForFileName()+" "+startTimeWindowStr
        startTimeWindow = dt.datetime.strptime(x, dateFormatter)

        if self.lineEdit_second.text() == "":
            deltaSeconds = 0
        else:
            deltaSeconds = int(self.lineEdit_second.text())
        endTimeWindow = startTimeWindow + dt.timedelta(seconds=deltaSeconds)
        arrangeSet = []
        valueAccumulate = 0
        quantityAccumulate = 0
        j = 0
        for i in range(len(rows)):
            x = self.getDateForFileName() + " " + endTimeWindowStr
            if startTimeWindow >= dt.datetime.strptime(x, dateFormatter):
                break
            rowsOffset = len(rows)-1-i
            row = rows[rowsOffset]
            x = self.getDateForFileName() + " " + row[1]
            curTime = dt.datetime.strptime(x, dateFormatter)
            if curTime < startTimeWindow:#최초 시작 시간보다 앞선 경우 처리하지 않고 넘긴다.(장중 데이터만 처리하는데 개장 전 데이터는 무시하기 위함.
                continue
            elif curTime >= startTimeWindow and curTime < endTimeWindow:#현재 데이터가 time window내에 있는경우 accumulate한다.
                valueAccumulate += int(row[4])
                quantityAccumulate += int(row[3])
            else:#현재 데이터가 time window끝보다 크면 이전 accumulate했던 값을 startTimeWindow시점에 기록하고 새로 accumulate한다.
                arrangeUnit = []#여기서 생성을 해야 2차원리스트로 삽입이 됨.
                arrangeUnit.append(j)
                j += 1
                arrangeUnit.append(startTimeWindow)
                if self.lineEdit_TransVolume.text() == "":  # 기준금액이 비어있으면 모두 sum
                    arrangeUnit.append(valueAccumulate)
                    arrangeUnit.append(quantityAccumulate)
                elif abs(valueAccumulate) >= int(self.lineEdit_TransVolume.text()): #기준금액이 채워져있으면 sum한 금액이 기준금액보다 큰 경우 append
                    arrangeUnit.append(valueAccumulate)
                    arrangeUnit.append(quantityAccumulate)
                else:#기준금액보다 작으면 0을 append
                    arrangeUnit.append(0)
                    arrangeUnit.append(0)

                arrangeSet.append(arrangeUnit)
                valueAccumulate = 0
                quantityAccumulate = 0
                valueAccumulate += int(row[4])
                quantityAccumulate += int(row[3])

                startTimeWindow = endTimeWindow
                endTimeWindow = startTimeWindow + dt.timedelta(seconds=deltaSeconds)
                if curTime >= endTimeWindow:#현재 체결시간이 shift한 time Window보다 크면 time Window를 더 shift해야한다.
                    while curTime >= endTimeWindow:#현재 데이터의 시간이 time window안에 들어올때까지 time window를 shift한다.
                        arrangeUnit = []  # 여기서 생성을 해야 2차원리스트로 삽입이 됨.
                        arrangeUnit.append(j)
                        j += 1
                        arrangeUnit.append(startTimeWindow)
                        arrangeUnit.append(0)#체결금액 0으로 clear
                        arrangeUnit.append(0)#체결량 0으로 clear
                        arrangeSet.append(arrangeUnit)
                        # 현재 데이터의 시간까지 도달하기 전까지는 time window의 값을 0으로 채운다.
                        startTimeWindow = endTimeWindow
                        endTimeWindow = startTimeWindow + dt.timedelta(seconds=deltaSeconds)
        j = 0
        for i, row in enumerate(arrangeSet, 0):
            if row[3] != 0:#기준 금액 이상의 거래가 있는 경우만 표시하도록 함. 안그러면 평균가 계산할때 0 나누기때문에 에러남.
                j += 1
        self.concTableWidget_2.setColumnCount(3)
        self.concTableWidget_2.setRowCount(j)#기준금액을 넘는 경우만 리스트에 표시
        self.concTableWidget_2.setHorizontalHeaderLabels(["시간", "체결금액", "체결평균가"])
        Transe_Volume_sum = 0
        j = 0
        for i, row in enumerate(arrangeSet, 0):
            insertTime = QTableWidgetItem(str(row[1]))
            insertPrice = QTableWidgetItem(str(format(row[2],",")))
            if row[3] != 0:#기준 금액 이상의 거래가 있는 경우만 표시하도록 함. 안그러면 평균가 계산할때 0 나누기때문에 에러남.
                insertmean = QTableWidgetItem(str(format(row[2]/row[3], ",")))
                self.concTableWidget_2.setItem(j, 0, insertTime)
                self.concTableWidget_2.setItem(j, 1, insertPrice)
                self.concTableWidget_2.setItem(j, 2, insertmean)
                j += 1
            Transe_Volume_sum += row[2]

        self.lineEdit_TransVolume_sum.setText(str(format(Transe_Volume_sum,",")))

    def dispConcludeTable(self):
        datestr = self.getDateForFileName()
        item = self.item_comboBox.currentText()
        con = sqlite3.connect("../../conclusions/" + datestr + ".db")
        cur = con.cursor()
        cur.execute("select * from %s;" %item)
        rows = cur.fetchall()
        con.close()
        self.concTableWidget.setColumnCount(4)
        self.concTableWidget.setRowCount(len(rows))
        self.concTableWidget.setHorizontalHeaderLabels(["시간","체결가","체결량","체결금액"])

        for i, row in enumerate(rows, 0):
            insertTime = QTableWidgetItem(row[1])
            insertPrice = QTableWidgetItem(str(format(row[2],",")))
            insertQty = QTableWidgetItem(str(format(row[3],",")))
            insertValue = QTableWidgetItem(str(format(row[4],",")))
            self.concTableWidget.setItem(i, 0, insertTime)
            self.concTableWidget.setItem(i, 1, insertPrice)
            self.concTableWidget.setItem(i, 2, insertQty)
            self.concTableWidget.setItem(i, 3, insertValue)

    def ConcludeBy_ItematDate(self):
        datestr = self.getDateForFileName()
        item = self.item_comboBox.currentText()
        con = sqlite3.connect("../../conclusions/" + datestr + ".db")
        cur = con.cursor()
        cur.execute("select * from %s;" %item)
        rows = cur.fetchall()
        con.close()
        y_data = []
        x_data = []
        dateFormatter = "%Y_%m_%d %H%M%S%f"
        prevCnt = 0
        x_prev = 0
        for i, row in enumerate(rows, 0):
            y_data.insert(0, int(row[4]))
            x = row[1]
            if x == x_prev:
                prevCnt -= 1
            else:
                prevCnt = 1000
            x_prev = x
            x = datestr+" "+x+str(prevCnt).zfill(6)
            time = dt.datetime.strptime(x, dateFormatter)
            x_data.insert(0, time)
        plt.plot_date(x_data,y_data,linestyle="-",marker=".")
        plt.show()

    def getPureText(self, tuple):
        item = str(tuple)
        item = item.replace("('", "")
        item = item.replace("',)", "")
        return item

    def getAllItemList(self):
        datestr = self.getDateForFileName()
        con = sqlite3.connect("../../conclusions/" + datestr + ".db")
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
        rows = cur.fetchall()
        con.close()
        #cur.execute("select * from KospiItems;")
        return rows
    def getDateForFileName(self):
        date = self.dateEdit.date()
        #datestr = "2021_01_08"
        datestr = date.toString("yyyy_MM_dd")
        return datestr
    def runBtn(self):
        date = self.dateEdit.date()
        datestr = date.toString("yyyy_MM_dd")

        con = sqlite3.connect("../../conclusions/" + datestr + ".db")
        cur = con.cursor()
        cur.execute("select * from KosdaqItems;")
        con.close()



def runItemSort():
    app = QApplication(sys.argv)
    myWindow = ItemSort.MyWindow()  # MyWindow 클래스를 생성하여 myWondow 변수에 할당
    myWindow.show()  # MyWindow 클래스를 노출
    app.exec_()  # 메인 이벤트 루프에 진입 후 프로그램이 종료될 때까지 무한 루프 상태 대기
def run_subProcItemSort():

    p_itemSort = mp.Process(target=runItemSort, args=())
    p_itemSort.start()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()


