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
from PyQt5.QtCore import QStringListModel

form_class = uic.loadUiType("ItemSort.ui")[0]  # ui 파일을 로드하여 form_class 생성
Interrestset = 0

class MyWindow(QMainWindow, form_class):  # MyWindow 클래스 QMainWindow, form_class 클래스를 상속 받아 생성됨
    def __init__(self):  # MyWindow 클래스의 초기화 함수(생성자)
        super().__init__()  # 부모클래스 QMainWindow 클래스의 초기화 함수(생성자)를 호출
        self.setupUi(self)  # ui 파일 화면 출력
        self.showItems()
        lb = QLabel()#자동완성기능을 사용하기 위한 객체 생성 아마도 자동완성기능은 QLabel객체의 기능을 활용하나 봄
        self.lineEditForAdd.textChanged.connect(lb.setText)#자동완성기능을 호출
        self.lineEditForSub.textChanged.connect(lb.setText)#자동완성기능을 호출
        self.lineEditForAdd.returnPressed.connect(self.findListFromLineAdd)
        self.lineEditForAdd.returnPressed.connect(self.findListFromLineSub)


        self.add.clicked.connect(self.addItem)  # 관심종목으로 추가 method 호출
        self.sub.clicked.connect(self.subItem)  # 관심종목에서 제거 method 호출
        self.add_2.clicked.connect(self.addItemFromLineEditer)
        self.sub_2.clicked.connect(self.subItemFromLineEditer)


    def findListFromLineAdd(self):
        return
    def findListFromLineSub(self):
        return



    def addItemFromLineEditer(self):
        self.addsub(1, "FromLine")
        self.lineEditForAdd.setText("")
    def subItemFromLineEditer(self):
        self.addsub(0, "FromLine")
        self.lineEditForSub.setText("")
    def addItem(self):
        self.addsub(1, "FromList")

    def subItem(self):
        self.addsub(0, "FromList")
    def addsub(self, addsub, fromWhere):

        if fromWhere == "FromList":#리스트에서 추가/삭제진행하는 경우
            if addsub == 1:#1이면 add
                rowsellected = self.Itemlist.currentItem().text()
            elif addsub == 0:#0이면 sub
                rowsellected = self.Itemlist_En.currentItem().text()
            code = rowsellected.split(':',maxsplit=1)[0]
            name = rowsellected.split(':',maxsplit=1)[1]
        else:#line에서 추가/삭제하는 경우
            if addsub == 1:#1이면 add
                name = self.lineEditForAdd.text()
                code = self.findCodeByName(name)
            elif addsub == 0:#0이면 sub
                name = self.lineEditForSub.text()
                code = self.findCodeByName(name)

        self.addsubDB(addsub, code, name)
    def findCodeByName(self, name):
        wasKosdaq = False
        con = sqlite3.connect("Items.db")
        cur = con.cursor()
        cur.execute("select * from KosdaqItems")
        rows = cur.fetchall()
        con.close()
        for row in rows:
            if row[1] == name:
                code = row[0]
                wasKosdaq = True
                break
        if wasKosdaq == False:
            con = sqlite3.connect("Items.db")
            cur = con.cursor()
            cur.execute("select * from KospiItems")
            rows = cur.fetchall()
            con.close()
            for row in rows:
                if row[1] == name:
                    code = row[0]
                    break
        return code

    def addsubDB(self, addsub, code, name):
        wasKosdaq = False  ##추가/제거한 아니템이 코스닥이 아니였다(즉 코스피항목을 추가로 찾아보아야한다는 것을 의미)
        con = sqlite3.connect("Items.db")
        cur = con.cursor()
        cur.execute("select * from KosdaqItems")
        rows = cur.fetchall()
        con.close()
        for row in rows:
            if row[0] == code:
                setvalue = row[2]
                if addsub == 1:  # 1이면 add
                    if row[2] == None:
                        setvalue = 0
                    con = sqlite3.connect("Items.db")
                    cur = con.cursor()
                    cur.execute("UPDATE KosdaqItems SET Interrest = %d where name = '%s'" % ((setvalue | 1<<Interrestset), str(name)))
                    con.commit()
                    con.close()
                    print(str(code))
                elif addsub == 0:  # 0이면 sub
                    con = sqlite3.connect("Items.db")
                    cur = con.cursor()
                    cur.execute("UPDATE KosdaqItems SET Interrest = %d where name = '%s'" % ((setvalue & ~(1<<Interrestset)), str(name)))
                    con.commit()
                    con.close()
                wasKosdaq = True
                break

        if wasKosdaq == False:
            con = sqlite3.connect("Items.db")
            cur = con.cursor()
            cur.execute("select * from KospiItems")
            rows = cur.fetchall()
            for row in rows:
                if row[0] == code:
                    setvalue = row[2]
                    if addsub == 1:  # 1이면 add
                        if row[2] == None:
                            setvalue = 0
                        cur.execute(
                            "UPDATE KospiItems SET Interrest = %d where name = '%s'" % ((setvalue | 1 << Interrestset), name))
                    elif addsub == 0:  # 0이면 sub
                        cur.execute("UPDATE KospiItems SET Interrest = %d where name = '%s'" % ((setvalue & ~(1 << Interrestset)), name))
                    break
            con.commit()
            con.close()

        self.showItems()
    def showItems(self):
        con = sqlite3.connect("Items.db")
        cur = con.cursor()
        cur.execute("select * from KosdaqItems")
        rows = cur.fetchall()
        #cur.execute("ADD COLUMN IF NOT EXISTS KosdaqItems ADD COLUMN Interrest integer")
        #cur.execute("ALTER TABLE KosdaqItems ADD COLUMN Interrest integer")
        #cur.execute("ADD COLUMN IF NOT EXISTS EXISTS KospiItems ADD COLUMN Interrest integer")
        #con.commit()
        con.close()
        con = sqlite3.connect("Items.db")
        cur = con.cursor()
        cur.execute("select * from KospiItems")
        rows += cur.fetchall()
        self.Itemlist.clear()
        self.Itemlist_En.clear()

        forAddList = []
        forSubList = []
        for row in rows:
            #self.pteLog.Itemlist("{}  {}  {}".format(row.num, row.name, row.Interrest))
            if row[0] != "":
                if self.retInterrestattr(row) == "NoInterrest":
                    self.Itemlist.addItem("{}:{}".format(row[0], row[1]))
                    forAddList.append(str(row[1]))
                elif self.retInterrestattr(row) == "Interrest":
                    self.Itemlist_En.addItem("{}:{}".format(row[0], row[1]))
                    forSubList.append(str(row[1]))

        forAddModel = QStringListModel()##관심종목으로 추가하기 위한 line editer의 completer model
        forAddModel.setStringList(forAddList)

        forSubModel = QStringListModel()##관심종목에서 제거하기 위한 line editer의 completer model
        forSubModel.setStringList(forSubList)

        forAddCompleter = QCompleter()  ##관심종목으로 추가하기 위한 line editer의 completer
        forAddCompleter.setModel(forAddModel)
        forSubCompleter = QCompleter()##관심에서 제거하기 위한 line editer의 completer
        forSubCompleter.setModel(forSubModel)


        self.lineEditForAdd.setCompleter(forAddCompleter)
        self.lineEditForSub.setCompleter(forSubCompleter)

    def retInterrestattr(self, row):
        if row[2] == None or row[2] == 0:
            return "NoInterrest"
        elif (row[2] and 0x1) == 1:
            return "Interrest"
        else:
            return "NoInterrest"
    # py 파일 실행시 제일 먼저 동작
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()  # MyWindow 클래스를 생성하여 myWondow 변수에 할당
    myWindow.show()  # MyWindow 클래스를 노출
    app.exec_()  # 메인 이벤트 루프에 진입 후 프로그램이 종료될 때까지 무한 루프 상태 대기