#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys,os,win32api
import win32com.client as win32
from PyQt5.QtWidgets import (QApplication,QWidget, QTableView,
                             QVBoxLayout)
from PyQt5.QtGui import QStandardItemModel,QStandardItem
from PyQt5.QtCore import Qt

class statusTableWidget(QWidget):

    def __init__(self,parent=None):
        super().__init__(parent)
        self.initUI()

    def initUI(self):
        #statusReplaceModel
        statusReplaceModel=QStandardItemModel()
        statusReplaceModel.setColumnCount(3)
        statusReplaceModel.setHorizontalHeaderLabels(["Файл","Найти","Заменить","Объект","Статус замены"])
        #/statusReplaceModel
        StatusTable=QTableView(self)
        StatusTable.setModel(statusReplaceModel)
        StatusTable.setToolTip("Отчетная таблица")

        #mainLayout
        mainLayout = QVBoxLayout()
        mainLayout.setContentsMargins(0,0,0,0)
        mainLayout.addWidget(StatusTable)
        self.setLayout(mainLayout)
        #/mainLayout
        #self

        self.statusTable = StatusTable
        #/self
        StatusTable.doubleClicked.connect(self.openFile)

    def openFile(self,QModindex):
        if QModindex.column()==0:
            model=self.statusTable.model()
            file=model.data(QModindex,Qt.UserRole)
            os.startfile(file)

    def clear(self):
        self.statusTable.model().removeRows(0,self.statusTable.model().rowCount())

    def appendRow(self,file="",find="",replace="",object="",status=False):
        status="Выполнено" if status else "Не найдено"
        cfile=os.path.split(file)[1]
        items=[QStandardItem(x) for x in [cfile,find,replace,object,status]]
        items[0].setData(file,Qt.UserRole)
        self.statusTable.model().appendRow(items)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = statusTableWidget()
    ex.show()
    sys.exit(app.exec_())
