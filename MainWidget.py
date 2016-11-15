#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import (QApplication, QHBoxLayout, QSplitter, QWidget, QTableView,
                             QVBoxLayout,QListView,QPushButton)
from PyQt5.QtCore import Qt,QObject,QEvent

class MainWidget(QWidget):

    def __init__(self,parent=None):
        super().__init__(parent)
        self.LeftLay=QVBoxLayout()
        self.LeftLay.setContentsMargins(0,0,0,0)

        self.RightLay=QVBoxLayout()
        self.RightLay.setContentsMargins(0,0,0,0)
        #btn=QPushButton("sdsdsd")
        #self.RightLay.setMenuBar(btn)

        self.BottomLay=QVBoxLayout()
        self.BottomLay.setContentsMargins(0,0,0,0)
        self.initUI()
        #btn.clicked.connect(self.btncl)
    """def btncl(self):
        #QSplitter.sizes()
        print(self.sender().parent().parent().sizes())
        self.sender().parent().parent().setSizes([1,0])"""

    def initUI(self):
        mainLayout = QHBoxLayout(self)
        LeftWidget=QWidget(self)
        RightWidget=QWidget(self)
        BottomWidget=QWidget(self)

        LeftWidget.setLayout(self.LeftLay)
        RightWidget.setLayout(self.RightLay)
        BottomWidget.setLayout(self.BottomLay)

        self.HSplitter = QSplitter(Qt.Horizontal)
        self.HSplitter.addWidget(LeftWidget)
        self.HSplitter.addWidget(RightWidget)

        self.VSplitter = QSplitter(Qt.Vertical)
        self.VSplitter.addWidget(self.HSplitter)
        self.VSplitter.addWidget(BottomWidget)
        #self.HSplitter.setSizes([300,300])
        #self.VSplitter.setSizes([300,500])
        mainLayout.addWidget(self.VSplitter)
        self.setLayout(mainLayout)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWidget()
    ex.LeftLay.addWidget(QListView(ex))
    ex.RightLay.addWidget(QTableView(ex))
    ex.BottomLay.addWidget(QTableView(ex))
    ex.show()
    sys.exit(app.exec_())
