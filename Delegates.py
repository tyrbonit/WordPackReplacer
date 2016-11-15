#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys,os
from PyQt5.QtWidgets import (QMainWindow, QAction, QApplication, QHBoxLayout, QSplitter, QWidget, QTableView,QGroupBox,
                             QVBoxLayout,QDialogButtonBox,QPushButton,QCheckBox,QFileDialog,QListView,QHeaderView,
                             QPlainTextEdit,QMessageBox,QProgressDialog,QItemDelegate,QToolBar,QStackedLayout,QTabWidget,
                             QToolBox,QStyleOptionViewItem,QStyle,QStyledItemDelegate,QStyleOptionButton)
from PyQt5.QtGui import QIcon,QStandardItemModel,QStandardItem,QPainter
from PyQt5.QtCore import Qt,QSize,QModelIndex,QEvent,QPoint,QRect,QObject,pyqtSignal

class PlainTextEditDelegate(QItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QPlainTextEdit(parent)
        #editor.textChanged.connect(self.changeWeight)
        editor.setMinimumHeight(50)
        return editor
    #def changeWeight(self):
        #print(self.sender)

#class Communicate(QObject):
    #clicked=pyqtSignal(QModelIndex,name="clicked")

class ButtonDelegate(QItemDelegate):
    """
    Делегат, который представляет полностью функционирующую кнопку QPushButton
    в каждой клетке колонки, в которой она применяется
    """
    clicked=pyqtSignal(QModelIndex)
    def __init__(self, parent):
        # parent является обязательным аргументом для делегата
        # так как мы ссылаемся на него в методе paint (см ниже)
        #self.__c=Communicate()
        #self.clicked=self.__c.clicked
        QItemDelegate.__init__(self, parent)


    def paint(self, painter, option, index):
        # This method will be called every time a particular cell is
        # in view and that view is changed in some way. We ask the
        # delegates parent (in this case a table view) if the index
        # in question (the table cell) already has a widget associated
        # with it. If not, create one with the text for this index and
        # connect its clicked signal to a slot in the parent view so
        # we are notified when its used and can do something.
        if not self.parent().indexWidget(index):
            pb=QPushButton(
                    index.data(Qt.DisplayRole),
                    self.parent(),
                    clicked=self.__clicked
                )
            pb.index=index
            self.parent().setIndexWidget(
                index,pb)
            #QPushButton(
                    #index.data(),
                    #self.parent(),
                    #clicked=self.parent().cellButtonClicked
                #)
            #)
    def __clicked(self):
        self.clicked.emit(self.sender().index)

class TableWidget(QWidget):

    def __init__(self,parent=None):

        super().__init__(parent)
        self.initUI()

    def initUI(self):

        Model=QStandardItemModel()
        Model.setColumnCount(3)
        Model.setHorizontalHeaderLabels(["clmn1","clmn2","clmn3"])
        Model.appendRow([QStandardItem(11),QStandardItem(12),QStandardItem(13)])
        Model.appendRow([QStandardItem(21),QStandardItem(22),QStandardItem(23)])

        Table=QTableView(self)
        Table.setModel(Model)
        ButDeleg=ButtonDelegate(Table)
        Table.setItemDelegateForColumn(1,ButDeleg)
        TXTDelegate=PlainTextEditDelegate()
        Table.setItemDelegateForColumn(0,TXTDelegate)
        ButDeleg.clicked.connect(self.Prnt)

        mainLayout = QVBoxLayout()
        mainLayout.setContentsMargins(0,0,0,0)
        mainLayout.addWidget(Table)
        self.setLayout(mainLayout)

    def Prnt(self,index):
        print(index.row())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TableWidget()
    ex.show()
    sys.exit(app.exec_())