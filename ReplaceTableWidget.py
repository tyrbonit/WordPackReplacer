#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys,os,win32api
import win32com.client as win32
from PyQt5.QtWidgets import (QMainWindow, QAction, QApplication, QHBoxLayout, QSplitter, QWidget, QTableView,QGroupBox,
                             QVBoxLayout,QDialogButtonBox,QPushButton,QCheckBox,QFileDialog,QListView,QHeaderView,
                             QPlainTextEdit,QMessageBox,QProgressDialog,QItemDelegate,QToolBar,QStackedLayout,QTabWidget,
                             QToolBox,QStyleOptionViewItem,QStyle,QStyledItemDelegate,QStyleOptionButton,QFormLayout,QLabel,QLineEdit,QGridLayout,QLayout,QSizePolicy)
from PyQt5.QtGui import QIcon,QStandardItemModel,QStandardItem,QPainter
from PyQt5.QtCore import Qt,QSize,QModelIndex,QEvent,QPoint,QRect,QObject,pyqtSignal
from MSWordConstants import wdFindContinue,wdReplaceAll,wdYellow,msoTextEffect,WdStoryType,WdInlineShapeType,MsoShapeType,wdColorBlack
from Delegates import PlainTextEditDelegate,ButtonDelegate
from OptionsWidget import Options
import pymorphy2

class ReplaceTableWidget(QWidget):

    def __init__(self,parent=None):
        super().__init__(parent)
        self.ReplaceList=[]
        self.ReplaceItemOptions=[]
        self.initUI()

    def initUI(self):
        #ReplaceModel
        ReplaceModel=QStandardItemModel()
        ReplaceModel.setColumnCount(3)
        ReplaceModel.setHorizontalHeaderLabels(["Найти","Заменить на","Настройки"])
        #/ReplaceModel

        #replaceTable
        replaceTable = QTableView(self)
        #replaceTable.setToolTip("Таблица поиска и замены,\nдобавьте строку и заполните текст для поиска и замены,\nвы также можете выбрать индивидуальные опции поиска для данной строки")
        #replaceTable.horizontalHeader().setToolTip("Найти - Текст для поиска\nЗаменить - Текст замены\nЗн. - Подстановочные знаки\nЦв. - Выделение цветом\nЖ - Выделение жирным шрифтом\nН - Изменять надписи\nWA - Изменять объкты WordArt")
        replaceTable.setModel(ReplaceModel)
        replaceTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        replaceTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        replaceTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        """replaceTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        replaceTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        replaceTable.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)"""
        replaceTable.horizontalHeader().resizeSection(0,500)
        replaceTable.setItemDelegateForColumn(0,PlainTextEditDelegate(replaceTable))
        replaceTable.setItemDelegateForColumn(1,PlainTextEditDelegate(replaceTable))

        #btnDeleg=ButtonDelegate(replaceTable)
        #btnDeleg.clicked.connect(self.setItemOpt)
        #replaceTable.clicked.connect(self.setItemOpt)

        #replaceTable.setItemDelegateForColumn(2,btnDeleg)
        #replaceTable.setItemDelegateForColumn(6,CheckBoxDelegate(replaceTable))
        #/replaceTable

        #buttonBox
        sklonPadejBtn=QPushButton("Просклонять")
        addRowBtn=QPushButton("Добавить строку")
        delRowBtn=QPushButton("Удалить строку")
        clearTable=QPushButton("Очистить")
        buttonBox = QDialogButtonBox(Qt.Horizontal)
        buttonBox.addButton(sklonPadejBtn, QDialogButtonBox.ActionRole)
        buttonBox.addButton(addRowBtn, QDialogButtonBox.ActionRole)
        buttonBox.addButton(delRowBtn, QDialogButtonBox.ActionRole)
        buttonBox.addButton(clearTable, QDialogButtonBox.ActionRole)
        #/buttonBox

        #mainLayout
        mainLayout = QVBoxLayout()
        mainLayout.setContentsMargins(0,0,0,0)

        self.optionWdgt=Options()
        mainLayout.setMenuBar(self.optionWdgt)

        mainLayout.addWidget(replaceTable)
        mainLayout.addWidget(buttonBox)
        #/mainLayout

        #self
        self.setLayout(mainLayout)
        self.ReplaceModel=ReplaceModel
        self.replaceTable=replaceTable
        #/self

        #connections
        sklonPadejBtn.clicked.connect(self.addSklonenie)
        addRowBtn.clicked.connect(self.addRow)
        delRowBtn.clicked.connect(self.delRow)
        clearTable.clicked.connect(self.clearReplaceTable)
        ReplaceModel.itemChanged.connect(self.ChangeReplaceList)

    def Sklonenie(self,Text):
        try:
            morph = pymorphy2.MorphAnalyzer()
            Text=Text.split(" ")
            varTxt=[[morph.parse(word)[0].inflect({"sing",sklon}).word for sklon in ["nomn","gent","datv","accs","ablt","loct"]] for word in Text]
            varTxt=zip(*varTxt)
            return [" ".join(x) for x in varTxt]
        except:
            return 5*Text

    def addSklonenie(self):
        row=self.replaceTable.currentIndex().row()
        find=self.ReplaceModel.item(row,0).data(Qt.DisplayRole)
        repl=self.ReplaceModel.item(row,1).data(Qt.DisplayRole)
        findList=self.Sklonenie(find)
        replList=self.Sklonenie(repl)
        for x in zip(findList,replList):
            if x[0]!=find or x[1]!=repl: self.addRow(FindRepl=x)

    def addRow(self,arg=None,FindRepl=["",""]):
        #print(arg,FindRepl)
        row=self.replaceTable.currentIndex().row()
        ToolTipText=["Текст для поиска","Текст замены",
                     "Использовать индивидуальные настройки для данной строки"]
        items=[]
        for i, TText in enumerate(ToolTipText):
            Item=QStandardItem()
            Item.setData(TText,Qt.ToolTipRole)
            if i>1:
                Item.setData(False,Qt.CheckStateRole)
                Item.setData("Задать",Qt.DisplayRole)
                Item.setData(self.optionWdgt,Qt.UserRole)
                Item.setEditable(False)
                Item.setCheckable(True)
            else:
                Item.setData(FindRepl[i],Qt.DisplayRole)
            items.append(Item)
        self.ReplaceModel.insertRow(row+1,items)
        self.replaceTable.selectRow(row+1)
        self.replaceTable.resizeRowsToContents()

    def delRow(self):
        row=self.replaceTable.currentIndex().row()
        self.ReplaceModel.removeRow(row)
        self.replaceTable.selectRow(row-1)
        self.ChangeReplaceList()

    def clearReplaceTable(self):
        msgBox = QMessageBox()
        msgBox.setInformativeText("Очистить таблицу?")
        msgBox.addButton(QMessageBox.Yes)
        msgBox.addButton(QMessageBox.No)
        msgBox.setDefaultButton(QMessageBox.No)
        ret = msgBox.exec_()
        if ret == QMessageBox.Yes:
            self.ReplaceModel.removeRows(0,self.ReplaceModel.rowCount())
            self.ReplaceList.clear()

    def ChangeReplaceList(self,QMitem=QStandardItem()):

        self.ReplaceModel.blockSignals(True)
        if QMitem.column()==2 and QMitem.data(Qt.CheckStateRole)==2:
            #item=self.ReplaceModel.item(index.row(),index.column())
            itemOptWdg=Options(self)
            itopt=QMitem.data(Qt.UserRole)
            itemOptWdg.setOptions(self.optionWdgt.OptionsDict)
            itemOptWdg.setWindowModality(Qt.WindowModal)
            #itemOptWdg.setGeometry(50, 100, 500, 400)
            itemOptWdg.setWindowTitle("Опции")
            itemOptWdg.setWindowFlags(Qt.Dialog|Qt.WindowMinMaxButtonsHint|Qt.WindowCloseButtonHint)
            itemOptWdg.show()
            QMitem.setData(itemOptWdg,Qt.UserRole)
        else:
            QMitem.setData(self.optionWdgt,Qt.UserRole)
        self.ReplaceModel.blockSignals(False)

        """Ограничение длины строки в Ворде 255 символов, при этом знаки переноса(конца абзаца)
        не считаются(в отличии от питона), поэтому строку обрезаем с учетом знаков переноса"""
        text=QMitem.text()
        n=text.count("\n")-text[255:].count("\n")
        if len(text)>255:QMitem.setText(text[:255-n])
        #print(len(QMitem.text()))

        self.replaceTable.resizeRowsToContents()

        model=self.ReplaceModel
        model.sort(0,Qt.AscendingOrder)
        self.ReplaceList.clear()
        for i in range(model.rowCount()):
            self.ReplaceList.append([model.item(i,0).text(),model.item(i,1).text(), model.item(i,2).data(Qt.UserRole).OptionsDict])
        #print(self.ReplaceList)
    def closeEvent(self, Event):
        print(self.ReplaceList)
        Event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ReplaceTableWidget()
    ex.resize(1024,768)
    ex.show()
    sys.exit(app.exec_())