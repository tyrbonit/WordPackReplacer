#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys,os,win32api
import win32com.client as win32
from PyQt5.QtWidgets import (QApplication, QHBoxLayout, QWidget,QGroupBox,
                             QVBoxLayout,QPushButton,QCheckBox,QFileDialog,QListView,
                             QPlainTextEdit,QLineEdit)
from PyQt5.QtGui import QIcon,QStandardItemModel,QStandardItem
from PyQt5.QtCore import QSize
class FilesTableWidget(QWidget):

    def __init__(self,parent=None):
        super().__init__(parent)
        self.FilesList=[]
        self.initUI()

    def initUI(self):
        #GrBox1
        GrBox=QGroupBox()
        vbox = QHBoxLayout(self)
        vbox.setContentsMargins(0,0,0,0)
        #GrBox.setFixedHeight(60)
        pathButton = QPushButton(QIcon('icons\\pack.png'),"Папка")
        pathButton.setIconSize(QSize(25,25))
        pathButton.setVisible(True)
        pathLable=QLineEdit()
        #pathLable.setReadOnly(True)
        subPathCheck=QCheckBox()
        subPathCheck.setText("Подпапки")
        subPathCheck.setCheckState(0)
        vbox.addWidget(pathLable)
        vbox.addWidget(subPathCheck)
        vbox.addWidget(pathButton)
        GrBox.setLayout(vbox)
        #/GrBox1

        #FilesTable
        FilesTable=QListView(self)
        FilesTable.setToolTip("Список файлов, выберите нужные файлы для обработки,\nдля просмотра файла дважды щелкните по нему")
        FilesTableModel = QStandardItemModel()
        FilesTable.setModel(FilesTableModel)
        #/FilesTable
        #mainLayout
        mainLayout = QVBoxLayout()
        mainLayout.setContentsMargins(0,0,0,0)

        mainLayout.setMenuBar(GrBox)
        mainLayout.addWidget(FilesTable)
        #/mainLayout

        #self
        self.setLayout(mainLayout)
        self.path=pathLable.text()
        self.pathLable=pathLable
        self.subPathCheck=subPathCheck
        self.FilesTableModel = FilesTableModel
        #/self
        #connections
        pathLable.textChanged.connect(self.setPath)
        pathButton.clicked.connect(self.selectPath)
        subPathCheck.clicked.connect(self.setPath)
        FilesTableModel.itemChanged.connect(self.ChangeFilesList)
        FilesTable.doubleClicked.connect(self.openFile)

    def openFile(self,QModindex):
        model=self.FilesTableModel
        file=self.path+os.sep+model.item(QModindex.row()).text()
        os.startfile(file)

    def selectPath(self):
        Path=os.path.normpath(QFileDialog.getExistingDirectory(directory="e:\\temp"))
        self.pathLable.setText(Path)
        self.setPath()

    def setPath(self):
        self.path=os.path.normpath(self.pathLable.text())
        if os.path.exists(self.path): self.scanPath()
        #print(self.path)

    def scanPath(self):
        path=self.path
        sbpCheck=self.subPathCheck.checkState()==2
        listdir=[]
        if sbpCheck:
            [[listdir.append(p.replace(path+os.sep,"")+os.sep+x if p.replace(path,"")!="" else x) for x in f] for (p,d,f) in os.walk(path)]
        else:
            listdir=os.listdir(path)
        #print(listdir)
        iconDict={"doc":QIcon("icons\\doc.png"),"docx":QIcon("icons\\docx.png")}
        model=self.FilesTableModel
        model.clear()
        self.FilesList.clear()
        for x in listdir:
            filepath=self.path+os.sep+x
            if not os.path.exists(filepath) or os.path.isdir(filepath): continue

            attrs=win32api.GetFileAttributes(filepath)
            #print(p,x, attrs,    win32con.FILE_ATTRIBUTE_HIDDEN) Исключаем все скрытые файлы
            if attrs in [2,3,34,35]: continue#win32con.FILE_ATTRIBUTE_HIDDEN=2,arch=32,readonly=1
            keyIcon=x.split(".")[-1]
            if keyIcon in iconDict.keys():
                item=QStandardItem(iconDict[keyIcon],x)
                item.setCheckable(True)
                item.setCheckState(2)
                item.setEditable(False)
                model.appendRow(item)
        self.ChangeFilesList()

    def ChangeFilesList(self,QItemModel=QStandardItem()):
        model=self.FilesTableModel
        self.FilesList.clear()
        [[self.FilesList.append(self.path+os.sep+model.item(i).text())
          if model.item(i).checkState()==2 else None]
         for i in range(model.rowCount())]
        #print(self.FilesList)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FilesTableWidget()
    ex.show()
    sys.exit(app.exec_())
