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

class FlowLayout(QLayout):
    def __init__(self, parent=None, margin=0, spacing=-1):
        super(FlowLayout, self).__init__(parent)

        if parent is not None:
            self.setContentsMargins(margin, margin, margin, margin)

        self.setSpacing(spacing)

        self.itemList = []
        self.MinimumItemWidth=0

    def __del__(self):
        item = self.takeAt(0)
        while item:
            item = self.takeAt(0)

    def addItem(self, item):
        self.itemList.append(item)

    def count(self):
        return len(self.itemList)

    def itemAt(self, index):
        if index >= 0 and index < len(self.itemList):
            return self.itemList[index]

        return None

    def takeAt(self, index):
        if index >= 0 and index < len(self.itemList):
            return self.itemList.pop(index)

        return None

    def expandingDirections(self):
        return Qt.Orientations(Qt.Orientation(0))

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        height = self.doLayout(QRect(0, 0, width, 0), True)
        return height

    def setGeometry(self, rect):
        super(FlowLayout, self).setGeometry(rect)
        self.doLayout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()

        for item in self.itemList:
            size = size.expandedTo(item.minimumSize())

        margin, _, _, _ = self.getContentsMargins()

        size += QSize(2 * margin, 2 * margin)
        return size

    def doLayout(self, rect, testOnly):
        x = rect.x()
        y = rect.y()
        lineHeight = 0

        for item in self.itemList:
            wid = item.widget()
            spaceX = self.spacing() + wid.style().layoutSpacing(QSizePolicy.PushButton, QSizePolicy.PushButton, Qt.Horizontal)
            spaceY = self.spacing() + wid.style().layoutSpacing(QSizePolicy.PushButton, QSizePolicy.PushButton, Qt.Vertical)
            nextX = x + item.sizeHint().width() + spaceX
            if nextX - spaceX > rect.right() and lineHeight > 0:
                x = rect.x()
                y = y + lineHeight + spaceY
                nextX = x + item.sizeHint().width() + spaceX
                lineHeight = 0

            if not testOnly:
                item.setGeometry(QRect(QPoint(x, y), item.sizeHint()))

            x = nextX
            lineHeight = max(lineHeight, item.sizeHint().height())

        return y + lineHeight - rect.y()

class SeachReplaceOptions(QWidget):

    def __init__(self,parent=None):
        super(SeachReplaceOptions, self).__init__(parent)
        self.selectedStory={}.fromkeys(["WdStoryType","WdInlineShapeType","MsoShapeType"])
        self.StoryTypeList=[]
        self.InlineShapeList=[]
        self.ShapeTypeList=[]

        self.initUI()

    def initUI(self):
        StoryTypeGroupBox=QGroupBox()
        StoryTypeLayer=FlowLayout()
        for key in WdStoryType.keys():
            widget=QCheckBox(WdStoryType[key][1] if WdStoryType[key][1]!="" else WdStoryType[key][0])
            widget.setObjectName("WdStoryType."+str(key))
            widget.setCheckState(WdStoryType[key][2])
            widget.setEnabled(widget.checkState()!=1)
            widget.setMinimumWidth(230)
            StoryTypeLayer.addWidget(widget)
            self.StoryTypeList.append(widget)
        StoryTypeGroupBox.setLayout(StoryTypeLayer)

        InlineShapeTypeGroupBox=QGroupBox()
        InlineShapeTypeLayer=FlowLayout()
        for key in WdInlineShapeType.keys():
            widget=QCheckBox(WdInlineShapeType[key][1] if WdInlineShapeType[key][1]!="" else WdInlineShapeType[key][0])
            widget.setObjectName("WdInlineShapeType."+str(key))
            widget.setCheckState(WdInlineShapeType[key][2])
            widget.setEnabled(widget.checkState()!=1)
            widget.setMinimumWidth(230)
            InlineShapeTypeLayer.addWidget(widget)
            self.InlineShapeList.append(widget)
        InlineShapeTypeGroupBox.setLayout(InlineShapeTypeLayer)

        ShapeTypeGroupBox=QGroupBox()
        ShapeTypeLayer=FlowLayout()
        for key in MsoShapeType.keys():
            widget=QCheckBox(MsoShapeType[key][1] if MsoShapeType[key][1]!="" else MsoShapeType[key][0])
            widget.setObjectName("MsoShapeType."+str(key))
            widget.setCheckState(MsoShapeType[key][2])
            widget.setEnabled(widget.checkState()!=1)
            widget.setMinimumWidth(230)
            ShapeTypeLayer.addWidget(widget)
            self.ShapeTypeList.append(widget)
        ShapeTypeGroupBox.setLayout(ShapeTypeLayer)

        """TabWidget=QTabWidget()
        TabWidget.addTab(StoryTypeGroupBox,"StoryType")
        TabWidget.addTab(InlineShapeTypeGroupBox,"InlineShapeType")
        TabWidget.addTab(ShapeTypeGroupBox,"ShapeType")
        TabWidget.setCurrentIndex(0)"""
        TabWidget=QToolBox()
        TabWidget.addItem(StoryTypeGroupBox,"StoryType")
        TabWidget.addItem(InlineShapeTypeGroupBox,"InlineShapeType")
        TabWidget.addItem(ShapeTypeGroupBox,"ShapeType")
        TabWidget.setCurrentIndex(0)

        hbox=QVBoxLayout()
        hbox.addWidget(TabWidget)

        btnGrp=QDialogButtonBox()
        btnClose=QPushButton("Закрыть")
        #btnApply=QPushButton("Применить")
        btnGrp.addButton(btnClose,QDialogButtonBox.ActionRole)
        #btnGrp.addButton(btnApply,QDialogButtonBox.ActionRole)
        hbox.addWidget(btnGrp)

        btnClose.clicked.connect(self.close)
        #btnApply.clicked.connect(self.apply)
        self.apply()
        self.setLayout(hbox)
        self.setGeometry(500, 100, 500, 400)
        self.setWindowTitle("Опции")
        self.setWindowFlags(Qt.Dialog|Qt.WindowMinMaxButtonsHint|Qt.WindowCloseButtonHint)
        #self.show()

    def apply(self):
        self.selectedStory["WdStoryType"]=[int(wdg.objectName().split(".")[-1]) for wdg in self.StoryTypeList if wdg.checkState()==2]
        self.selectedStory["WdInlineShapeType"]=[int(wdg.objectName().split(".")[-1])for wdg in self.InlineShapeList if wdg.checkState()==2]
        self.selectedStory["MsoShapeType"]=[int(wdg.objectName().split(".")[-1])for wdg in self.ShapeTypeList if wdg.checkState()==2]
        #print(self.selectedStory)

    def setOptions(self,selectedStory={}):
        if selectedStory=={}:return
        self.selectedStory.update(selectedStory)
        [wdg.setCheckState(0) for wdg in self.StoryTypeList if wdg.checkState()==2]
        [wdg.setCheckState(2) for wdg in self.StoryTypeList if int(wdg.objectName().split(".")[-1]) in selectedStory["WdStoryType"]]
        [wdg.setCheckState(2) for wdg in self.StoryTypeList if int(wdg.objectName().split(".")[-1]) in selectedStory["WdInlineShapeType"]]
        [wdg.setCheckState(2) for wdg in self.StoryTypeList if int(wdg.objectName().split(".")[-1]) in selectedStory["MsoShapeType"]]

    def closeEvent(self, Event):
        self.apply()
        Event.accept()

class searchOptions(QWidget):
    closed=pyqtSignal(list)
    def __init__(self,parent=None):
        super(searchOptions, self).__init__(parent)
        self.WdStoryType=[]
        self.StoryTypeWdgList=[]
        self.initUI()

    def initUI(self):
        StoryTypeLayer=FlowLayout()
        for key in WdStoryType.keys():
            widget=QCheckBox(WdStoryType[key][1] if WdStoryType[key][1]!="" else WdStoryType[key][0])
            widget.setObjectName("WdStoryType."+str(key))
            widget.setCheckState(WdStoryType[key][2])
            widget.setEnabled(widget.checkState()!=1)
            widget.setMinimumWidth(230)
            StoryTypeLayer.addWidget(widget)
            self.StoryTypeWdgList.append(widget)
        self.apply()
        hbox=QVBoxLayout()
        hbox.addLayout(StoryTypeLayer)

        btnGrp=QDialogButtonBox()
        btnClose=QPushButton("Закрыть")
        btnGrp.addButton(btnClose,QDialogButtonBox.ActionRole)
        hbox.addWidget(btnGrp)
        btnClose.clicked.connect(self.close)
        self.setLayout(hbox)
        self.setGeometry(500, 100, 500, 400)
        self.setWindowTitle("Опции")
        self.setWindowFlags(Qt.Dialog|Qt.WindowMinMaxButtonsHint|Qt.WindowCloseButtonHint)
        #self.show()

    def apply(self):
        self.WdStoryType=[int(wdg.objectName().split(".")[-1]) for wdg in self.StoryTypeWdgList if wdg.checkState()==2]

    def setOptions(self,selectedStory=[]):
        #if selectedStory==[]:return
        self.WdStoryType=selectedStory

        [wdg.setCheckState(0) for wdg in self.StoryTypeWdgList if wdg.checkState()==2]
        [wdg.setCheckState(2) for wdg in self.StoryTypeWdgList if int(wdg.objectName().split(".")[-1]) in selectedStory]

    def closeEvent(self, Event):
        self.apply()
        self.closed.emit(self.WdStoryType)
        Event.accept()

class Options(QWidget):
    def __init__(self,parent=None):
        super(Options,self).__init__(parent)
        self.OptionsDict={"MatchWildcards":False,"Highlight":False,"Bold":False,"ChangeTextboxes":False,"ChangeWordArt":False,"searchOptions":[]}
        self.initUI()

    def initUI(self):
        pz_chkbx=QCheckBox("Использовать подстановочные знаки")
        pz_chkbx.setObjectName("MatchWildcards")
        cv_chkbx=QCheckBox("Выделить цветом замененный текст")
        cv_chkbx.setObjectName("Highlight")
        gf_chkbx=QCheckBox("Выделить жирным шрифтом замененный текст")
        gf_chkbx.setObjectName("Bold")
        sh_chkbx=QCheckBox("Изменять текст в фигурах")
        sh_chkbx.setObjectName("ChangeTextboxes")
        wa_chkbx=QCheckBox("Изменять текст в объктах WordArt")
        wa_chkbx.setObjectName("ChangeWordArt")
        optbtn=QPushButton("Элементы поиска")
        optbtn.setObjectName("searchOptions")

        layout = QHBoxLayout()
        replaceLay=QVBoxLayout()
        replaceLay.addWidget(pz_chkbx)
        replaceLay.addWidget(cv_chkbx)
        replaceLay.addWidget(gf_chkbx)
        layout.addLayout(replaceLay)

        findLay=QVBoxLayout()
        findLay.addWidget(sh_chkbx)
        findLay.addWidget(wa_chkbx)
        findLay.addWidget(optbtn)
        layout.addLayout(findLay)

        self.setLayout(layout)
        self.searchOptions=searchOptions(self)
        self.searchOptions.setWindowModality(Qt.WindowModal)
        self.OptionsDict["searchOptions"]=self.searchOptions.WdStoryType

        for chbx in [pz_chkbx,cv_chkbx,gf_chkbx,sh_chkbx,wa_chkbx]:
            chbx.clicked.connect(self.changeOption)
        self.searchOptions.closed.connect(self.changeSearchOptions)
        optbtn.clicked.connect(self.searchOptions.show)

    def changeOption(self,checkState):
        key=self.sender().objectName()
        self.OptionsDict[key]=checkState
        #print(self.OptionsDict)
        #s=self.findChild(QCheckBox,"MatchWildcards")
        #print(s.checkState())
        #print(self.OptionsDict)

    def changeSearchOptions(self,WdStoryType):
        self.OptionsDict["searchOptions"]=WdStoryType

    def setOptions(self,Optionsdict={}):
        self.OptionsDict.update(Optionsdict)

        for key in Optionsdict.keys():
            CheckBox=self.findChild(QCheckBox,key)
            if CheckBox!=None: CheckBox.setCheckState(2 if Optionsdict[key]!=0 else 0)
        if Optionsdict.get("searchOptions")!=None:
            self.searchOptions.setOptions(Optionsdict["searchOptions"])
        #print(self.OptionsDict)

    def closeEvent(self, Event):
        #print(self.OptionsDict)
        Event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex=Options()
    ex.setOptions({"MatchWildcards":True,"Highlight":False,"Bold":False,"ChangeTextboxes":True,"ChangeWordArt":False,"searchOptions":[]})
    ex.show()
    sys.exit(app.exec_())