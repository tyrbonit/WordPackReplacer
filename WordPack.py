#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys,os
import win32com.client as win32
from PyQt5.QtWidgets import (QMainWindow, QAction, QApplication,QWidget,
                             QVBoxLayout,QPushButton,QCheckBox,QMessageBox,QProgressDialog)
from PyQt5.QtGui import QIcon,QStandardItem
from PyQt5.QtCore import Qt,QSize
from MSWordConstants import wdFindContinue,wdReplaceAll,wdYellow,msoTextEffect,WdStoryType,WdInlineShapeType,MsoShapeType,wdColorBlack

from MainWidget import MainWidget
from FilesTableWidget import FilesTableWidget
from ReplaceTableWidget import ReplaceTableWidget
from statusTable import statusTableWidget

class UnselectOptionWidget(QWidget):

    def __init__(self,parent=None,Fileslist=[]):
        super(UnselectOptionWidget, self).__init__(parent)
        self.Fileslist=Fileslist
        self.initUI()
        self.setGeometry(500, 200, 300, 200)
        self.setWindowTitle("Опции")
        self.setWindowIcon(QIcon('icons\\highlight.png'))
        self.setWindowFlags(Qt.Dialog|Qt.WindowMinMaxButtonsHint|Qt.WindowCloseButtonHint)
        self.show()

    def initUI(self):
        Vbox = QVBoxLayout(self)
        check1=QCheckBox("Снять выделение текста цветом")
        check2=QCheckBox("Установить черный цвет шрифта")
        button=QPushButton("ОК")
        #progress=QProgressDialog()
        #progress.setHidden(True)
        Vbox.addWidget(check1)
        Vbox.addWidget(check2)
        #Vbox.addWidget(progress)
        Vbox.addWidget(button)

        self.setLayout(Vbox)
        self.check1=check1
        self.check2=check2
        self.button=button

        button.clicked.connect(self.UnSelectText)

    def UnSelectText(self):
        ch1=self.check1.checkState()==2
        ch2=self.check2.checkState()==2
        SelectedWdStoryType=WdStoryType.keys()

        if self.Fileslist!=[] and (ch1 or ch2):
            app = win32.Dispatch("Word.Application")
            app.Visible = 0
            app.DisplayAlerts = 0
            for i,file in enumerate(self.Fileslist):
                doc=app.Documents.Open(file)
                for oRngStory in doc.StoryRanges:
                    if oRngStory.StoryType in SelectedWdStoryType:
                        if ch1:oRngStory.HighlightColorIndex = 0
                        if ch2:oRngStory.Font.Color=wdColorBlack

                    #Надписи, расположенные поверх текста
                    if oRngStory.ShapeRange.Count > 0:
                        for oShp in oRngStory.ShapeRange:
                            self.UselectTextInShape(oShp,ch1,ch2)

                doc.Close(SaveChanges=True)
            app.Quit()
        self.close()

    def UselectTextInShape(self,oShp,ch1,ch2):
        if oShp.TextFrame.HasText:
            if ch1:oShp.TextFrame.TextRange.HighlightColorIndex = 0
            if ch2:oShp.TextFrame.TextRange.Font.Color=wdColorBlack
        elif oShp.Type==6:
            for i in range(1,oShp.GroupItems.Count+1):
                self.UselectTextInShape(oShp.GroupItems.Item(i),ch1,ch2)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        FWidget=FilesTableWidget(self)
        RWidget=ReplaceTableWidget(self)
        SWidget=statusTableWidget(self)
        MWidget=MainWidget(self)
        MWidget.LeftLay.addWidget(FWidget)
        MWidget.RightLay.addWidget(SWidget)
        MWidget.BottomLay.addWidget(RWidget)
        self.setCentralWidget(MWidget)
        #QActions

        exitAction = QAction(QIcon('icons\\exit.jpg'), 'Выход', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Закрыть приложение')
        exitAction.triggered.connect(self.close)

        AboutAction=QAction(QIcon('icons\\Question.png'), 'О программе', self)
        AboutAction.setShortcut('Ctrl+F1')
        AboutAction.setStatusTip('О программе')
        AboutAction.triggered.connect(self.About)

        DelSelectionText=QAction(QIcon('icons\\highlight.png'), 'Снять &выделение', self)
        DelSelectionText.setShortcut('Ctrl+Shift+A')
        DelSelectionText.setStatusTip('Снять выделение цветом для выбранных файлов')
        DelSelectionText.triggered.connect(self.UnSelectText)

        ReplaceAction=QAction(QIcon('icons\\search-replace.png'), 'Найти и заменить', self)
        ReplaceAction.setShortcut('Ctrl+R')
        ReplaceAction.setStatusTip('Приступить к замене в отмеченных файлах')
        ReplaceAction.triggered.connect(self.startReplace)

        SelectPathAction=QAction(QIcon('icons\\pack.png'), 'Выбрать папку', self)
        SelectPathAction.setShortcut('Ctrl+P')
        SelectPathAction.setStatusTip('Выбрать папку с документами')
        SelectPathAction.triggered.connect(FWidget.selectPath)

        #/QActions
        #menuBar
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&Файл')
        fileMenu.addAction(DelSelectionText)
        fileMenu.addAction(exitAction)

        helpMenu = menubar.addMenu('&Справка')
        helpMenu.addAction(AboutAction)
        #/menuBar

        toolbar = self.addToolBar('Панель инструментов')
        #QToolBar
        toolbar.setIconSize(QSize(50,50))
        toolbar.addAction(exitAction)
        toolbar.addAction(DelSelectionText)
        toolbar.addAction(ReplaceAction)
        toolbar.addAction(SelectPathAction)
        toolbar.addAction(AboutAction)

        self.MWidget=MWidget
        self.SWidget=SWidget
        self.RWidget=RWidget
        self.FWidget=FWidget
        self.FilesList=self.FWidget.FilesList
        self.ReplaceList=self.RWidget.ReplaceList

        self.statusBar()
        self.resize(1024,768)
        desktop=QApplication.desktop()
        x=(desktop.width()-self.frameSize().width())//2
        y=(desktop.height()-self.frameSize().height())//2
        self.move(x, y)
        self.setWindowTitle('Обработчик пакетов Word документов')
        self.setWindowIcon(QIcon('icons\\pack.png'))
        self.show()

    def About(self):
        QMessageBox.about(self,"О программе",
        """Программа выполняет поиск и замену текста в нескольких документах MS Office
        Разработал: Москвитин И.В.
        Версия 1.3 2016г.
        Email: ityrbonit@yandex.ru""")

    def UnSelectText(self):
        FilesList=self.FilesList
        UnselectOptionWidget(self,FilesList)

    def startReplace(self):
        self.SWidget.clear()
        app = win32.Dispatch("Word.Application")
        app.Options.DefaultHighlightColorIndex = wdYellow
        app.Visible = False
        app.DisplayAlerts = False
        self.progress=QProgressDialog()
        progress=self.progress
        progress.setMinimum(0)
        progress.setMaximum(len(self.FilesList))
        progress.show()

        for i,file in enumerate(self.FilesList):
            if progress.wasCanceled(): break
            progress.setValue(i)
            progress.setLabelText("Отрыт файл: "+os.path.split(file)[1])
            doc=app.Documents.Open(file)
            for oRngStory in doc.StoryRanges:

                for find_str,replace_str,options in self.ReplaceList:
                    MatchWildcards=options["MatchWildcards"]
                    Highlight=options["Highlight"]
                    Bold=options["Bold"]
                    ChangeTextboxes=options["ChangeTextboxes"]
                    ChangeWordArt=options["ChangeWordArt"]
                    SelectedWdStoryType=options["searchOptions"]
                    ReplaceParamList=[find_str,replace_str,MatchWildcards,Highlight,Bold,ChangeTextboxes,ChangeWordArt]

                    if find_str=='': continue
                    if oRngStory.StoryType in SelectedWdStoryType:
                        fnd=self.ReplaceInWordRange(oRngStory,ReplaceParamList)
                        self.SWidget.appendRow(file,find_str,replace_str,WdStoryType[oRngStory.StoryType][1],fnd )
                    #Надписи, расположенные поверх текста
                    if oRngStory.ShapeRange.Count > 0 and ChangeTextboxes:
                        for oShp in oRngStory.ShapeRange:
                            fnd=self.ReplaceInShape(oShp,ReplaceParamList)
                            self.SWidget.appendRow(file,find_str,replace_str,MsoShapeType[oShp.Type][0],fnd)
                    #Замена в объектах WordArt
                    if oRngStory.InlineShapes.Count > 0 and ChangeWordArt:
                        for oInShp in oRngStory.InlineShapes:
                            try:
                                if oInShp.Type!=3: continue
                                Text=oInShp.TextEffect.Text
                                if oInShp.TextEffect.Text!= "":
                                    oInShp.TextEffect.Text = Text.replace(find_str, replace_str)
                                    fnd=replace_str in oInShp.TextEffect.Text
                                    self.SWidget.appendRow(file,find_str,replace_str,WdInlineShapeType[oInShp.Type][0],fnd)
                            except:
                                continue
            doc.Close(SaveChanges=True)
        progress.close()
        app.Quit()

    def ReplaceInWordRange(self,WordRangeObj,ReplaceParamList):
        find_str,replace_str,MatchWildcards,Highlight,Bold,ChangeTextboxes,ChangeWordArt=ReplaceParamList

        Find=WordRangeObj.Find

        find_str=find_str.replace("\n","^0013") if MatchWildcards else find_str.replace("\n","^p")
        replace_str=replace_str.replace("\n","^p")

        Find.ClearFormatting()
        Find.Replacement.ClearFormatting()
        Find.Replacement.Highlight = Highlight #Выделение цветом
        Find.Replacement.Font.Bold=Bold
        #Почему то при компиляции в ехе работает только нижеследующий метод - когда все параметры в аргументах метода .Execute
        return Find.Execute(FindText=find_str,MatchCase=False, MatchWholeWord=False, MatchWildcards=MatchWildcards, MatchSoundsLike=False, MatchAllWordForms=False, \
            Forward=True, Wrap=wdFindContinue, Format=True, ReplaceWith=replace_str, Replace=wdReplaceAll)

    def ReplaceInShape(self,oShp,ReplaceParamList):
        find_str,replace_str,MatchWildcards,Highlight,Bold,ChangeTextboxes,ChangeWordArt=ReplaceParamList
        found1,found2,found3=False,False,False
        if oShp.TextFrame.HasText:
            found1=self.ReplaceInWordRange(oShp.TextFrame.TextRange,ReplaceParamList)

        elif oShp.Type == msoTextEffect and ChangeWordArt:
            oShp.TextEffect.Text=oShp.TextEffect.Text.replace(find_str, replace_str)
            found2=replace_str in oShp.TextEffect.Text
        elif oShp.Type==6:
            #print(oShp.GroupItems.Count)
            for i in range(1,oShp.GroupItems.Count+1):
                #print(i)
                found3=self.ReplaceInShape(oShp.GroupItems.Item(i),ReplaceParamList)
        return True in [found1, found2, found3]

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())

