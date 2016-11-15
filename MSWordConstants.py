#from enum import Enum,unique,IntEnum
wdFindContinue = 1
wdFindAsk = 2
wdFindStop = 0

wdReplaceNone = 0
wdReplaceOne = 1
wdReplaceAll = 2

msoTextEffect = 15
wdYellow = 7
wdColorBlack=0

"""class WdFindWrap(IntEnum):
    wdFindStop = 0
    wdFindContinue = 1
    wdFindAsk = 2

class WdReplace(IntEnum):
    wdReplaceNone = 0
    wdReplaceOne = 1
    wdReplaceAll = 2

class WdColorIndex(IntEnum):
    __doc__="Индексы цвета MSWord"
    wdAuto = 0
    wdBlack = 1
    wdBlue = 2
    wdBrightGreen = 4
    wdByAuthor = -1 #(&HFFFFFFFF)
    wdDarkBlue = 9
    wdDarkRed = 13
    wdDarkYellow = 14
    wdGray25 = 16 #(&H10)
    wdGray50 = 15
    wdGreen = 11
    wdNoHighlight = 0
    wdPink = 5
    wdRed = 6
    wdTeal = 10
    wdTurquoise = 3
    wdViolet = 12
    wdWhite = 8
    wdYellow = 7

@unique
class WdStoryType(IntEnum):
    __doc__ = "Константа WdStoryType – Возвращает тип истории"
    wdMainTextStory=1
    wdFootnotesStory=2
    wdEndnotesStory=3
    wdCommentsStory=4
    wdTextFrameStory=5
    wdEvenPagesHeaderStory=6
    wdPrimaryHeaderStory=7
    wdEvenPagesFooterStory=8
    wdPrimaryFooterStory=9
    wdFirstPageHeaderStory=10
    wdFirstPageFooterStory=11
    wdFootnoteSeparatorStory=12
    wdFootnoteContinuationSeparatorStory=13
    wdFootnoteContinuationNoticeStory=14
    wdEndnoteSeparatorStory=15
    wdEndnoteContinuationSeparatorStory=16
    wdEndnoteContinuationNoticeStory=17

@unique
class WdInlineShapeType(IntEnum):
    wdInlineShapeEmbeddedOLEObject=1
    wdInlineShapeLinkedOLEObject=2
    wdInlineShapePicture=3
    wdInlineShapeLinkedPicture=4
    wdInlineShapeOLEControlObject=5
    wdInlineShapeHorizontalLine=6
    wdInlineShapePictureHorizontalLine=7
    wdInlineShapeLinkedPictureHorizontalLine=8
    wdInlineShapePictureBullet=9
    wdInlineShapeScriptAnchor=10
    wdInlineShapeOWSAnchor=11
    wdInlineShapeChart=12
    wdInlineShapeDiagram=13
    wdInlineShapeLockedCanvas=14
    wdInlineShapeSmartArt=15

@unique
class MsoShapeType(IntEnum):
    msoShapeTypeMixed=-2
    msoAutoShape=1
    msoCallout=2
    msoChart=3
    msoComment=4
    msoFreeform=5
    msoGroup=6
    msoEmbeddedOLEObject=7
    msoFormControl=8
    msoLine=9
    msoLinkedOLEObject=10
    msoLinkedPicture=11
    msoOLEControlObject=12
    msoPicture=13
    msoPlaceholder=14
    msoTextEffect=15
    msoMedia=16
    msoTextBox=17
    msoScriptAnchor=18
    msoTable=19
    msoCanvas=20
    msoDiagram=21
    msoInk=22
    msoInkComment=23
    msoIgxGraphic=24
    msoSlicer=25
    msoWebVideo=26
    msoContentApp=27"""

WdStoryType={
            1:["wdMainTextStory","Основной текст",2],
            2:["wdFootnotesStory","Сноски",2],
            3:["wdEndnotesStory","Примечания",2],
            4:["wdCommentsStory","Комментарии",2],
            5:["wdTextFrameStory","Текст рамки",2],
            6:["wdEvenPagesHeaderStory","Четные страницы заголовка",2],
            7:["wdPrimaryHeaderStory","Верхний колонтитул",2],
            8:["wdEvenPagesFooterStory","Четные страницы нижнего колонтитула",2],
            9:["wdPrimaryFooterStory","Нижний колонтитул",2],
            10:["wdFirstPageHeaderStory","Верхний колонтитул первой страницы",2],
            11:["wdFirstPageFooterStory","Нижний колонтитул первой страницы",2],
            12:["wdFootnoteSeparatorStory","Нижняя сноска разделителя",2],
            13:["wdFootnoteContinuationSeparatorStory","Сноска продолжения разделителя",2],
            14:["wdFootnoteContinuationNoticeStory","Сноска продолжения уведомления",2],
            15:["wdEndnoteSeparatorStory","Сноска разделитель",2],
            16:["wdEndnoteContinuationSeparatorStory","Сноска продолжения разделителя",2],
            17:["wdEndnoteContinuationNoticeStory","Сноска продолжения уведомления",2]
            }
WdInlineShapeType={
                    1:["wdInlineShapeEmbeddedOLEObject","",1],
                    2:["wdInlineShapeLinkedOLEObject","",1],
                    3:["wdInlineShapePicture","",1],
                    4:["wdInlineShapeLinkedPicture","",1],
                    5:["wdInlineShapeOLEControlObject","",1],
                    6:["wdInlineShapeHorizontalLine","",1],
                    7:["wdInlineShapePictureHorizontalLine","",1],
                    8:["wdInlineShapeLinkedPictureHorizontalLine","",1],
                    9:["wdInlineShapePictureBullet","",1],
                    10:["wdInlineShapeScriptAnchor","",1],
                    11:["wdInlineShapeOWSAnchor","",1],
                    12:["wdInlineShapeChart","",0],
                    13:["wdInlineShapeDiagram","",0],
                    14:["wdInlineShapeLockedCanvas","",1],
                    15:["wdInlineShapeSmartArt","",0]
                    }
MsoShapeType={
                -2:["msoShapeTypeMixed","Фигура смешанного типа",2],
                1:["msoAutoShape","Автофигура",2],
                2:["msoCallout","Выноска",1],
                3:["msoChart","График",1],
                4:["msoComment","Комментарий",2],
                5:["msoFreeform","Свободная форма",1],
                6:["msoGroup","Группа",2],
                7:["msoEmbeddedOLEObject","Встроенный OLE объект",1],
                8:["msoFormControl","Элемент управления формы",1],
                9:["msoLine","Линия",1],
                10:["msoLinkedOLEObject","Связанный OLE объект",1],
                11:["msoLinkedPicture","Связанное фото",1],
                12:["msoOLEControlObject","OLE объект управления",1],
                13:["msoPicture","Картинка",1],
                14:["msoPlaceholder","Заполнитель",1],
                15:["msoTextEffect","Текст эффект",2],
                16:["msoMedia","Медиа",1],
                17:["msoTextBox","Текстовое поле",1],
                18:["msoScriptAnchor","Ссылка на скрипт",1],
                19:["msoTable","Таблица",1],
                20:["msoCanvas","Полотно(Холст)",1],
                21:["msoDiagram","Диаграмма",1],
                22:["msoInk","Чернила",1],
                23:["msoInkComment","Чернила коментарий",1],
                24:["msoIgxGraphic"," SmartArt(Чернила) графика",1],
                25:["msoSlicer","Разделитель",1],
                26:["msoWebVideo","Веб-видео",1],
                27:["msoContentApp","Содержимое приложения",1]
                }


"""import sys

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
class ComboDelegate(QItemDelegate):
    editorItems=['Combo_Zero', 'Combo_One','Combo_Two']
    height = 25
    width = 200
    def createEditor(self, parent, option, index):
        editor = QListWidget(parent)
        # editor.addItems(self.editorItems)
        # editor.setEditable(True)
        editor.currentItemChanged.connect(self.currentItemChanged)
        return editor

    def setEditorData(self,editor,index):
        z = 0
        for item in self.editorItems:
            ai = QListWidgetItem(item)
            editor.addItem(ai)
            if item == index.data():
                editor.setCurrentItem(editor.item(z))
            z += 1
        editor.setGeometry(0,index.row()*self.height,self.width,self.height*len(self.editorItems))

    def setModelData(self, editor, model, index):
        editorIndex=editor.currentIndex()
        text=editor.currentItem().text() 
        model.setData(index, text)
        # print '\t\t\t ...setModelData() 1', text

    @pyqtSlot()
    def currentItemChanged(self): 
        self.commitData.emit(self.sender())"""
"""class MyModel(QAbstractTableModel):
    def __init__(self, parent=None, *args):
        QAbstractTableModel.__init__(self, parent, *args)
        self.items=[]
        keyList=list(WdStoryType.keys())
        keyList.sort()
        n=len(keyList)//2
        ost=len(keyList)%2
        n+=ost
        keyList1=keyList[:n]
        keyList2=keyList[n:]
        if ost==1:keyList2.append("")
        print(keyList1,keyList2)
        for i,key1 in enumerate(keyList1):
            key2=keyList2[i]
            val1=WdStoryType.get(key1)[1] if WdStoryType.get(key1)[1]!="" else WdStoryType.get(key1)[0]
            val2=WdStoryType.get(key2)[1] if WdStoryType.get(key2)[1]!="" else WdStoryType.get(key2)[0]
            self.items.append([val1,val2])


    def rowCount(self, QModelIndex_parent=None, *args, **kwargs):
        return len(self.items)

    def columnCount(self, QModelIndex_parent=None, *args, **kwargs):
        return len(self.items[0])

    def data(self, index, role):
        if not index.isValid(): return QVariant()

        row=index.row()
        col=index.column()

        item=self.items[row][col]

        if row>len(self.items): return QVariant()
        if col>len(self.items[0]): return QVariant()

        if role == Qt.DisplayRole:
            return QVariant(item)

    def flags(self, index):

        return Qt.ItemIsEditable | Qt.ItemIsEnabled |Qt.ItemIsUserCheckable

    def setData(self,index, value,Role):
        print(value,Role)
        self.items[index.row()][index.column()]=value
        return True

if __name__ == '__main__':
    app = QApplication(sys.argv)

    model = MyModel()
    tableView = QTableView()
    tableView.setModel(model)

    #delegate = ComboDelegate()

    #tableView.setItemDelegate(delegate)
    for i in range(0,tableView.model().rowCount()):
        tableView.setRowHeight(i,tableView.itemDelegate().height)
    for i in range(0,tableView.model().columnCount()):
        tableView.setColumnWidth(i,tableView.itemDelegate().width)"""
"""tableView.show()
    sys.exit(app.exec_())"""