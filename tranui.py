from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
import start
import sys

class findFile(QWidget):
    listCount = 1
    def __init__(self,parent = None):
         super(findFile,self).__init__(parent)
         layout=QVBoxLayout()
         layout_2=QHBoxLayout()

         self.inputFile = None
         self.inputFileList = []
         self.outPutFile = None
         self.Title = None
         self.dia = QFileDialog()

         #Set window style

         self.resize(400,200)

         #Create and Connect two buttons
         self.btnGetInputFile = QPushButton("输入(日报)")
         self.btnGetOutputFile = QPushButton("输出(日报)")
         self.btnSelectFileForWeek = QPushButton("周报(输入)")
         self.btnClearTextEdit = QPushButton("清除")
         self.btnConfirmTitle = QPushButton("确认标题")
         self.btnExcuteDaliy = QPushButton("生成日报")
         self.btnExcuteWeekly = QPushButton("生成周报")
         self.btnQuit = QPushButton("退出") #Quit program

         #Set a text edit
         self.textEdit = QTextEdit()

         #Set a line edit
         self.lineEdit = QLineEdit()

         #Connect to click signal
         self.btnGetInputFile.clicked.connect(self.GetInputFile)
         self.btnGetOutputFile.clicked.connect(self.GetOutputFile)
         self.btnClearTextEdit.clicked.connect(self.textEdit.clear)
         self.btnClearTextEdit.clicked.connect(self.lineEdit.clear)
         self.btnClearTextEdit.clicked.connect(self.clearAll)
         self.btnConfirmTitle.clicked.connect(self.ConfirmTitle)
         self.btnQuit.clicked.connect(self.Quit)
         self.btnExcuteDaliy.clicked.connect(self.ExcuteDaliy)
         self.btnSelectFileForWeek.clicked.connect(self.GetInputFileList)
         self.btnExcuteWeekly.clicked.connect(self.ExcuteWeekly)

         #Add buttons to layout
         layout.addWidget(self.btnGetInputFile)
         layout.addWidget(self.btnGetOutputFile)
         layout.addWidget(self.btnSelectFileForWeek)
         layout.addWidget(self.btnClearTextEdit)

         #Add Line Edit
         layout.addWidget(self.lineEdit)
         layout.addWidget(self.btnConfirmTitle)

         #Add text edit to layout
         layout.addWidget(self.textEdit)
         layout.addWidget(self.btnExcuteDaliy)
         layout.addWidget(self.btnExcuteWeekly)
         layout.addWidget(self.btnQuit)

         #Set layout
         self.setLayout(layout)
         self.setWindowTitle("报告生成器")
         self.labelGetfile=QLabel("")

    #Wrap textEdit insert
    def __writetext(self,content):
        self.textEdit.textCursor().insertText(content)

    def GetInputFile(self):
        fname,_=self.dia.getOpenFileName(self,'Open file',"/home/ming")
        self.inputFile = fname
        self.__writetext("Input:\r\n" + self.inputFile + "\r\n")

    def GetOutputFile(self):
        fname,_=self.dia.getOpenFileName(self,'Open file',"/home/ming")
        self.outputFile = fname
        self.__writetext("Output:\r\n" + fname + "\r\n")

    #Select a series of files
    def GetInputFileList(self):
        if self.dia.exec():
            filenames = self.dia.selectedFiles()
            self.inputFileList.append(filenames[0])
            if self.listCount == 1:
                self.__writetext("Files List:\r\n")
            self.__writetext(f"{self.listCount}." + filenames[0] + "\r\n")
            self.listCount += 1

    def GetReportTitle(self):
        return self.lineEdit.text()

    def ConfirmTitle(self):
        self.Title = self.lineEdit.text()
        self.__writetext("Title:\r\n" + self.Title + "\r\n")

    def ExcuteDaliy(self):
        start.exec_daliy(excel=self.inputFile,word=self.outputFile,title=self.Title)
        self.__writetext("\"" + self.Title + "\"" + "generated,please check!\r\n")
        self.listCount = 1

    def ExcuteWeekly(self):
        #start.exec_weekly(excel=self.inputFile,word=self.outputFile,title=self.Title,inputFileList=self.inputFileList)
        #start.exec_weekly(excelIn=self.inputFile,inputFileList=self.inputFileList)
        start.exec_weekly(inputFileList=self.inputFileList)
        self.__writetext("\"" + self.Title + "\"" + "generated,please check!\r\n")
        self.listCount = 1

    def clearAll(self):
        self.listCount = 1
        self.inputFileList = []
        self.outPutFile = None
        self.Title = None

    def Quit(self):
        sys.exit()

if __name__ == "__main__":
     app=QApplication(sys.argv)
     window=findFile()
     window.show()
     app.exec()
