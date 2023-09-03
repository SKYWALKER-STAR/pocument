from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
import start
import sys

class findFile(QWidget):
    def __init__(self,parent = None):
         super(findFile,self).__init__(parent)
         layout=QVBoxLayout()

         self.inputFile = None
         self.outPutFile = None
         self.Title = None

         #Create and Connect two buttons
         self.btnGetInputFile = QPushButton("Input")
         self.btnGetOutputFile = QPushButton("Output")
         self.btnClearTextEdit = QPushButton("Clear")
         self.btnConfirmTitle = QPushButton("Confirm Title")
         self.btnExcute = QPushButton("Excute")
         self.btnQuit = QPushButton("Quit") #Quit program

         #Set a text edit
         self.textEdit = QTextEdit()

         #Set a line edit
         self.lineEdit = QLineEdit()

         #Connect to click signal
         self.btnGetInputFile.clicked.connect(self.GetInputFile)
         self.btnGetOutputFile.clicked.connect(self.GetOutputFile)
         self.btnClearTextEdit.clicked.connect(self.textEdit.clear)
         self.btnClearTextEdit.clicked.connect(self.lineEdit.clear)
         self.btnConfirmTitle.clicked.connect(self.ConfirmTitle)
         self.btnQuit.clicked.connect(self.Quit)
         self.btnExcute.clicked.connect(self.Excute)


         #Add buttons to layout
         layout.addWidget(self.btnGetInputFile)
         layout.addWidget(self.btnGetOutputFile)
         layout.addWidget(self.btnClearTextEdit)

         #Add Line Edit
         layout.addWidget(self.lineEdit)
         layout.addWidget(self.btnConfirmTitle)

         #Add text edit to layout
         layout.addWidget(self.textEdit)
         layout.addWidget(self.btnExcute)
         layout.addWidget(self.btnQuit)

         #Set layout
         self.setLayout(layout)
         self.setWindowTitle("File Open")

         self.labelGetfile=QLabel("")


    #Wrap textEdit insert
    def __writetext(self,content):
        self.textEdit.textCursor().insertText(content)

    def GetInputFile(self):
        fname,_=QFileDialog.getOpenFileName(self,'Open file',"/home/ming")
        self.inputFile = fname
        self.__writetext("Input:\r\n" + self.inputFile + "\r\n")

    def GetOutputFile(self):
        fname,_=QFileDialog.getOpenFileName(self,'Open file',"/home/ming")
        self.outputFile = fname
        self.__writetext("Output:\r\n" + fname + "\r\n")

    def GetReportTitle(self):
        return self.lineEdit.text()

    def ConfirmTitle(self):
        self.Title = self.lineEdit.text()
        self.__writetext("Title:\r\n" + self.Title + "\r\n")

    def Excute(self):
        start.exec(excel=self.inputFile,word=self.outputFile,title=self.Title)
        self.__writetext("\"" + self.Title + "\"" + " generated,please check!\r\n")

    def Quit(self):
        sys.exit()

if __name__ == "__main__":
     app=QApplication(sys.argv)
     window=findFile()
     window.show()
     app.exec()
