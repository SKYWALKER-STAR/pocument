# *coding=utf-8
from __future__ import annotations
from abc import ABC,abstractmethod
import pandas as pd
import numpy as np

#Abstract factories for input 
class AbstractInputFactory(ABC):
    @abstractmethod
    def createFormDefault(self):
        pass

#Concrete factories which inhere from AbstractInputFactory
class InputForExcel(AbstractInputFactory):
    def __init__(self,docPath=None,sheetName=None):
        self.docPath = docPath
        self.sheetName=sheetName
    def createFormDefault(self):
        return ExcelInputFormDefault(self.docPath,self.sheetName)
    def createExcelWeekReport(self,inputFileList=None,keyWords=None):
        return ExcelWeekReport(self.docPath,self.sheetName,inputFileList=inputFileList,keyWords=keyWords)

#Abstract input form for excel
class AbstractInputFormExcel(ABC):
    contents:list

    @abstractmethod
    def __init__(self,docPath=None):
        pass

    @abstractmethod
    def read(self):
        pass
    
#Concrete input form for excel:Default
#np.nan representing the missing data
class ExcelInputFormDefault(AbstractInputFormExcel):
    contents:list
    def __init__(self,docPath=None,sheetName=None):
        self.df = pd.read_excel(docPath,sheetName)
        self.savePath = docPath
        self.sheetName = sheetName
        self.contents = []

    def read(self): 
        print(self.df)
        return 0

    def getColumnNameArray(self,df=None,flag=0):
        if flag == 0:
           for i in self.df.columns.values:
              yield i
        else:
            for i in df.columns.values:
                yield i

    def matchName(self,pattern=None):
        pass

    def getColumnContent(self,columnName):
        for i in self.df[columnName]:
            if i is np.nan:
                continue
            else:
                yield i

    def getColLoc(self,colName):
        return self.df.columns.get_loc(colName)

    def getRowIndex(self,colName,taskName):
        return self.df.index[self.df[colName] == taskName]

    def getContentAccordingIndex(self,index):
        return self.df[index]

    def traversal(self,start,endFlag):
            rv = self.df.iloc(start)
            if rv == endFlag:
                yield -1
            else:
                yield rv

    def getColAccordingName(self,colName):
        for i in self.df.columns.values:
            if i == colName:
                return i

class ExcelWeekReport(ExcelInputFormDefault):

    def __init__(self,docPath=None,sheetName="Sheet1",keyWords=None,inputFileList=None):
        if docPath is not None:
            self.df = pd.read_excel(docPath,sheet_name=sheetName)

        self.fileList = inputFileList
        self.dfList = []
        self.sheetName = sheetName
        self.keyWords = keyWords

    def read(self):
        pass
    def save(self,df:pd.DataFrame,savePath,sheetName="Sheet1"):
        df.to_excel(savePath,sheetName)
    def saveFromList(self,seriData,savePath,sheetName="Sheet1"):
        df = pd.DataFrame(data=seriData)
        df.to_excel(savePath,sheetName)

    def getContent(self,df,colNames:list,keyWord=None):
        flag = 0
        infoList = []
        for i in colNames:
            infoDic = dict()
            infoDic['Name'] = i
            msgList = []
            for content in df[i]:
                if content == keyWord:
                    infoDic['Task'] = content
                    flag = 1
                elif content is np.nan:
                    continue
                elif (content == "END" or content == "NOTHING") and flag == 1:
                    infoList.append(infoDic)
                    break
                elif flag == 1:
                    msgList.append(content)
                    infoDic['Msg'] = msgList
        return infoList

    def getColumnNameArray(self,df=None,flag=0):
        liColList = []
        if flag == 0:
           for i in self.df.columns.values:
               liColList.append(i)
        else:
            for i in df.columns.values:
               liColList.append(i)
        return liColList

    def crossFileReadTask(self,colNames:list):
        for i in self.fileList:
            tmpdf = pd.read_excel(i,self.sheetName)
            rv = self.getContent(tmpdf,colNames)
            yield rv

    def getDfList(self,inputFileList):
        for i in inputFileList:
            yield pd.read_excel(i,"Sheet1")

if __name__ == '__main__':
    print("Abstrace factory for input")
