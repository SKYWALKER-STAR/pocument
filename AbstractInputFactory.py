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

    def getColumnNameArray(self):
        for i in self.df.columns.values:
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

if __name__ == '__main__':
    print("Abstrace factory for input")

