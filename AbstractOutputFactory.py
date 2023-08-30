# *coding=utf-8
from __future__ import annotations
from abc import ABC,abstractmethod
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

#Abstract factories for output
class AbstractOutputFactory(ABC):
    @abstractmethod
    def createFormDefault(self):
        pass

#Concrete factories which inhere from AbstractOutputFactory
class OutputForWord(AbstractOutputFactory):
    def __init__(self,docPath=None):
        self.docPath = docPath

    def createFormDefault(self): 
        return WordFormDefault(self.docPath)

#Abstract class for word output
class AbstractOutputFormWord(ABC):
    @abstractmethod
    def write(self):
        pass
    @abstractmethod
    def OutputTitle(self):
        pass
    @abstractmethod
    def OutputParagraph(self):
        pass

    def PageBreak(self):
        pass


class WordFormDefault(AbstractOutputFormWord):
    def __init__(self,docPath=None):
        self.document = Document()
        self.savePath = docPath

    def write(self,savePath=None):
        if savePath != None:
            self.document.save(savePath)
        else:
            self.document.save(self.savePath)

    def OutputTitle(self,content=None,level=1):
        if content != None:
            p = self.document.add_heading(content,level)
        else:
            p = self.document.add_heding("Default Heading",level)

        p.style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.write(savePath=None)

    def OutputParagraph(self,content=None):
        if content != None:
            p = self.document.add_paragraph(content)
        else:
            p = self.document.add_paragraph("Paragraph example")

        p.add_run(' bold ').bold = True
        p.add_run(' and some ')
        p.add_run(' italic. ').italic = True

        #p.style.font.size = Pt(20) font size
        self.write(savePath=None)

    def PageBreak(self):
         self.document.add_page_break()
