# *coding=utf-8
from __future__ import annotations
import AbstractOutputFactory as ab

def client_code(factory:ab.AbstractOutputFactory): 
    wordOut = factory.createFormDefault()

    wordOut.OutputTitle("This is My Main title")

    wordOut.OutputTitle("This is My Third title",level=3)

    wordOut.OutputTitle("This is My Second title",level=2)

    wordOut.OutputParagraph("Hello world from ming")

    #wordOut.TestStyle()



if __name__ == "__main__":
    client_code(ab.OutputForWord("./test.docx"))
