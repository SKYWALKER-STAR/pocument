# *coding=utf-8
from __future__ import annotations
import AbstractOutputFactory as ab
from docx.shared import RGBColor

def client_code(factory:ab.AbstractOutputFactory): 
    wordOut = factory.createFormDefault()

    wordOut.OutputTitle("门户系统",typeFace=u"黑体",level=1)
    wordOut.OutputTitle("1.非常规任务",typeFace=u"黑体",center=0,level=2)
    wordOut.OutputParagraph("门户系统的第一个自然段",left_ident=0.2)
    wordOut.PageBreak()

    wordOut.OutputTitle("服务平台",typeFace=u"黑体")
    wordOut.PageBreak()

    wordOut.OutputTitle("共享交换",typeFace=u"黑体")
    wordOut.PageBreak()

    wordOut.OutputTitle("数据资源管理",typeFace=u"黑体")
    wordOut.PageBreak()

    wordOut.OutputTitle("前置机",typeFace=u"黑体")

    #wordOut.TestStyle()



if __name__ == "__main__":
    client_code(ab.OutputForWord("./test.docx"))
