# *coding=utf-8

#This is file should be implemented by user them self using AbstractOutputFactory and AbstractInputFactory
from __future__ import annotations
import AbstractOutputFactory as abo
from docx.shared import RGBColor
import AbstractInputFactory as abi
import re
import argparse
import sys

#global noTask
taskName = ''
noTask = "今日无"
taskCount = 1

def taskNo(task):
    if task == "非常态化任务":
        task = "一、非常态化任务"
        return task
    elif task == "监控巡检":
        task = "二、监控巡检"
        return task
    elif task == "故障处理":
        task = "三、故障处理"
        return task
    elif task == "安全整改":
        task = "四、安全整改"
        return task
    elif task == "变更发版":
        task = "五、变更发版"
        return task

#According to content of excel write differendly using regular expression
def matchwrite(content,out:AbstractOutputFactory):
    keywords = {'非常态化任务','监控巡检','安全整改','变更发版','故障处理'}
    token_specification = [
                ('BINGO',r'[\u4e00-\u9fa5a-zA-Z]*'),
                ('CONTENT',r'[\u4e00-\u9fa5]|[\W]'),
                ('END',r'END'),
                ('SKIP',r'None'),
                ('NONE',r'NOTHING'),
                ]

    tok_regex = '|'.join('(?P<%s>%s)'% pair for pair in token_specification)
    line_num = 1
    line_start = 0

    contentList = []
    for mo in re.finditer(tok_regex,content):
        kind = mo.lastgroup
        value = mo.group()
        column = mo.start() - line_start
        if kind == 'BINGO':
            if value in keywords:
                out.OutputTitle(taskNo(value),level=3,center=0,size=12)
                global taskName
                taskName = value
            else:
                contentList.append(value)
        elif kind == 'END':
            global taskCount 
            taskCount = 1
            continue
        elif kind == 'NONE':
            global noTask
            content = noTask + taskName
            out.OutputParagraph(content)

    new_list = list(filter(lambda x:x.strip(),contentList))
    contentStr = ""
    if len(new_list) != 0:
        contentStr = ",".join(new_list)
        yield contentStr

def daliyReport(factoryi:abi.AbstractInputFactory,factoryo:abo.AbstractOutputFactory,title): 

    excelIn = factoryi.createFormDefault()
    wordOut = factoryo.createFormDefault()

    colNames = excelIn.getColumnNameArray()
    taskList = excelIn.getColumnContent

    wordOut.OutputTitle(title,level=1,size=18)
    i = 0
    for colName in colNames:
        if i == 0:
          wordOut.OutputTitle(colName,level=2,size=14)
          i = 1
        else:
          wordOut.PageBreak()
          wordOut.OutputTitle(colName,level=2,size=14)
        global taskCount
        taskCount = 1
        for task in taskList(colName):
            for content in matchwrite(task,wordOut):
                content = f"{taskCount}." + content
                wordOut.OutputParagraph(content)
                taskCount += 1
def exec(excel,word,title,excelSheet="Sheet1"):
    daliyReport(abi.InputForExcel(excel,excelSheet),
                abo.OutputForWord(word),title)
    return 0


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Input and Output")
    parser.add_argument('--excel-source',dest="exs",type=str,help='excel input',required=True)
    parser.add_argument('--excel-source-sheet',dest='exss',type=str,help='excel input sheet',required=True)
    parser.add_argument('--word-output',dest='wot',type=str,help='world output sheet',required=True)
    parser.add_argument('--report-title',dest='title',type=str,help='report title',required=True)
    args = parser.parse_args()

    exec(excel=args.exs,excelSheet=args.exss,word=args.wot,title=args.title)
