#### Author: ming  
#### Email: 1546274931@qq.com  
#### Usage: 用来处理日常报告，主要用在word、excel、ppt等office套件上  
#### Create Date: 2023/8/31  
#### Last Modify Date: 2024/1/16  

# I.概述
## 0.介绍

本脚本最主要使用到的库有PyQT6、PySide6、pandas、numpy、python-docx.

其中，PyQT6和PySide6用来完成图形化界面。因为它们是通用的GUI框架，因此，
本脚本的GUI界面不受操作系统限制。

pandas与numpy用来完成Excel的数据处理,python-docx用来完成Word文档的处理

## 1. 安装
执行以下命令,安装本脚本需要的依赖

`pip install -r requirments.txt`

## 2. 本程序是自动生成日报的程序

`AbstractInputFactory.py` 文件中包含了处理输入文件的类;

`AbstractOutputFactory.py` 文件中包含了处理输出文件的类;

`start.py` 是命令行启动脚本;

`tranui.py` 是GUI启动脚本;

## 3. GUI启动命令
`python tranui.py`

## 4. 命令行启动命令

`python start.py --excel-source FILE1 --excel-source-sheet SHEETNAME --word-output FILE2 --report-title TITLE`

其中，FILE1是excel表格输入，SHEETNAME是sheet名，默认为“Sheet1”，FILE2是输出文件，TITLE是生成的报告的标题.

## 5. 扩展

为了方便扩展，本程序采用抽象工厂的设计模式,分别有AbstarctInput抽象类和AbstractOutput抽象类，如果需要增加不同类型、不同

格式的Input和Output，只需要新增一个继承上述两个抽象类的具体类即可。详细结构请参见架构图
