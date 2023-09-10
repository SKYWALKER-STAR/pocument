>Author: ming 
>Email: 1546274931@qq.com 
>Usage: 用来处理日常报告，主要用在word、excel、ppt等office套件上 
>Create Date: 2023/8/31 
>Last Modify Date: 2023/9/10

# 0.介绍 
本程序是自动生成日报的程序
AbstractInputFactory.py文件中包含了处理输入文件的类;
AbstractOutputFactory.py文件中包含了处理输出文件的类;
start.py是命令行的启动脚本
tranui.py是GUI启动脚本

GUI启动的命令为：
python tranui.py

命令行启动的命令为：
python start.py --excel-source FILE1 --excel-source-sheet SHEETNAME --word-output FILE2 --report-title TITLE
其中，FILE1是excel表格输入，SHEETNAME是sheet名，默认为“Sheet1”，FILE2是输出文件，TITLE是生成的报告的标题.

为了方便扩展，本程序采用抽象工厂的设计模式,分别有AbstarctInput抽象类和AbstractOutput抽象类，如果需要增加不同类型、不同
格式的Input和Output，只需要新增一个继承上述两个抽象类的具体类即可。详细结构请参见架构图
