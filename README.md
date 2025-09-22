# GET2.0
to get the calculation of the weighted average easier.

该脚本旨在简化全年级同学加权平均分计算的excel操作。web页面上传包含成绩表.xlsx（列：学号，姓名，各课程及成绩），可视化操作设置加权平均分计算规则，运行得到结果表.xlsx（列：学号，姓名，加权平均分，所有选择的课程及成绩）

Web驱动自动化：搭建交互式web页面，前后端分离架构。网页接收文件上传，app.py采用Flask框架作为核心调度器，组织对应的路由处理逻辑。
## QuickStart Guide
环境：win10+python3.8.10+vscode

1.下载代码，创建并激活虚拟环境venv，安装python包（pandas、openpyxl、flask）
```
py -m venv venv
.\venv\Scripts\activate
py -m pip install pandas
py -m pip install openpyxl
py -m pip install flask
```
2.启动本地服务器
```
py app.py
```
3.打开web页面，开始使用。
