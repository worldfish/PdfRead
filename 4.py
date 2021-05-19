# -*- coding: utf-8 -*-
"""
Created on 0721 2019
@author: Chan
请确保你在运行这个代码的时候，已经安装了pdfplumber库
如果没有安装，请在[附件-命令提示符]下输入：
pip install pdfplumber
"""
import time

import pdfplumber
import xlwt


# 定义保存Excel的位置

now_time = time.strftime("%Y%m%d_%H%M%S")
s = '1'


path = input("请输入PDF文件位置：")
#path = "aaaaaa.PDF"  # 导入PDF路径
pdf = pdfplumber.open(r"D:\AS3000\PdfRead\20210512.pdf")
print('\n')
print('开始读取数据')
print('\n')


for page in pdf.pages:

    # 获取当前页面的全部文本信息，包括表格中的文字
    print(page.extract_text())
    page.e
    workbook = xlwt.Workbook()  #定义workbook
    for table in page.extract_tables():#全部表格
        print(table)
        i = 0 # Excel起始位置
        sheet = workbook.add_sheet('Sheet1')  #添加sheet
        for row in table:#一个表格
            print(row)
            for j in range(len(row)):#一行
                sheet.write(i, j, row[j])#一列一列
            i += 1
        print('---------- 分割线 ----------')
        workbook.save('D:/AS3000/pdffile/'+s+'.xls')
        s += '1'


pdf.close()


# 保存Excel表
#workbook.save('D:/AS3000/文件名.xls')
#print('\n')
#print('写入excel成功')
#print('保存位置：')
#print('保存路径/文件名.xls')
#print('\n')
input('PDF取读完毕，按任意键退出')
