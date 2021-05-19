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

#now_time = time.strftime("%Y%m%d_%H%M%S")
now_time = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
s = '1'
#global flag
flag = True
index = False

path = input("请输入PDF文件位置：")
#path = "aaaaaa.PDF"  # 导入PDF路径
pdf = pdfplumber.open(r"D:\AS3000\PdfRead\20210513.pdf")
print('\n')
print('开始读取数据')
print('\n')


for page in pdf.pages:

    # 获取当前页面的全部文本信息，包括表格中的文字
    print(page.extract_text())
    text = page.extract_text()
    workbook = xlwt.Workbook()  #定义workbook
    if text.find("附件"):
        sheet = workbook.add_sheet('Sheet1')  #添加sheet
        i = 0 # Excel起始位置
        #if flag == True:
        for table in page.extract_tables():#全部表格
            print(table)
            for row in table:#一个表格
                if row[0] != '电厂名称':
                    flag = False
                if row[0] == '电厂名称':
                    flag = True
                print(row)
                for j in range(len(row)):#一行
                    sheet.write(i, j, row[j])#一列一列
                i += 1
            print('---------- 分割线 ----------')
        workbook.save(r'D:/AS3000/pdffile/'+s+'.xls')
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
