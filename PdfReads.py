import pdfplumber
import numpy as np
with pdfplumber.open(r'D:\AS3000\PdfRead\20210304.pdf') as pdf:
    #前提所有列必须固定
    table1= np.array([])
    table2 = np.array([])
    table3 = np.array([])
    table4 = np.array([ ])
    table5 = np.array([])
    table6 = np.array([])

    table1List1 = []  # 表1数据集合
    table1List2 = []  # 表1数据集合
    table1List3 = []  # 表1数据集合
    table1List4 = []  # 表1数据集合
    table1List5 = []  # 表1数据集合
    table1List6 = []  # 表1数据集合
    for page in pdf.pages:
        # page = pdf.pages[index]  # 页数 从0开始
        top1 = []
        index = 0
        for row in page.extract_table():
            if index == 0:
                for cell in row:
                    top1.append("".join(cell.split()))  # 将第一行的集合去除 \n
                if row[0] == "序号": #下一页 没有表头 则继续上一页为同一个表格
                    flag = 0
            else:
                # 1.正常读取行数 进行数据插入
                if flag == 1:  # 插入表1数据
                    table1List1.append(row)
                if flag == 2:  # 插入表1数据
                    table1List2.append(row)
                if flag == 3:  # 插入表1数据
                    table1List3.append(row)
                if flag == 4:  # 插入表1数据
                    table1List4.append(row)
                if flag == 5:  # 插入表1数据
                    table1List5.append(row)
                if flag == 6:  # 插入表1数据
                    table1List6.append(row)
            index = 1
            if flag==0 and len(table1)==len(top1) and (table1 == np.array(top1)).all():  # 判定是否为表1的数据
                flag = 1
            if flag==0 and len(table2)==len(top1) and (table2 == np.array(top1)).all():
                flag = 2
            if flag==0 and len(table3)==len(top1) and (table3 == np.array(top1)).all():
                flag = 3
            if flag==0 and len(table4)==len(top1) and (table4 == np.array(top1)).all():
                flag = 4
            if flag==0 and len(table5)==len(top1) and (table5 == np.array(top1)).all():
                flag = 5
            if flag==0 and len(table6)==len(top1) and (table6 == np.array(top1)).all():
                flag = 6

    print(table1List1)
    print(table1List2)
    print(table1List3)
    print(table1List4)
    print(table1List5)
    print(table1List6)

         # 2.特殊处理 同一张表在不同页上
            #采用序号对比来进行判定
         # 3.特殊处理 一条序号横跨多行。 即多级电厂数据

