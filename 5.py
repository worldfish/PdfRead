import pdfplumber
import pandas as pd

# 创建一个空数据框
df = pd.DataFrame()

# 使用with语句打开pdf文件
with pdfplumber.open(r"D:\AS3000\PdfRead\20210512.pdf") as pdf:
        # 使用for循环遍历每个pages
    for page in pdf.pages:
        # for table in page.extract_tables():
        #         # 取出当前页表格，结果为列表
        #     d=page.extract_table()
        #    # 将列表转为数据框
        #     df1 = pd.DataFrame(d[1:], columns=d[0])
        #    #添加至df数据框中
        #     df = df.append(df1)
        for table in page.extract_tables():
            print(table)
            df = pd.DataFrame(table[1:], columns=table[0])
            for row in table:
                for cell in row:
                    # print(cell, end="\t|")
                    print(cell)
# 写入到Excel表中
df.to_excel(excel_writer = r"D:\AS3000\pdffile\a.xlsx",sheet_name = "从PDF中提取出来的表格数据")
# 写入到Excel表中并去掉Python中默认的索引
df.to_excel(excel_writer = r"D:\AS3000\pdffile\a1.xlsx",sheet_name = "从PDF中提取出来的表格数据1",index = False)

