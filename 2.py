# 导入pdfplumber
import pandas as pd
import pdfplumber

pdf =  pdfplumber.open("E:\\nba.pdf")# 读取pdf文件，保存为pdf实例

# 访问第二页
first_page = pdf.pages[1]

    # 自动读取表格信息，返回列表
    table = first_page.extract_table()

    table
# 将列表转为df
table_df = pd.DataFrame(table_2[1:],columns=table_2[0])

# 保存excel
table_df.to_excel('test.xlsx')

table_df
