import os
import re
import time
from openpyxl import Workbook
import pdfplumber
import pandas as pd
# 这是返回文件的绝对路径写法

PATH = lambda p: os.path.abspath(
    os.path.join(os.path.dirname(__file__), p)
)
document_path = PATH('./')
list1 = []
for file in os.listdir(document_path):
    if file.endswith(".pdf"):
        if file not in list1:
            list1.append(file)


# results_name可以用时间戳创建，mode='a'是追加
now_time = time.strftime("%Y%m%d_%H%M%S")
df_first = pd.DataFrame()
results_name = 'results_%s.xlsx' % now_time
df_first.to_excel(results_name)
#writer = pd.ExcelWriter(results_name, mode='a', engine='openpyxl')
writer = pd.ExcelWriter(results_name,mode='a',engine='openpyxl')


for i in list1:
    pdf = pdfplumber.open(PATH(i))
    for page in pdf.pages:
        if page.page_number == 3:
            content = page.extract_text()
            # 利用正则或其他方法获取pdf表格开始页与结束页
            start_page = int(re.findall('3.4.*?(\d+)', content)[0])
            end_page = int(re.findall('3.5.*?(\d+)', content)[0])

start_page = start_page+4  # +4是因为实际页数比标明页数多4页
end_page = end_page+4
table_range = [m for m in range(start_page,end_page+1)]
# 这里的dataframe是根据自己实际情况创建，list2是表头，比如['a','b','c','d','e','f','g']
#df = pd.DataFrame(columns=list2)
df = pd.DataFrame(columns=list1)

# 再次遍历pdf的page
for page in pdf.pages:
    df_len = len(df)
    if page.page_number in table_range:
        # 有可能存在多个表格，起始页取最后一个，结束页取第一个
        table = page.extract_tables()[-1]
        # table[1],table[2]...就是要保存的数据
        for j in table[1:]:
            # 有7列数据，这里写死了。可根据实际情况变通
            for k in range(0, 7):
                # 替换第一个值得第一个换行符为空
                if (k == 0) and (j[k] is not None):
                    df.loc[df_len, df.columns[k]] = j[k].replace('\n', '', 1)
                else:
                    df.loc[df_len, df.columns[k]] = j[k]
            df_len = len(df)

pdf.close()
writer.save()
# 所有操作结束后还要关闭writer对象
writer.close()
