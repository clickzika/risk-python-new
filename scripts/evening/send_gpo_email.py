import sys
import os
import win32com.client
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from config import GPO_FIXED_FILE, GPO_EQ_FILE, EMAIL_RECIPIENTS

df1 = pd.read_excel(GPO_FIXED_FILE, sheet_name='Benchmark - PI', header=4)
df2 = pd.read_excel(GPO_EQ_FILE, sheet_name='Benchmark - PI', header=4)

table_str1 = df1.iloc[:, [0, 1, 2, 4, 5]].tail(1).to_html(header=True, index=False)
table_str2 = df2.iloc[:, [0, 2]].tail(1).to_html(header=True, index=False)

ol = win32com.client.Dispatch('Outlook.Application')
newmail = ol.CreateItem(0x0)
newmail.Subject = 'อัพเดท GPO'
newmail.To = EMAIL_RECIPIENTS
newmail.HTMLBody = f'''
<html>
<head>
<style>
    body {{ font-size: 10px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    table, th, td {{ border: 1px solid black; }}
    th, td {{ padding: 8px; text-align: left; }}
    tr:nth-child(even) {{ background-color: #f2f2f2; }}
    tr:hover {{ background-color: #ddd; }}
    th {{ background-color: #4CAF50; color: white; }}
</style>
</head>
<body>
<p style="font-size:18px; color: navy;">GPO เรียบร้อยครับ</p>
<p></p>
<p style="font-size:13px;">GPO-FIXED </p>
<p style="font-size:10px;">{table_str1}</p>
<p></p>
<p style="font-size:13px;">GPO-EQ</p>
<p style="font-size:10px;">{table_str2}</p>
</body>
</html>
'''
newmail.Send()
