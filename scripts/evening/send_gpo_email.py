import win32com.client
import pandas as pd

fileGPO = r'P:\Bloomberg\Management Fee for PVD\Management Fee for PVD_GPO-FIXED - LHFUND_REVISED_RATE.xls'
fileGPO2 = r'P:\Bloomberg\Management Fee for PVD\Management Fee for PVD_GPO-EQ - LHFUND  REVISED_RATE.xls'

df1 = pd.read_excel(fileGPO,sheet_name = 'Benchmark - PI',header=4 )
df2 = pd.read_excel(fileGPO2,sheet_name = 'Benchmark - PI',header=4)


table_str1 = df1.iloc[:,[0,1,2,4,5]].tail(1).to_html(header=True, index=False)
table_str2 = df2.iloc[:, [0, 2]].tail(1).to_html(header=True, index=False)

#table_str1 = df1.iloc[:, :7].tail(3).to_html(header=True, index=False)
#table_str2 = df2.iloc[:, [0, 2, 4]].tail(3).to_html(header=True, index=False)

ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'อัพเดท GPO'
newmail.To = 'risk@lhfund.co.th ; operation@lhfund.co.th'
#newmail.To = 'Panisarap@lhfund.co.th ; kornwipad@lhfund.co.th ; Amornsiris@lhfund.co.th'
#newmail.To = 'Amornsiris@lhfund.co.th'
############*********************************
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

# แนบไฟล์ (ถ้ามี)
# attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)

# ส่งอีเมล
newmail.Send()
