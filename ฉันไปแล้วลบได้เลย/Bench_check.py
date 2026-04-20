###### ของแทร่
from datetime import datetime, timedelta
import pandas as pd
import win32com.client

today = (datetime.now() - timedelta(days=0)).strftime('%Y-%m-%d')

file_benchmark = r'\\w2fsspho101.lhfund.net\FM-RI$\risk\1.Risk Report\1.Daily Risk Report\Benchmark.xlsm'


# อ่านข้อมูลจากไฟล์ Excel
df1 = pd.read_excel(file_benchmark, sheet_name='Bench', header=4)

# ตรวจสอบว่าคอลัมน์ที่ 3 ถึงสุดท้ายมีค่า NaN เหมือนกันในแถวเดียวกันหรือไม่
mask = df1.iloc[:, 2:].isna().all(axis=1)

# ลบบรรทัดที่มีค่า NaN ในคอลัมน์ที่ 3 ถึงสุดท้าย
df1_cleaned = df1[~mask]

#print("DataFrame หลังการลบ:")
#print(df1_cleaned)


df1_cleaned = df1_cleaned[df1_cleaned["Ticker"].dt.weekday < 5]

# เลือกข้อมูล 5 แถวสุดท้าย
df2 = df1_cleaned.tail(3)

print(df2)


path = fr'\\w2fsspho101.lhfund.net\FM-RI$\risk\Amornsiri\Benchmark_duplicate\{today}.xlsx'

df2.to_excel(path)


#------------------------------------------------------------------------------

# ตรวจสอบเฉพาะคอลัมน์ที่ 3 ถึงสุดท้าย
columns_to_check = df2.columns[2:]
duplicated_columns = []

value_to_zero_columns = []


# บรรทัดล่างสุดและบรรทัดก่อนหน้า
last_row = df2.iloc[-1]  # บรรทัดล่างสุด
previous_row = df2.iloc[-2]  # บรรทัดก่อนหน้า

# ตรวจสอบค่าระหว่างบรรทัด
for col in columns_to_check:
    # ตรวจสอบเงื่อนไข: previous_row มีตัวเลข และ last_row ไม่มีค่า หรือมีค่าเป็น 0
    if pd.notna(previous_row[col]) and previous_row[col] != 0:
        if pd.isna(last_row[col]) or last_row[col] == 0:
            # เก็บคอลัมน์ที่ผ่านเงื่อนไขลงใน table2
            value_to_zero_columns.append(col)

# สร้าง DataFrame สำหรับ table2
if value_to_zero_columns:
    table2_df = pd.DataFrame({"ตัวที่อยู่ๆก็เกิด 0 ": value_to_zero_columns})
    print(table2_df)
else:
    print("ไม่พบคอลัมน์ที่ผ่านเงื่อนไข")





# ตรวจสอบว่ามีค่าซ้ำในแต่ละคอลัมน์
for col in columns_to_check:
    # นับจำนวนแต่ละค่าในคอลัมน์ (ไม่รวม NaN)
    value_counts = df2[col].value_counts(dropna=True)
    
    # ถ้าค่าซ้ำ >= 5 ครั้งให้เพิ่มชื่อคอลัมน์ใน duplicated_columns
    if (value_counts >= 3).any():
        duplicated_columns.append(col)

# รายการคอลัมน์ที่ต้องยกเว้น
exclude_cols = ['SCBT3MD', 'BBLT3MD', 'TFBT3MD', 'TAVG3MD3', 'SCBT6MD', 'BBLT6MD', 'TFBT6MD',
                'KTBT6MD', 'TAVG6MD4', 'SCBT1YD', 'BBLT1YD', 'TFBT1YD', 'KTBT1YD', 'TAVG1YD3',
                'TAVG1YD4', 'SCBT1YD(FUND)', 'BBLT1YD(FUND)', 'TFBT1YD(FUND)', 'KTBT1YD(FUND)',
                'TAVG1YD4(FUND)', 'US0003M', 'US0006M', 'US0012M' ,'TBCITOTR', 'TBC1TOTR', 
                'TBC2TOTR', 'TBPBTOTR', 'TBP1TOTR', 'TBP2TOTR', 'TBBNTOTR', 'SETTHSIT', 
                'ACEMKEI LX', 'BGWHUBA ID']

# ลบคอลัมน์ที่อยู่ใน exclude_cols จาก DataFrame
columns_to_remove = [col for col in duplicated_columns if col in exclude_cols]
df_filtered = df1_cleaned.drop(columns=columns_to_remove, errors='ignore')

# ฟ้องว่ามีค่าซ้ำในคอลัมน์ที่เหลือหลังจากลบ
remaining_duplicated = [col for col in duplicated_columns if col not in exclude_cols]

if remaining_duplicated:
    print(f"พบค่าซ้ำ 2 ครั้งในคอลัมน์ที่ต้องตรวจสอบ: {', '.join(remaining_duplicated)}")
else:
    print("ไม่พบค่าซ้ำ 2 ครั้งในคอลัมน์ที่ตรวจสอบ")

remaining_duplicated_df = pd.DataFrame(list(remaining_duplicated), columns=["Duplicated Values"])    
table_str1 = remaining_duplicated_df.to_html(header=True, index=False)
table2_df  = table2_df.to_html(header=True, index=False)

ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'Benchmark วันนี้มีอะไรซ้ำ'
#newmail.To = 'risk@lhfund.co.th ; operation@lhfund.co.th'
newmail.To = 'Panisarap@lhfund.co.th ; kornwipad@lhfund.co.th ; Amornsiris@lhfund.co.th'
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
<p></p>

<p style="font-size:10px;">{table_str1}</p>
<br><br>
<tr></tr>
<tr></tr>
<p style="font-size:10px;">{table2_df}</p>
<p></p>
<p style="font-size:18px; color: navy;">สามารถดูเพิ่มเติมได้ที่ :</p>
<p style="font-size:18px; color: navy;">{path}</p>
</body>
</html>
'''

# แนบไฟล์ (ถ้ามี)
# attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)

# ส่งอีเมล
newmail.Send()

