import mysql.connector
from openpyxl import Workbook

# Database
db = mysql.connector.connect(
    host ='localhost',
    port=3306,
    user='root',
    password='Admin@123',
    database='Plug_want_to_buy'
)

curses = db.cursor()
sql ='''
    select *
    from products;
'''
curses.execute(sql)
products = curses.fetchall()

for p in products:
    print(p)

curses.close()
db.close()

# Excel
workbook = Workbook()
sheet = workbook.active
sheet.append(['ID', 'ชื่อสินค้า', 'ราคา', 'ความต้องการ', 'วันที่บันทึก'])

for p in products:
    print(p)
    sheet.append(p)

workbook.save(filename='exported.xlsx')