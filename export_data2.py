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
    select p.id as id, p.title as title, p.price as price, c.title as category
    from products as p
    left join categories as c
    on p.category_id = c.id
'''
curses.execute(sql)
products = curses.fetchall()


# Excel
workbook = Workbook()
sheet = workbook.active
sheet.append(['ID', 'ชื่อสินค้า', 'ราคาสินค้า', 'ประเภทสินค้า'])

for p in products:
    print(p)
    sheet.append(p)

workbook.save(filename='exported2.xlsx')

curses.close()
db.close()