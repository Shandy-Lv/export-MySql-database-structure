import pymysql
import xlsxwriter
import xlrd


def write_data(worksheet, row, values):
    col = 0
    for v in values:
        worksheet.write(row, col, v)
        col += 1


database_info_sheet = xlrd.open_workbook("database_info.xlsx").sheets()[0]
rows = database_info_sheet.nrows
row = 0
dic = {"": ""}
while row < rows:
    dic[database_info_sheet.cell_value(row, 0)] = str(database_info_sheet.cell_value(row, 1))
    row += 1
workbook = xlsxwriter.Workbook('Result.xlsx')
worksheet = workbook.add_worksheet()
dbecs = pymysql.connect(host=dic["host"], user=dic["user"], passwd=dic["passwd"], database=dic["database"],
                        port=int(dic["port"]), charset=dic["charset"])
cursor1 = dbecs.cursor()
cursor1.execute("SHOW TABLES")
tables = cursor1.fetchall()
row = 0
write_data(worksheet, row, ["Field Name", "Datatype", "Collation", "Allow Null?", "PK?", "Auto Incr?", "Comment"])
for table in tables:
    print(table)
    # exec1 = "DESC" + " " + str(table)[2:-3]
    exec1 = "SHOW FULL FIELDS FROM" + " " + str(table)[2:-3]
    worksheet.write(row, 0, str(table)[2:-3])
    row += 1
    cursor1.execute(exec1)
    result = cursor1.fetchall()
    for r in result:
        print(r)
        values = [r[0], r[1], r[2], r[3], r[4], r[6], r[8]]
        write_data(worksheet, row, values)
        row += 1
workbook.close()
