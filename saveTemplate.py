import sqlite3
from openpyxl import load_workbook

def create_connection(db_file):

    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print("connected")
    except sqlite3.Error as e:
        print(e)
    return conn
    

def create_table(conn):

    print("about to create table")
    create_table_sql = """CREATE TABLE IF NOT EXISTS testTemplate2 
    (id INTEGER PRIMARY KEY AUTOINCREMENT,
    workSheetData TEXT
    )"""

    try:
        cursor = conn.cursor()
        cursor.execute(create_table_sql)
        print("table created")
    except sqlite3.Error as e:
        print(e)
    
def insert(conn, data):
    InsertSql = """INSERT INTO testTemplate2(workSheetData)
    VALUES(?)"""

    cursor = conn.cursor()
    cursor.execute(InsertSql, data)
    conn.commit()
    print("data added")
    return cursor.lastrowid
    

def selectTemplates(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM testTemplate2")
    conn.commit()

    rows = cursor.fetchall()
    return rows[0]

def deleteTemplate(conn, id):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM testTemplate2 WHERE id=?", (id,))

conn = create_connection("test.db")

#with conn:
    #create_table(conn)

    #insert(conn, ['1,2,3', '4,5,6', '7,8,9'])

    #selectTemplates(conn)

    #deleteTemplate(conn, 1)




def saveTemplate(excelFile):
    wb = load_workbook("/Users/user-1/Desktop/TestDispatchGenerator.xlsx")
    dataList = []

    countSinceData = 0
    for column in wb["Data Sheet"].iter_cols():
        dataList.append([])
        for cell in column:
            if cell.value != None:
                countSinceData = 0
            elif cell.value == None:
                countSinceData += 1
            
            dataList[-1].append(cell.value)
        if countSinceData > 3:
            dataList[-1] = dataList[-1][:1-countSinceData]        
        dataList[-1].append(None)

    dataString = ""
    for column in dataList:
        dataString += "@"
        for cell in column:
            dataString += f"{cell}"
            dataString += "~"
        dataString = dataString[1:-1]
    


    with conn:
        #create_table(conn)
        #insert(conn, [dataString,])
        newDataList = selectTemplates(conn)
        newDataList = dataString.split("@")
        for i in range(len(newDataList)):
            newDataList[i] = newDataList[i].split("~")
        

        newTemplate = wb.create_sheet("Example")



        for col_idx, col_data in enumerate(newDataList, start=1):
            for row_idx, value in enumerate(col_data, start=1):
                if value == "None":
                    newTemplate.cell(row=row_idx, column=col_idx, value=None)
                else:
                    newTemplate.cell(row=row_idx, column=col_idx, value=value)
        wb.save("/Users/user-1/Desktop/TestDispatchGenerator.xlsx")


saveTemplate("hi")