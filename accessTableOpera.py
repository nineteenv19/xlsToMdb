import os.path

import pyodbc
from comtypes.client import GetModule, CreateObject
# from comtypes.gen import Access
import shutil


class OperaTable:
    def __init__(self):
        self.row_count = None
        self.src_table_name = None
        self.conn = None
        self.cursor = None

    # 创建数据库
    # def createDataBase(self, DataBaseName):  # , TableName, TableInfo):
    #     try:
    #         access = CreateObject('Access.Application')
    #         DBEngine = access.DBEngine
    #         db = DBEngine.CreateDatabase(DataBaseName, Access.DB_LANG_GENERAL)
    #         db.BeginTrans()
    #
    #         # db.Execute(TableName)
    #         # db.Execute(TableInfo)
    #         # db.Execute("CREATE TABLE test (ID Integer, numapples Integer)")
    #         # db.Execute("INSERT INTO test VALUES (1, ' ')")
    #
    #         db.CommitTrans()
    #         db.Close()
    #         access.Quit()
    #         return 1
    #     except Exception as e:
    #         print("数据库创建失败")
    #         return 0

    # 移动文件
    def myMoveFile(self, src_path_file, dst_path_file):
        if not os.path.isfile(src_path_file):
            print("该路径下文件不存在" + src_path_file)
            return 0
        else:
            file_path, file_name = os.path.split(dst_path_file)
            if not os.path.exists(file_path):
                os.makedirs(file_path)
            shutil.move(src_path_file, dst_path_file)
            print("目标文件路径：" + dst_path_file)
            return 1

    # 连接数据库
    def connectDataBase(self, path):
        try:
            str1 = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + path + ";Uid=;Pwd=;"
            self.conn = pyodbc.connect(str1)
            self.cursor = self.conn.cursor()
            return 1
        except Exception as e:
            print("数据库连接失败")
            return 0

    def queryTableName(self):
        try:
            self.src_table_name = [table.table_name for table in self.cursor.tables(tableType='Table')]
            # print(self.src_table_name)
            return 1
        except Exception as e:
            print("查询数据库所有表名失败")
            return 0

    # 创建表单
    def createTable(self, tableName, tableField):
        try:
            str1 = "create table " + tableName + "(" + tableField + ")"
            self.cursor.execute(str1)
            self.cursor.commit()
            return 1
        except Exception as e:
            print("创建" + tableName + "表单失败")
            return 0

    # 添加表单字段
    def addTableField(self, tableName, fieldName, fieldType):
        try:
            for i in range(len(fieldName)):
                str1 = "SELECT * FROM " + tableName
                self.cursor.execute(str1)
                columns = [column[0] for column in self.cursor.description]

                if fieldName[i] not in columns:
                    str1 = "alter table " + tableName + " add column " + fieldType[i]
                    self.cursor.execute(str1)
                    self.cursor.commit()

            return 1
        except Exception as e:
            print("添加" + tableName + "表单字段失败")
            return 0

    def queryTableRow(self, tableName):
        try:
            str1 = "SELECT COUNT(*) FROM " + tableName
            self.cursor.execute(str1)
            self.row_count = self.cursor.fetchone()[0]
            return 1
        except Exception as e:
            print("查询" + tableName + "表中行数失败")
            return 0

    # 插入表单数据
    def insertTableInfo(self, tableName, field, values):
        try:
            str1 = "Insert Into " + tableName + "(" + field + ")" + " Values (" + values + ")"
            self.cursor.execute(str1)
            self.cursor.commit()
            return 1
        except Exception as e:
            print("插入表单数据失败")
            return 0

    # 查询字段数据
    def selectTableInfo(self, tableName, field):
        try:
            str1 = "select " + field + " from " + tableName
            self.cursor.execute(str1)
            # row = self.cursor.fetchone()
            # print(row)
            # for row in self.cursor.execute(str1):
            #     print(row)
            return 1
        except Exception as e:
            print("查询字段数据失败")
            return 0

    # 更新字段数据
    def updateTableInfo(self, tableName, fields_values, value_index):
        try:
            str1 = "update " + tableName + " set " + fields_values + " where ID=" + value_index
            self.cursor.execute(str1)
            self.cursor.commit()
            return 1
        except Exception as e:
            print("更新字段数据失败")
            return 0

    def deleteTable(self, tableName):
        try:
            str1 = "DROP TABLE" + tableName
            self.cursor.execute(str1)
            self.conn.commit()
            return 1
        except Exception as e:
            print("清空表中数据失败")
            return 0

    def deleteAllInfoTable(self, tableName):
        try:
            str1 = "delete from " + tableName
            self.cursor.execute(str1)
            self.conn.commit()
            str2 = "ALTER TABLE " + tableName + " ALTER COLUMN ID COUNTER(1,1)"
            self.cursor.execute(str2)
            self.conn.commit()

            return 1
        except Exception as e:
            print("清空表中数据失败")
            return 0

    # 删除字段诗句
    def deleteTableInfo(self, tableName, field, value):
        try:
            str1 = "delete from " + tableName + " where " + field + "=?"
            self.cursor.execute(str1, value)
            self.cursor.commit()
            return 1
        except Exception as e:
            print("删除字段数据失败")
            return 0

    # 断开数据库连接
    def disconnectDataBase(self):
        try:
            self.cursor.close()
            self.conn.close()
            return 1
        except Exception as e:
            print("断开连接异常")
            return 0


if __name__ == '__main__':
    # str1 = r'C:\\Program Files\\Microsoft Office\\root\\Office16\\MSACC.OLB'
    # GetModule(str1) # 生成对应包装器

    test1 = OperaTable()
    # test1.createDataBase("test.mdb", "CREATE TABLE test (ID Integer, numapples Integer)", "INSERT INTO test VALUES (1, 1)")
    # test1.myMoveFile('C:/Users/nineteen/Documents/test.mdb', 'D:/DataBase')
    path = "D:/DataBase/test.mdb"
    tableName = "table_test1"
    fieldType = "ceshi5 text"
    field = ""
    value = "('4','6')"
    ret = test1.connectDataBase(path)
    if ret == 0:
        quit()

    ret = test1.queryTableName()
    if ret == 0:
        quit()

    # ret = test1.createTable(tableName)
    if ret == 0:
        quit()

    # ret = test1.addTableField(tableName, fieldType)
    if ret == 0:
        quit()

    field = ""
    value = "('1', '0')"
    # ret = test1.insertTableInfo(tableName, field, value)
    if ret == 0:
        quit()

    # field = "*"
    # ret = test1.selectTableInfo(tableName, field)
    # if ret == 0:
    #     quit()

    field = "ID"
    value = 'ttt'
    field2 = "Remark"
    value2 = '1'
    # ret = test1.updateTableInfo(tableName, field, field2, value, value2)
    if ret == 0:
        quit()

    ret = test1.deleteTableInfo(tableName, field, value)
    if ret == 0:
        quit()

    field = "*"
    ret = test1.selectTableInfo(tableName, field)
    if ret == 0:
        quit()

    ret = test1.disconnectDataBase()
    if ret == 0:
        quit()

    print("测试结束")
