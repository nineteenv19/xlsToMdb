import sys
from accessTableOpera import OperaTable
from read_xlsx_to_mdb import Opera_xls
import os
from mainWindow import Ui_Form
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QIcon, QColor
from itertools import zip_longest
import json
import datetime
import time


class mainUI(QWidget, Ui_Form):
    def __init__(self):
        super(mainUI, self).__init__()
        self.setupUi(self)
        self.settings = QSettings("./FormConfig.ini", QSettings.IniFormat)
        self.settings.setIniCodec("UTF8")

        self.field_name = None
        self.table_name = None
        self.column_name = None
        self.var_name = None
        self.xls_model_name = None
        self.xls_path = None
        self.build_path = None
        self.xlsOpera = None
        self.i_checked_items = []

        self.database_model = QStandardItemModel()
        self.xls_model = QStandardItemModel()
        self.read_settings()

        self.field_dict = None
        self.table_field_dic = None
        self.table_field_list = None  # "[ID] AUTOINCREMENT, [Part Number] text, [Part Type] text, [Manufacturer Part Number] text, [Value] text, [Description] text, [Manufacturer] text, [Datasheet] text, [PCB Footprint] text, [Schematic Part] text, [更新人] text, [更新时间] DATETIME,[状态] text"
        # self.field_dict = {
        #     "Part Number": "零部件编码-end",
        #     "Part Type": "小类名称",
        #     "Manufacturer Part Number": "供应商物料编码",
        #     "Value": "值",
        #     "Description": "器件描述",
        #     "Manufacturer": "制造商",
        #     "Datasheet": "",
        #     "PCB Footprint": "封装",
        #     "Schematic Part": "",
        #     "更新人": "申请人",
        #     "更新时间": "最后更新时间",
        #     "状态": ""
        # }

        self.pushButton.clicked.connect(self.choose_src_path_file)
        self.pushButton_2.clicked.connect(self.choose_save_path_file)
        self.pushButton_3.clicked.connect(self.click_check)

        self.pushButton_xls_edit.clicked.connect(self.edit_xls_status)
        self.pushButton_database_edit.clicked.connect(self.edit_database_status)
        self.pushButton_database_add.clicked.connect(self.add_database_status)
        self.pushButton_database_delete.clicked.connect(self.delete_database_status)
        self.init_tree_view()

    def init_settings(self):
        self.settings.setValue("CONFIG/XLS_PATH", "./探维物料编码数据库.xlsx")
        self.settings.setValue("CONFIG/BUILD_PATH", "./")
        self.settings.setValue("CONFIG/XLS_MODEL_NAME", "大类名称")

        field_name = ["零部件编码-end", "小类名称", "供应商物料编码", "值", "器件描述", "制造商", "Datasheet", "封装",
                      "Schematic Part", "申请人", "最后更新时间"]
        column_name = ["Part Number", "Part Type", "Manufacturer Part Number", "Value", "Description", "Manufacturer",
                       "Datasheet", "PCB Footprint", "Schematic Part", "更新人", "更新时间"]
        var_name = ["text", "text", "text", "text", "text", "text", "text", "text", "text", "text",
                    "DATETIME"]

        fn = ','.join(field_name)
        cn = ','.join(column_name)
        vn = ','.join(var_name)
        self.settings.setValue("XLS_DATABASE/FIELD_NAME", fn)
        self.settings.setValue("DATABASE/COLUMN_NAME", cn)
        self.settings.setValue("DATABASE/VAR_NAME", vn)
        self.settings.setValue("DATABASE/TABLE_NAME", "TanwayDataBase.mdb")

    def read_settings(self):
        if not os.path.exists("./FormConfig.ini"):
            self.print_str("创建初始化配置文件: " + self.settings.fileName())
            self.init_settings()
        try:
            self.xls_path = self.settings.value("CONFIG/XLS_PATH")
            self.build_path = self.settings.value("CONFIG/BUILD_PATH")
            self.xls_model_name = self.settings.value("CONFIG/XLS_MODEL_NAME")
            self.lineEdit.setText(self.xls_path)
            self.lineEdit_2.setText(self.build_path)

            column_name = self.settings.value("DATABASE/COLUMN_NAME")
            field_name = self.settings.value("XLS_DATABASE/FIELD_NAME")
            var_name = self.settings.value("DATABASE/VAR_NAME")

            self.field_name = field_name.split(',')
            self.column_name = column_name.split(',')
            self.var_name = var_name.split(',')

            self.table_name = self.settings.value("DATABASE/TABLE_NAME")

        except Exception as e:
            self.print_str("读取参数不存在，重新生成配置文件")
            self.init_settings()
            self.read_settings()

    def write_settings(self):
        fn = ','.join(self.field_name)
        cn = ','.join(self.column_name)
        vn = ','.join(self.var_name)
        self.settings.setValue("XLS_DATABASE/FIELD_NAME", fn)
        self.settings.setValue("DATABASE/COLUMN_NAME", cn)
        self.settings.setValue("DATABASE/VAR_NAME", vn)
        self.settings.setValue("CONFIG/XLS_PATH", str(self.xls_path))
        self.settings.setValue("CONFIG/BUILD_PATH", str(self.build_path))

    def init_tree_view(self):
        if not os.path.exists(self.xls_path):
            str1 = "该路径下：" + self.xls_path + " 文件不存在"
            self.print_str(str1)
            return
        self.xlsOpera = Opera_xls()
        ret = self.xlsOpera.read_src_xls(self.xls_path, 0)
        if ret == 0:
            str1 = "读取文件：" + self.xls_path + " 失败"
            self.print_str(str1)
            return

        self.populate_xls_model(self.xls_model, self.field_name, self.xlsOpera.df)
        self.treeView_xls.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.treeView_xls.header().setSectionResizeMode(QHeaderView.Stretch)
        self.treeView_xls.setModel(self.xls_model)
        self.get_xls_selected_name(self.xls_model)

        self.populate_data_model(self.database_model, self.field_name, self.column_name, self.var_name)
        self.treeView_database.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.treeView_database.header().setSectionResizeMode(QHeaderView.Stretch)
        self.treeView_database.setModel(self.database_model)

    def closeEvent(self, event):
        self.write_settings()

    # treeview model
    def get_xls_selected_name(self, model):
        for row in range(model.rowCount()):
            index = model.index(row, 2)  # 获取第二列的索引
            model.setData(index, "")  # 将项目设置为空字符串

        count = 0
        checked_items = []
        for i in range(model.rowCount()):
            item = model.item(i, 1)
            if item.checkState() == 2:
                checked_items.append(item.text())
                value_item = QStandardItem(item.text())
                model.setItem(count, 2, value_item)
                count = count + 1
        self.i_checked_items = checked_items

    @staticmethod
    def populate_data_model(model, data1, data2, data3):
        model.clear()
        model.setHorizontalHeaderItem(0, QStandardItem('序号'))
        model.setHorizontalHeaderItem(1, QStandardItem('Xls Name'))
        model.setHorizontalHeaderItem(2, QStandardItem('DataBase Name'))
        model.setHorizontalHeaderItem(3, QStandardItem('Type'))
        count = 1
        for value1, value2, value3 in zip_longest(data1, data2, data3, fillvalue=""):
            value_item = QStandardItem(str(count))
            value_item1 = QStandardItem(str(value1))
            value_item2 = QStandardItem(str(value2))
            value_item3 = QStandardItem(str(value3))
            model.appendRow([value_item, value_item1, value_item2, value_item3])
            count = count + 1

    @staticmethod
    def populate_xls_model(model, field_data, all_data):
        # if isinstance(data, dict):
        data = set(all_data).intersection(field_data)
        model.clear()
        model.setHorizontalHeaderItem(0, QStandardItem('序号'))
        model.setHorizontalHeaderItem(1, QStandardItem('All Name'))
        model.setHorizontalHeaderItem(2, QStandardItem('Selected Name'))
        count = 1
        for value in all_data:
            value_item = QStandardItem(str(value))
            value_item.setCheckable(True)  # 设置父项为可勾选
            value_item1 = QStandardItem(str(count))
            if value in data:
                value_item.setCheckState(2)
            value_item.setFlags(value_item.flags() & ~Qt.ItemIsUserCheckable)  # 禁用勾选框
            model.appendRow([value_item1, value_item])
            count = count + 1

    @staticmethod
    def set_data_status(value_list, data):
        for value in value_list:
            if value in data:
                value_item = QStandardItem(str(value))
                value_item.setCheckState(2)

    def edit_xls_status(self):
        enable_interaction = not (self.xls_model.item(0, 1).flags() & Qt.ItemIsUserCheckable)
        for row in range(self.xls_model.rowCount()):
            item = self.xls_model.item(row, 1)
            if item is not None:
                if enable_interaction:
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsUserCheckable)

        if enable_interaction:
            self.pushButton_xls_edit.setText("保存当前状态")
        else:
            self.pushButton_xls_edit.setText("编辑")
            checked_items = []

            def save_item_state(item):
                if item.checkState() == 2:
                    checked_items.append(item.text())
                for row in range(item.rowCount()):
                    save_item_state(item.child(row, 1))

            save_item_state(self.xls_model.invisibleRootItem())
            self.i_checked_items = checked_items
            self.get_xls_selected_name(self.xls_model)
            self.print_str("xls字段重新勾选完成，请重新编辑数据库内容")

    def update_database_status(self):
        def read_column_data(model, idx):
            data = []
            for row in range(model.rowCount()):
                index = model.index(row, idx)
                item = model.itemFromIndex(index)
                data.append(item.text())
            return data

        self.field_name = read_column_data(self.database_model, 1)
        self.column_name = read_column_data(self.database_model, 2)
        self.var_name = read_column_data(self.database_model, 3)

        if "" in self.column_name:
            self.print_str("第三列存在空数据，请检查")
        if "" in self.field_name:
            self.print_str("第二列存在空数据，请检查")
        if "" in self.var_name:
            self.print_str("第四列存在空数据，请检查")

    def edit_database_status(self):
        if self.treeView_database.editTriggers() == QAbstractItemView.NoEditTriggers:
            self.treeView_database.setEditTriggers(QAbstractItemView.DoubleClicked)
            self.pushButton_database_edit.setText("保存")
        else:
            self.treeView_database.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.pushButton_database_edit.setText("编辑")
            self.update_database_status()
            self.print_str("数据列表修改成功")

    def add_database_status(self):
        count = self.database_model.rowCount() + 1
        item_len = len(self.i_checked_items)
        for i in range(item_len):
            value = str(self.i_checked_items[i])
            if value not in self.field_name:
                item1 = QStandardItem(str(count))
                item2 = QStandardItem(value)
                item3 = QStandardItem("")
                item4 = QStandardItem("text")
                self.database_model.appendRow([item1, item2, item3, item4])
                count = count + 1
        self.update_database_status()
        self.print_str("新增数据库数据段完成")

    def delete_database_status(self):
        last_index = self.database_model.rowCount() - 1
        self.database_model.removeRow(last_index)
        self.update_database_status()
        self.print_str("删除数据库数据段完成")

    def print_str(self, text):
        currentTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.plainTextEdit.appendPlainText(currentTime + " " + str(text))

    def choose_src_path_file(self):
        srcFile, ss = QFileDialog.getOpenFileName(self, "选择xls原始路径", "./", "Excel Files (*.xlsx)")
        if srcFile != "":
            # self.src_path_file = srcFile
            self.xls_path = srcFile
            self.lineEdit.setText(self.xls_path)
            self.init_tree_view()

    def choose_save_path_file(self):
        savePath = QFileDialog.getExistingDirectory(self, "选择保存路径", "")
        if savePath != "":
            self.build_path = savePath
            self.lineEdit_2.setText(self.build_path)

    def click_check(self):
        if self.xlsOpera is None:
            self.print_str("该路径下xls文件不存在，请重新选择xls文件地址")
            return 0

        ret = self.xlsOpera.classifyTableName(self.xls_model_name)
        if ret == 0:
            self.print_str("xls文件解析失败")
            return 0

        databaseOpera = OperaTable()
        database_path_file = self.build_path + '/' + self.table_name
        database_path = self.build_path
        if not os.path.exists(database_path):
            os.makedirs(database_path)

        if not os.path.isfile(database_path_file):
            # ret = databaseOpera.createDataBase(self.database_name)
            # database_src_path_file = self.database_src_path_file
            # ret = databaseOpera.myMoveFile(database_src_path_file, database_path_file)
            ret = 0
            if ret == 0:
                self.print_str("数据库原文件不存在")
                return 0

        len1 = len([s for s in self.field_name if s != ""])
        len2 = len([s for s in self.column_name if s != ""])
        len3 = len([s for s in self.var_name if s != ""])

        if len1 != len2 != len3:
            str1 = "创建表长度不同，请检查field_name长度：" + str(len1) + "column_name长度" + str(
                len2) + "var_name长度" + str(len3)
            self.print_str(str1)
            return 0

        ret = databaseOpera.connectDataBase(database_path_file)
        if ret == 0:
            self.print_str("数据库原连接失败")
            return 0

        ret = databaseOpera.queryTableName()
        if ret == 0:
            self.print_str("数据库表查询失败")
            return 0

        field_name_arr = self.column_name[:]
        field_name_arr.append('状态')
        field_dict = dict(zip_longest(field_name_arr, self.field_name, fillvalue=""))
        field_name_column_arr = [f'[{item}]' if isinstance(item, str) else item for item in field_name_arr]
        table_tmp = ', '.join(f'[{key}] {value}' for key, value in zip_longest(field_name_arr, self.var_name, fillvalue="text"))
        table_field = ('[ID] AUTOINCREMENT,' + table_tmp)
        field_type_arr = table_field.split(",")
        field_name_arr.insert(0, 'ID')

        for i in range(len(self.xlsOpera.table_name)):
            table_name = "[" + self.xlsOpera.table_name[i] + "]"
            if self.xlsOpera.table_name[i] in databaseOpera.src_table_name:
                ret = databaseOpera.deleteTable(table_name)
                if ret == 0:
                    self.print_str("数据库表删除失败")
                    return 0
            # else:
            ret = databaseOpera.createTable(table_name, table_field)
            if ret == 0:
                self.print_str("数据库表创建失败")
                return 0

            # ret = databaseOpera.addTableField(table_name, field_name_arr, field_type_arr)
            # if ret == 0:
            #     self.print_str("数据库表添加字段失败")
            #     return 0

            ret = databaseOpera.queryTableRow(table_name)
            if ret == 0:
                self.print_str("数据库表查询失败")
                return 0

            ret = self.xlsOpera.choiceXlsData(i, field_name_arr, field_dict)
            if ret == 0:
                self.print_str("数据库表解析失败")
                return 0

            # for row in range(databaseOpera.row_count):
            #     str1 = ""
            #     for col in range(len(field_name_column_arr)):
            #         str1 = str1 + field_name_column_arr[col] + "=" + xlsOpera.values[row][col]
            #         if col != len(field_name_column_arr) - 1:
            #             str1 = str1 + ", "
            #
            #     ret = databaseOpera.updateTableInfo(table_name, str1, str(row + 1))
            #     if ret == 0:
            #         return 0
            # ret = databaseOpera.deleteAllInfoTable(table_name)
            # if ret == 0:
            #     self.print_str("数据库表删除失败")
            #     return 0

            for count in range(len(self.xlsOpera.values)):
                str1 = ""
                str2 = ""
                for col in range(len(field_name_column_arr)):
                    str1 = str1 + field_name_column_arr[col]
                    str2 = str2 + self.xlsOpera.values[count][col]
                    if col != len(field_name_column_arr) - 1:
                        str1 = str1 + ","
                        str2 = str2 + ","

                ret = databaseOpera.insertTableInfo(table_name, str1, str2)
                if ret == 0:
                    self.print_str("数据库插入表单失败")
                    return 0

        ret = databaseOpera.disconnectDataBase()
        if ret == 0:
            self.print_str("数据库表断开连接失败")
            return 0

        self.print_str("***数据库表生成成功***")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_ui = mainUI()
    main_ui.setWindowIcon(QIcon("1.ico"))
    main_ui.setWindowTitle("探维物料编码数据库转换脚本V2.0")
    main_ui.show()
    app.exec_()
