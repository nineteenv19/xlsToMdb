import pandas as pd
import numpy as np


class Opera_xls:
    def __init__(self):
        self.table_name = None
        self.df = None
        self.table_info = []
        self.values = []

    def read_src_xls(self, src_path_file, index):
        try:
            self.df = pd.read_excel(src_path_file, sheet_name=index)
            return 1
        except Exception as e:
            print("读取xlxs文件失败")
            return 0

    def classifyTableName(self, str1):
        try:
            self.table_name = self.df[str1].unique().astype(str)
            for i in range(len(self.table_name)):
                if self.table_name[i] == 'nan':
                    self.table_name[i] = '其他项'
                    self.table_info.append(self.df.loc[self.df[str1].isnull()].values)
                else:
                    self.table_info.append(self.df.loc[self.df[str1] == self.table_name[i]].values)
            return 1
        except Exception as e:
            print("解析xls列名失败")
            return 0

    def choiceXlsData(self, table_index, field_name, field_dict):
        try:
            self.values.clear()
            for i in range(len(self.table_info[table_index])):
                str1 = []
                for key, value_dx in field_dict.items():
                    column_name = value_dx
                    value = ""
                    if column_name != "":
                        column_index = self.df.columns.get_loc(column_name)
                        value = str(self.table_info[table_index][i][column_index])
                        if value == "nan":
                            value = "NULL"
                        else:
                            value = "'" + value + "'"
                    else:
                        value = "NULL"

                    str1.append(value)
                self.values.append(str1)
            return 1
        except Exception as e:
            print("解析xls数据失败")
            return 0


if __name__ == '__main__':
    src_path_file = 'D:/PycharmProjects/FeiShuForm/探维物料编码数据库.xlsx'
    opera = Opera_xls()
    ret = opera.read_src_xls(src_path_file, 0)
    model_name = '大类名称'
    ret = opera.classifyTableName(model_name)
    print(opera.table_name)
    print("测试结束")
