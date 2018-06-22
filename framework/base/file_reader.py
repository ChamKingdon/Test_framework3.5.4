# coding=utf-8
"""
文件读取。YamlReader读取yaml文件，ExcelReader读取excel。
"""
import yaml
import os
from xlrd import open_workbook
from datetime import datetime
from xlrd import xldate_as_tuple


class YamlReader:
    def __init__(self, yamlf):
        if os.path.exists(yamlf):
            self.yamlf = yamlf
        else:
            raise FileNotFoundError("the file is not exit")
        self._data = None

    @property
    def data(self):
        # 如果是第一次调用data，读取yaml文档，否则直接返回之前保存的数据
        if not self._data:
            with open(self.yamlf, 'rb') as f:
                # load后是个generator，用list组织成列表
                self._data = list(yaml.safe_load_all(f))
        return self._data


class SheetTypeError(Exception):
    pass


class ExcelReader:
    """
    读取excel文件中的内容。返回list。
    1、打开Excel文件读取数据
        data = xlrd.open_workbook('excel.xls')
    2、获取一个工作表
        table = data.sheets()[0]          #通过索引顺序获取
        table = data.sheet_by_index(0) #通过索引顺序获取
        table = data.sheet_by_name(u'Sheet1')#通过名称获取
    3、获取整行和整列的值（返回数组）
        table.row_values(i)
        table.col_values(i)
    4、获取行数和列数　
        table.nrows
        table.ncols
    5、获取单元格
　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     　　     table.cell(0,0).value
        table.cell(2,3).value
    6、表格的数据类型
        table.cell(i,j).ctype :  0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
    如：
    excel中内容为：
    | A  | B  | C  |
    | A1 | B1 | C1 |
    | A2 | B2 | C2 |

    如果 print(ExcelReader(excel, title_line=True).data)，输出结果：
    [{A: A1, B: B1, C:C1}, {A:A2, B:B2, C:C2}]

    如果 print(ExcelReader(excel, title_line=False).data)，输出结果：
    [[A,B,C], [A1,B1,C1], [A2,B2,C2]]


    可以指定sheet，通过index或者name：
    ExcelReader(excel, sheet=2)
    ExcelReader(excel, sheet='BaiDuTest')
    """

    def __init__(self, excel, sheet=0, title_line=True):
        if os.path.exists(excel):
            self.excel = excel
        else:
            raise FileNotFoundError('文件不存在！')
        self.sheet = sheet
        self.title_line = title_line
        self._data = list()

    # 按行
    @property
    def row_data(self):
        if not self._data:
            workbook = open_workbook(self.excel)  # 打开Excel文件读取数据
            if type(self.sheet) not in [int, str]:
                raise SheetTypeError(
                    'Please pass in <type int> or <type str>, not {0}'.format(
                        type(
                            self.sheet)))   # 判断索引是否输入正确
            elif isinstance(self.sheet, int):
                s = workbook.sheet_by_index(self.sheet)  # 通过索引获取
            else:
                s = workbook.sheet_by_name(self.sheet)  # 通过名称获取

            if self.title_line:

                # for row in range(0, s.nrows):
				#
                #     if s.cell(row, 0).ctype in (0, 2, 3, 4, 5):
				#
                #         if row == 0:
                #             title = []
				#
                #         if s.cell(
                #             row,
                #                 0).ctype == 0:  # ctype :  0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
                #             title.append("null")
                #         elif s.cell(row, 0).ctype == 2 and s.cell(row, 0).value % 1 == 0:
                #             title.append(int(s.cell(0, row).value))
                #         elif s.cell(row, 0).ctype == 3:
                #             # 转成datetime对象
                #             date = datetime(
                #                 *xldate_as_tuple(s.cell(row, 0), 0))
                #             title.append(date.strftime('%Y/%d/%m %H:%M:%S'))
                #         elif s.cell(row, 0).ctype == 4:
                #             title = True if title == 1 else False
                #         elif s.cell(row, 0).ctype == 5:
                #             title.append("The data type of the table is error")
                #     else:
				#
                #         title.append(s.cell(row, 0).value)  # 首行为title
                #     print(title)

                for col in range(1, s.ncols):   # s.ncols 获取列数
                    # 依次遍历其余行，与首行组成dict，拼到self._data中
                    # s.col_values 获取整列数


                    for row in range(0, s.nrows):
                        table = [[1 for j in range(1, row)] for i in range(1, col)]
                        print(table)

                        if s.cell(row, col).ctype in (0, 2,  3, 4, 5):

                            if row == 0 and col == 1:
                                pass
                            if s.cell(
                                row,
                                    col).ctype == 0:  # ctype :  0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
                                table[row][col].append("null")

                            elif s.cell(row, col).ctype == 2 and s.cell(row, col).value % 1 == 0:
                                print(str(row)+"==========="+str(col))
                                table[row][col] = int(s.cell(row, col).value)

                            elif s.cell(row, col).ctype == 3:
                                # 转成datetime对象
                                date = datetime(
                                    *xldate_as_tuple(s.cell(row, 0), 0))
                                table[row][col].append(
                                    date.strftime('%Y/%d/%m %H:%M:%S'))

                            elif s.cell(row, col).ctype == 4:
                                table[row][col] = True if table[col][row] == 1 else False

                            elif s.cell(row, col).ctype == 5:
                                table[row][col].append(
                                    "The data type of the table is error")
                        else:

                            table[row][col] = s.cell(row, col).value     # 首行为title
                        print(table)

                    #self._data.append(dict(zip((title), s.col_values(row))))

            else:
                print("two:" + str(self.title_line))
                for row in range(0, s.nrows):
                    # 遍历所有行，拼到self._data中
                    self._data.append(
                        s.row_values(row))    # s.row_values 获取整行数
        return self._data

    # 按列
    @property
    def col_data(self):
        if not self._data:
            workbook = open_workbook(self.excel)  # 打开Excel文件读取数据
            if type(self.sheet) not in [int, str]:
                raise SheetTypeError(
                    'Please pass in <type int> or <type str>, not {0}'.format(
                        type(
                            self.sheet)))   # 判断索引是否输入正确
            elif isinstance(self.sheet, int):
                s = workbook.sheet_by_index(self.sheet)  # 通过索引获取
            else:
                s = workbook.sheet_by_name(self.sheet)  # 通过名称获取

            if self.title_line:
                print("one:" + str(self.title_line))
                title = s.row_values(0)  # 首行为title
                for col in range(1, s.nrows):   # s.ncols 获取列数
                    # 依次遍历其余行，与首行组成dict，拼到self._data中
                    # s.row_values 获取整行数
                    self._data.append(dict(zip(title, s.row_values(col))))

            else:
                print("two:" + str(self.title_line))
                for col in range(0, s.nrows):
                    # 遍历所有行，拼到self._data中
                    self._data.append(
                        s.col_values(col))    # s.col_values 获取整列数
        return self._data


if __name__ == '__main__':
    # y = 'E:\Python\Project\\venv3.5.4\\framework\config\config.yml'
    # reader = YamlReader(y)
    # print(reader.data)

    e = 'E:/Python/Project/venv3.5.4/framework/data/test.xlsx'

    readers = ExcelReader(e, title_line=True)
    print(readers.row_data)
