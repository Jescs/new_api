from common.read_excel import ReadExcel


class ReadBaseData(ReadExcel):
    num = 0
    id = 0
    method = 0
    path = 1
    start_data = 3
    end_data = 11 # load_workbook 统计行数跟列数是从1开始，使用get_xlrd时，行数与列数都需要在row，col的基础上进行加一
    expect = 10
    rows = 7
    is_need = 11  # open_workbook 统计行数与列数是从0开始
    cell = 12

    def __init__(self):
        super().__init__(self.num)

    def get_id(self, row):
        return super().get_cell_value(row, self.id)

    def get_method(self, row):
        return super().get_cell_value(row, self.method)

    def get_path(self, row):
        return super().get_cell_value(row, self.path)

    def get_expect(self, row):
        return super().get_cell_value(row, self.expect)

    def get_is_need_response(self, row):
        return super().get_cell_value(row, self.is_need)

    def get_cell(self, row):
        return super().get_cell_value(row, self.cell)

    def get_data(self, row):
        for j in range(row):
            keys = []
            values = []
            for i in range(self.start_data, self.end_data):
                if i % 2 == 0:
                    value = super().get_xlrd(row+1, i)
                    values.append(value)
                    values = [i for i in values if i is not None]
                else:
                    key = super().get_xlrd(row+1, i)
                    keys.append(key)
                    keys = [i for i in keys if i is not None]
            datas = dict(zip(keys, values))
            return datas

    def get_all_data(self, row):
        data_list = [self.get_method(row), self.get_path(row),
                     self.get_data(row), self.get_expect(row), self.get_is_need_response(row), self.get_cell(row)]
        return data_list

    def get_all_datas(self):
        data_lists = []
        for i in range(self.rows):
            data_lists.append(self.get_all_data(i))
        return data_lists


if __name__ == '__main__':
    excel = ReadBaseData()
    print(excel.get_data(3))
