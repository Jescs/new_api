from common.read_excel import ReadExcel


class ReadBaseData(ReadExcel):
    num = 0
    id = 0
    method = 1
    path = 2
    start_data = 3
    end_data = 11
    expect = 11
    rows = 7
    is_need = 12
    cell = 13

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
                    value = super().get_xlrd(row, i)
                    values.append(value)
                    values = [i for i in values if i is not None]
                else:
                    key = super().get_xlrd(row, i)
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
    print(excel.get_cell(4))
    # print(excel.get_path(6))
    # print(excel.get_all_data(4))
    # print(len(excel.get_all_datas()))
    # print(excel.get_all_datas())
