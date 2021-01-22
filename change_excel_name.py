import os
import xlrd


if __name__ == '__main__':
    path = 'E://test/'
    count = 1
    for file in os.listdir(path):
        print('-------------------')
        print(file)
        file_path = path + file
        try:
            data = xlrd.open_workbook(file_path)
            sheet = data.sheet_by_name("Sheet1")
            name = sheet.cell_value(0, 0)
            if name == '附件1：':
                name = sheet.cell_value(1, 0)
            new_name = str(count) + str(name) + '.xls'
            print(new_name)
            old_file = path + '/' + file
            new_file = path + '/' + new_name
            os.rename(os.path.join(path, file), os.path.join(path, new_name))
            # os.rename(old_file, new_file)
            count = count + 1
        except Exception as e:
            print(e)