import class_ali
import class_qcc
from openpyxl import Workbook

if __name__ == '__main__':
    wb = Workbook()
    ali_excle = wb.active
    ali_excle['A1'] = 'name'
    wb2 = Workbook()
    qcc_excle = wb2.active
    qcc_excle['A1'] = 'name'
    ali_error_list = []
    qcc_error_list = []
    name_list = ['linyi paper', 'linyi car', 'linyi pen']
    for name in name_list:
        print('当前搜素名字' + name)
        ali = class_ali.Ali(name)
        try:
            ali.deep_crawl_main()
            print(ali.search_name + '阿里爬取完成保存文件名: ' + ali.save_name)
        except Exception as e:
            ali.driver.quit()
            print(ali.search_name + '阿里爬取失败')
            print('阿里错误信息:' + str(e))
            ali_error_list.append(name)
            continue
        qcc = class_qcc.Qcc(name)
        try:
            qcc.star()
            print(qcc.search_name + '企查查爬取完成保存文件名: ' + qcc.save_name)
        except Exception as e:
            print(qcc.search_name + '企查查爬取失败')
            print('企查查错误信息:' + str(e))
            qcc_error_list.append(name)
    print('以下检索信息没有爬取到\n阿里:')
    print(ali_error_list)
    print('企查查:')
    print(qcc_error_list)
    count = 2
    for i in ali_error_list:
        ali_excle['A' + str(count)] = i
        count = count + 1
    count = 2
    for i in qcc_error_list:
        qcc_excle['A' + str(count)] = i
        count = count + 1
    wb.save('ali_error.xls')
    wb2.save('qcc_error.xls')