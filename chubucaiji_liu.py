import json
import os

from main import liu_chubucaiji
import pandas as pd
import logging
import xlsxwriter


def getLog():
    dirname = r'log'
    filename = dirname + '/chubucaiji_6_log.log'
    if not os.path.exists(dirname):
        os.makedirs(dirname)
    logging.basicConfig(
        filename=filename,
        level=logging.DEBUG,
        format='%(asctime)s:%(levelname)s:%(message)s'
    )
    return logging


def show_conmpaies_data1():
    filename = 'return_data.json'
    with open(filename) as file_obj:
        names = json.load(file_obj)
    # print(names)
    return names


def start():
    qiyedata = show_conmpaies_data1()
    i = 0
    for qiye in qiyedata:
        i += 1
        scompanyname = qiye['sCompanyName']
        scompanyid = qiye['sCompanyID']
        num = len(qiyedata)
        test(scompanyid, scompanyname, i, num)


def test(scompanyid, scompanyname, i, num):
    ipagecount = 100
    ipageindex = 1
    # scompanyid = '2cbc869a-4ff8-4f78-bcd0-a6fa222b08f6'
    # scompanyname = '测试公司'
    data = liu_chubucaiji(scompanyid, ipageindex, ipagecount)
    all_count = data['Table1'][0]['iCount']
    all_ipagecount = round(all_count / ipagecount)
    logging.info('{}\tEpoch {}/{}\t{}/{}页'.format(scompanyname, i, num, ipageindex, all_ipagecount))
    print('{}\tEpoch {}/{}\t{}/{}页'.format(scompanyname, i, num,  ipageindex, all_ipagecount))
    # print(data)
    data_table = data['Table']
    if ipageindex < all_ipagecount:
        for index in range(1, all_ipagecount):
            print('{}\tEpoch {}/{}\t{}/{}页'.format(scompanyname, i, num, (index + 1), all_ipagecount))
            logging.info('{}\tEpoch {}/{}\t{}/{}页'.format(scompanyname, i, num, (index + 1), all_ipagecount))
            tm_data = liu_chubucaiji(scompanyid, (index + 1), ipagecount)
            for x in tm_data['Table']:
                data_table.append(x)
    df = pd.DataFrame()
    print('%s--数据长度%d' % (scompanyname, len(data_table)))
    logging.info('%s--数据长度%d' % (scompanyname, len(data_table)))
    for v in data_table:
        df = df.append(
            [
                {
                    'sCompanyID': scompanyid,
                    'sCompanyName': scompanyname,
                    'sCity': v['sCity'],
                    'sInstalmentID': v['sInstalmentID'],
                    'sSchemeName': v['sSchemeName'],
                    'fPurposeArea': v['fPurposeArea'],
                    'fPriceAverage': v['fPriceAverage'],
                    'fFloorPrice': v['fFloorPrice'],
                    'sBuildcyc': v['sBuildcyc'],
                    'sOpenDate': v['sOpenDate']
                }
            ]
        )
    df.to_excel(excel_writer=('excel/六/初步采集/' + scompanyname + '.xlsx'), sheet_name='Sheet1', index=None)


if __name__ == '__main__':
    logging = getLog()
    start()
    # test()
