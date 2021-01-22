#-*- coding = utf-8 -*-
#@Author : PeterLong
import os,sys,re,csv,datetime,time,glob,shutil
from itertools import islice
import xlrd
root_folder = 'D:\\newbee'
file_folder = root_folder + '\\new'
excel_message = []
product_regext_list = []
country_regext_list = []
country_regext_ga_list=[]
csvfile = open(os.path.join(file_folder, "productregexlist.csv"),encoding='utf-8')
csv_reader = csv.reader(csvfile, dialect='excel')
for line in islice(csv_reader, 1, None):
    product_regext_list.append(line)
    # print(line)
csvfile.close()
csvfile = open(os.path.join(file_folder, "countryregexlist.csv"),encoding='utf-8')
csv_reader = csv.reader(csvfile, dialect='excel')
for line in islice(csv_reader, 1, None):
    country_regext_list.append(line)
    # print(line)
csvfile.close()

csvfile = open(os.path.join(file_folder, "country for ga.csv"),encoding='utf-8')
csv_reader = csv.reader(csvfile, dialect='excel')
for line in islice(csv_reader, 1, None):
    country_regext_ga_list.append(line)
    # print(line)
csvfile.close()

def read_csv():
    global file_folder, excel_message, product_regext_list, country_regext_list,country_regext_ga_list
    csv_list=glob.glob('*.csv')
    print(csv_list)
    for j in csv_list:
        csv_right=re.findall('Analytic.*',j,re.I)
        csv_right = str(csv_right).replace('[','').replace(']','').replace('\'','')
        print(csv_right)

        # for i in csv_right:

        if csv_right != '':
            csvfile = os.path.join(file_folder,(csv_right))
            csv_file = open(csvfile, "r", encoding="utf-8")
            csv_reader = csv.reader(csv_file, dialect='excel')
            csvfile = os.path.join(file_folder, (str(csv_right)).replace('[','').replace(']','').replace('\'',''))
            csv_file = open(csvfile, "r", encoding="utf-8", errors='ignore')
            numline = len(csv_file.readlines())
            csv_file = open(csvfile, "r", encoding="utf-8")
            csv_reader = csv.reader(csv_file, dialect='excel')
            for line in islice(csv_reader, 7, numline):
                # if len(line[0]) == 0:
                #     break
                try:
                    tmp_country = 1
                    for li in country_regext_ga_list:
                        if re.search(r'' + li[1] + r'', csv_right.split('_')[0], re.I):
                            tmp_country = li[0]
                            print(tmp_country)
                            break
                except:
                    tmp_country = 'no'
                tmp_datasource = 'GA'
                tmp_source=str(line[3]).strip().lower()
                tmp_Medium=str(line[2].split('/')[-1].strip().lower())
                if tmp_Medium.startswith('affiliat'):
                    tmp_Medium='affiliate'
                if tmp_Medium.startswith('face'):
                    tmp_Medium = 'fbig'
                if tmp_Medium.startswith('cpc'):
                    tmp_Medium = 'google'
                if tmp_Medium.startswith('paid'):
                    tmp_Medium = 'google'
                tmp_Campaign=line[1].strip().lower()
                tmp_user=line[4].strip().lower().replace(',','')
                tmp_new_user=line[5].strip().lower().replace(',','')
                tmp_Sessions=line[6].strip().lower().replace(',','')
                tmp_Bounce_Rate = (float(line[7].replace('%',''))/100)
                tmp_Pages_Session = line[8]
                tmp_Avg_Session_Duration = str(line[9].strip().lower())
                tmp_Ecommerce_Conversion_Rate = (float(line[10].replace('%',''))/100)
                tmp_Transactions=line[11]
                Revenue=line[12]
                t_list = list(Revenue)
                for i in range(len(t_list)):
                    try:
                        int(t_list[i])
                        break
                    except:
                        del t_list[i]
                tmp_Revenue = ''.join(t_list)
                tmp_Revenue = tmp_Revenue.replace(',', '').replace('?', '')
                if tmp_country == '':  # 英镑
                    tmp_Revenue = float(tmp_Revenue) * 1.3 / 1.18
                # tmp_Revenue=re.compile(r"\d+\.?\d*",Revenue)

                print(tmp_Revenue)
                date=time.strptime(line[0],"%Y%m%d")
                tmp_date=time.strftime("%Y-%m-%d",date).strip()
                print(tmp_date)



                tmp_line = [tmp_date,tmp_country,tmp_datasource, tmp_source, tmp_Medium, tmp_Campaign, tmp_user, tmp_new_user,tmp_Sessions,tmp_Bounce_Rate,tmp_Pages_Session,tmp_Avg_Session_Duration,tmp_Ecommerce_Conversion_Rate,tmp_Transactions,tmp_Revenue]
                excel_message.append(tmp_line)
                print(tmp_line)
            else:
                pass




if __name__ == '__main__':
    read_csv()

import cx_Oracle
os.environ['nls_lang'] = 'AMERICAN_AMERICA.AL32UTF8'
oracle_db = cx_Oracle.connect('dashboard_read', 'hihonordb20', '10.116.146.42:8080/HONOROve4755')
print(time.strftime('%H:%M:%S', time.localtime(time.time())),":已成功连接Oracle数据库，准备上传，数据库版本为：",oracle_db.version)
oracle_cur = oracle_db.cursor()

db_table = 'GoogleA_RAW_1'

oracle_sql = 'insert into '+ db_table +' values (:1, :2, :3, :4, :5, :6, :7,:8,:9,:10,:11,:12,:13,:14,:15)'

error_list = []
value_list = []
i = 0


def list2db(isend):
    global value_list
    if i % 500 == 0 or isend > 0:
        # 汇总上传法 - Oracle
        try:
            oracle_cur.executemany(oracle_sql, value_list)
            oracle_db.commit()
        except:
            for row_li in value_list:
                oracle_cur.execute(oracle_sql, row_li)
                try:
                    oracle_cur.execute(oracle_sql, row_li)
                except:
                    error_list.append([row_li, 'ORA'])
                    print('数据有误无法上传:', row_li)
            oracle_db.commit()

        print(time.strftime('%H:%M:%S', time.localtime(time.time())), ':已上传：', i, ' 行数据')
        value_list = []


for row in excel_message:
    i += 1
    value_list.append(row)
    list2db(0)

list2db(1)

# 对上传出错的行尝试再次上传
error_list_final = []
for li in error_list:
    try:
        if (li[1] == 'ORA'):
            oracle_cur.execute(oracle_sql, li[0])
            oracle_db.commit()
    except:
        error_list_final.append(li[0])
        print('数据有误无法上传:', li[0])

oracle_cur.close()
oracle_db.close()

print(time.strftime('%H:%M:%S', time.localtime(time.time())), ':上传完毕，共上传了：', i, ' 行数据')