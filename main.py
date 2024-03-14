import copy
import csv
import json

from excel2json import convert_from_file
import os
import glob
import re



pattern = r'\b\d{7,}\b'


def convert_file():
    """
    конвертирует xlsx в json
    :return:
    """
    file = []
    path = os.getcwd()
    for filename in glob.glob(os.path.join(path, '*.xlsx')):
        file.append(filename)
    filename = file[0]
    convert_from_file(filename)



def write_numbers():
    path = os.getcwd()
    for file in glob.glob(os.path.join(path, '*.json')):
        with open(file, 'r', encoding='utf-8') as file:
            numbers_list=json.load(file)

    result_numbers=[]
    for i in numbers_list:
        numbers_dict = {}
        one_owner_numbers=[]
        numbers=re.findall(pattern, i['Телефоны'])
        for number in numbers:
            if number.startswith('7'):
                number=number.replace('7','8',1)
            prefix=number[1:4]
            if prefix in ['900', '901', '902', '904', '908', '950', '951', '952', '953', '958', '969', '977', '991', '992', '993', '994']:
                operator = 'Tele2'
            else:
                operator='unknown'
            numbers_dict['number']=number
            numbers_dict['operator']=operator
            one_owner_numbers.append(copy.deepcopy(numbers_dict))
            numbers_dict.clear()
        result_numbers.append(one_owner_numbers)

    with open('результат.csv', 'w', newline='', encoding='utf-8-sig') as file:
        writer=csv.writer(file, delimiter=';')
        writer.writerow(['Теле2', 'Другие'])
        tele2_rows=[]
        unknown_rows=[]
        for i in result_numbers:
            for j in i:
                if j['operator']=='Tele2':
                    tele2_rows.append(j['number'])
                else:
                    unknown_rows.append(j['number'])

        if len(tele2_rows)>len(unknown_rows):
            flag = True
        else:
            flag=False
        if flag:
            for i in range(len(tele2_rows)):
                number_tele2=tele2_rows[i]
                try:
                    unknown_number=unknown_rows[i]
                except:
                    unknown_number=''
                writer.writerow([number_tele2, unknown_number])
        else:
            for i in range(len(unknown_rows)):
                unknown_number=unknown_rows[i]
                try:
                    number_tele2=tele2_rows[i]
                except:
                    number_tele2=''
                writer.writerow([number_tele2, unknown_number])


def delete_json():
    path = os.getcwd()
    for file in glob.glob(os.path.join(path, '*.json')):
        os.remove(file)


convert_file()
write_numbers()
delete_json()