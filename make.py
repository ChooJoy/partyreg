# coding=utf-8

import xlrd
import xlwt
from xlutils.copy import copy

LINE_LENGTH = 40


#нахождение позиций полей в эксель файле
# c = column
# r = row

STR_C = 69
STR_R = 3
SURNAME_C = 15
SURNAME_R = 12
NAME_C = 15
NAME_R = 14
F_NAME_C = 15
F_NAME_R = 17
BIRTHDAY_C = 23
BIRTHDAY_R = 24
BIRTHPLACE_C = 0
BIRTHPLACE_R = 28
PASSPORT_CODE_C = 22
PASSPORT_CODE_R = 36
PASSPORT_C = 30
PASSPORT_R = 39
PASSPORT_DATE_C = 18
PASSPORT_DATE_R = 41
PASSPORT_KEM_C = 18
PASSPORT_KEM_R = 43
PASSPORT_PODR_C = 24
PASSPORT_PODR_R = 49

INDEX_C = 24
INDEX_R = 12
REGION_C = 87
REGION_R = 12
RAYON_C = 0
RAYON_R = 16
RAYON_NAME_C = 36
RAYON_NAME_R = 16 
CITY_C = 0
CITY_R = 22
CITY_NAME_C = 36
CITY_NAME_R = 22 
TOWN_C = 0
TOWN_R = 26
TOWN_NAME_C = 36
TOWN_NAME_R = 26
STREET_C = 0
STREET_R = 32
STREET_NAME_C = 36
STREET_NAME_R = 32
DOM_C = 0
DOM_R = 38
DOM_NUM_C = 33
DOM_NUM_R = 38 
KORP_C = 63
KORP_R = 38
KORP_NUM_C = 96 
KORP_NUM_R = 38
FLAT_C = 33
FLAT_R = 40
FLAT_NUM_C = 96
FLAT_NUM_R = 40

page_counter = 3


# функция создает один документ

def make_one(row):

    surname = row[1]
    name = row[2]
    f_name = row[3]
    birthday = row[4]
    birthplace = row[5]
    passport_code = "21"
    passport = row[8]
    passport_date = row[9]
    passport_kem = row[10]
    passport_podr = row[11]

    index = row[12]
    region_text = row[13]
    region = row[14]
    rayon = row[15]
    rayon_name = row[16]
    city = row[17]
    city_name = row[19]
    town = row[20]
    town_name = row[21]
    street = row[24]
    street_name = row[26]
    dom = row[27]
    dom_num = row[28]
    korp = row[29]
    korp_num = row[31]
    flat = row[32]
    flat_num = row[33]

    print region_text
    print surname, name, f_name


    shab = xlrd.open_workbook('shab.xls',formatting_info=True)
    wb = copy(shab)
    write_sheet = wb.get_sheet(4)

    borders = xlwt.Borders()  
    borders.left = xlwt.Borders.DOTTED 
    borders.right = xlwt.Borders.DOTTED
    borders.top = xlwt.Borders.DOTTED
    borders.bottom = xlwt.Borders.DOTTED
    borders.left_colour = 0x37 
    borders.right_colour = 0x37
    borders.top_colour = 0x37
    borders.bottom_colour = 0x37 
    style = xlwt.XFStyle()  
    style.borders = borders 

    #вставляем номер страницы
    global page_counter
    page_count = str(page_counter)
    while len(page_count) < 3:
        page_count = "0" + page_count
    i = 0
    while i<len(page_count):
        write_sheet.write(STR_R, STR_C+3*i, page_count[i], style)
        i += 1
    page_counter += 1

    #вставляем фамилию
    i = 0
    while i<len(surname):
        write_sheet.write(SURNAME_R, SURNAME_C+3*i, surname[i].upper(), style)
        i += 1
    
    #Вставляем имя
    i = 0
    while i<len(name):
        write_sheet.write(NAME_R, NAME_C+3*i, name[i].upper(), style)
        i += 1
    
    #Вставляем отчество
    i = 0
    while i<len(f_name):
        write_sheet.write(F_NAME_R, F_NAME_C+3*i, f_name[i].upper(), style)
        i += 1
    
    #Вставляем дату рождения
    i = 0
    if isinstance(birthday, float): birthday = str(int(birthday))
    while i<len(birthday):
        if i != 2 and i != 5:
            write_sheet.write(BIRTHDAY_R, BIRTHDAY_C+3*i, birthday[i], style)
        i += 1

    #Вставляем место рождения
    i = 0
    line_num = 0
    while i<len(birthplace):
        write_sheet.write(BIRTHPLACE_R+line_num*2, BIRTHPLACE_C+3*i-LINE_LENGTH*3*line_num, birthplace[i].upper(), style)
        i += 1
        if i == LINE_LENGTH: line_num = 1
    
    #Вставляем код паспорта
    i = 0
    while i<len(passport_code):
        write_sheet.write(PASSPORT_CODE_R, PASSPORT_CODE_C+3*i, passport_code[i], style)
        i += 1
    
    #Вставляем серию и номер паспорта
    i = 0
    passport_temp = passport[0:2] + " " + passport[2:]
    if len(passport_temp) != 12: print 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    while i<len(passport_temp):
        write_sheet.write(PASSPORT_R, PASSPORT_C+3*i, passport_temp[i], style)
        i += 1
    
    #Вставляем дату выдачи паспорта
    i = 0
    if isinstance(passport_date, float): passport_date = str(int(passport_date))
    while i<len(passport_date):
        if i != 2 and i != 5:
            write_sheet.write(PASSPORT_DATE_R, PASSPORT_DATE_C+3*i, passport_date[i], style)
        i += 1
    
    #Вставляем кем выдан паспорт
    i = 0
    line_num = 0
    small_line_length = 34
    small_line_num = 0
    while i<len(passport_kem):
        write_sheet.write(PASSPORT_KEM_R+line_num*2, PASSPORT_KEM_C+3*i-LINE_LENGTH*3*line_num, passport_kem[i].upper(), style)
        i += 1
        if i == small_line_length: line_num = 1
        if i == small_line_length+LINE_LENGTH: line_num = 2

    #Вставляем код подразделения из паспорта
    i = 0
    while i<len(passport_podr):
        if i != 3:
            write_sheet.write(PASSPORT_PODR_R, PASSPORT_PODR_C+3*i, passport_podr[i], style)
        i += 1
    
    
    
    write_sheet = wb.get_sheet(5)
    
    page_count = str(page_counter)
    while len(page_count) < 3:
        page_count = "0" + page_count
    i = 0
    while i<len(page_count):
        write_sheet.write(STR_R, STR_C+3*i, page_count[i], style)
        i += 1
    page_counter += 1
    
    #Вставляем индекс
    i = 0
    index = str(int(index))
    while i<len(index):
        write_sheet.write(INDEX_R, INDEX_C+3*i, index[i], style)
        i += 1
    
    #Вставляем регион
    i = 0
    if isinstance(region, float): region = str(int(region))
    while i<len(region):
        write_sheet.write(REGION_R, REGION_C+3*i, region[i], style)
        i += 1
    
    #Вставляем слово "район"
    i = 0
    while i<len(rayon):
        write_sheet.write(RAYON_R, RAYON_C+3*i, rayon[i].upper(), style)
        i += 1
    
    #Вставляем название района
    i = 0
    small_line_length = 28
    line_num = 0
    while i<len(rayon_name):
        write_sheet.write(RAYON_NAME_R+line_num*2, RAYON_NAME_C+3*i-LINE_LENGTH*3*line_num, rayon_name[i].upper(), style)
        i += 1
        if i == small_line_length: line_num = 1
    
    #Вставляем слово "город"
    i = 0
    while i<len(city):
        write_sheet.write(CITY_R, CITY_C+3*i, city[i].upper(), style)
        i += 1
    
    #Вставляем название города
    i = 0
    while i<len(city_name):
        write_sheet.write(CITY_NAME_R, CITY_NAME_C+3*i, city_name[i].upper(), style)
        i += 1
    
    #Вставляем слово "село"
    i = 0
    while i<len(town):
        write_sheet.write(TOWN_R, TOWN_C+3*i, town[i].upper(), style)
        i += 1
    
    #Вставляем название села
    i = 0
    small_line_length = 28
    line_num = 0
    while i<len(town_name):
        write_sheet.write(TOWN_NAME_R+line_num*2, TOWN_NAME_C+3*i-LINE_LENGTH*3*line_num, town_name[i].upper(), style)
        i += 1
        if i == small_line_length: line_num = 1
    
    #Вставляем слово "улица"
    i = 0
    while i<len(street):
        write_sheet.write(STREET_R, STREET_C+3*i, street[i].upper(), style)
        i += 1
    
    #Вставляем название улицы
    i = 0
    small_line_length = 28
    line_num = 0
    while i<len(street_name):
        write_sheet.write(STREET_NAME_R+line_num*2, STREET_NAME_C+3*i-LINE_LENGTH*3*line_num, street_name[i].upper(), style)
        i += 1
        if i == small_line_length: line_num = 1
    
    #Вставляем слово "дом"
    i = 0
    while i<len(dom):
        write_sheet.write(DOM_R, DOM_C+3*i, dom[i].upper(), style)
        i += 1
    
    #Вставляем номер дома
    i = 0
    if isinstance(dom_num, float): dom_num = str(int(dom_num))
    while i<len(dom_num):
        write_sheet.write(DOM_NUM_R, DOM_NUM_C+3*i, dom_num[i].upper(), style)
        i += 1
    
    #Вставляем слово "корпус"
    i = 0
    while i<len(korp):
        write_sheet.write(KORP_R, KORP_C+3*i, korp[i].upper(), style)
        i += 1
    
    #Вставляем номер корпуса
    i = 0
    if isinstance(korp_num, float): korp_num = str(int(korp_num))
    while i<len(korp_num):
        write_sheet.write(KORP_NUM_R, KORP_NUM_C+3*i, korp_num[i].upper(), style)
        i += 1
    
    #Вставляем слово "квартира"
    i = 0
    while i<len(flat):
        write_sheet.write(FLAT_R, FLAT_C+3*i, flat[i].upper(), style)
        i += 1
    
    #Вставляем номер квартиры
    i = 0
    if isinstance(flat_num, float): flat_num = str(int(flat_num))
    while i<len(flat_num):
        write_sheet.write(FLAT_NUM_R, FLAT_NUM_C+3*i, flat_num[i].upper(), style)
        i += 1
    file_name = 'done/' + region + '-' + surname + "_" + name + '.xls'
    wb.save(file_name)
    print file_name, 'готово!'    
    print '--------------------------------------------------------------'


rb = xlrd.open_workbook('data.xls',formatting_info=True)
sheet = rb.sheet_by_index(0)

for rownum in range(sheet.nrows):
    row = sheet.row_values(rownum)
    make_one(row)
