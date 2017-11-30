import openpyxl
import re

# variables
headers = [['num', 'TV', 'Family name', 'BOM-code', 'Разрешение', 'Размер дисплея', 'Тип дисплея', 'Серия', 'Фильтрация', 'Особенности']]
letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']

# filters
tv_types = ['QLED', 'Full HD', 'Premium UHD TV', 'UHD', 'The Frame']

# regexp
is_curved = re.compile(r'Curved')
serie = re.compile(r'Q\d|Series \d')
screen_size = re.compile(r'E\d{2}')
find_num = re.compile(r'\d')
find_num2 = re.compile(r'\d{2}')

# r = re.search(screen_size, 'UE55MU6470UXRU')
# print(r.group())

# defs:
def find_serie(val1, val2):
    serie_raw = re.search(serie, val1) or re.search(serie, val2)
    if (serie_raw):
        return str(re.search(find_num, serie_raw.group()).group())

def find_screen_size(input_string):
    screen_raw = re.search(screen_size, input_string)
    if (screen_raw):
        res_screen = int(re.search(find_num2, screen_raw.group()).group())

        if (res_screen <= 49):
            return '40"~49"'
        elif (res_screen > 49 and res_screen <= 59):
            return '50"~59"'
        elif (res_screen > 59 and res_screen <= 69):
            return '60"~69"'
        elif (res_screen > 69):
            return 'диагональ более 70 дюймов'

def fill_sheet(sheet, array):
    for row in array:
        index = array.index(row) + 1
        for i in range(0, len(row)):
            sheet[letters[i] + str(index)].value = row[i]

def fill_array(filters, source_rows):
    res_array = headers + []

    for i in filters:
        for j in source_rows:
            row = list(j)
            tv_type = row[1].value
            if (tv_type == i):

                curved = re.search(is_curved, j[6].value)
                res_serie = find_serie(j[2].value, j[6].value)

                result_row = [
                    '',                                                         # num
                    tv_type,                                                    # TV
                    j[2].value,                                                 # Family name
                    j[3].value,                                                 # BOM-code
                    '',                                                         # Разрешение
                    find_screen_size(j[3].value),                               # Размер дисплея
                    'изогнутный' if curved else 'плоский',                      # Тип дисплея
                    'Серия ' + res_serie if res_serie else '',                  # Серия
                    '',                                                         # Фильтрация
                    ''                                                          # Особенности
                ]

                res_array.append(result_row)

    return res_array

wb = openpyxl.load_workbook('filters.xlsx')
sheet = wb.get_sheet_by_name('Product')
type_sheet = wb.get_sheet_by_name('Type')
resolution_sheet = wb.get_sheet_by_name('Resolution')
display_sheet = wb.get_sheet_by_name('Display')
tech_sheet = wb.get_sheet_by_name('Technologies')

# manipulate with cells
rows = list(sheet.rows)
rows.remove(rows[0])

# 1.pull tv types
fill_sheet(type_sheet, fill_array(tv_types, rows))

wb.save('filters.xlsx')
