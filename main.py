#!/usr/bin/env python3
import openpyxl
from openpyxl import load_workbook
import sys
import os
import grafik_functions as g


CRED = "\033[91m"
CEND = '\033[0m'

# wskazanie komórki z datą grafiku
month_cell = 'D1'
# wskazanie kolumny z dniami miesiąca
g.collumn_of_days = 'A'
# wskazanie kolumny z godzinami
g.collumn_of_hours = 'C'
g.dateformat = '%d/%m/%Y'


try:
    grafik_filename = sys.argv[1]
except IndexError:
    print('Uruchom skrypt z nazwą pliku jako argumentem: \n' + CRED +
          'python3 ' + CRED + str(sys.argv[0]) + ' grafik.xlsx' + CEND)
    sys.exit(1)

try:
    wb = load_workbook(filename=grafik_filename, data_only=True)

    # pobranie pierwszego arkusza
    ws = wb.worksheets[0]
    # pobranie liczby dni w miesiącu
    num_days = g.get_days_numbers(ws, month_cell)
    # wskazane linijki zawierajacej pierwszy dyżur
    g.line_of_first_day = 5
    # wskazanie ostatniej linii z grafikiem
    g.line_of_last_day = num_days * 3 + g.line_of_first_day - 1

    try:
        g.month = g.get_month_year(ws, month_cell)[0]
        g.year = g.get_month_year(ws, month_cell)[1]
    except:
        print("Wskazano złą komórkę z miesiącem, bądź w nazwie miesiąca występuje literówka")
        sys.exit(1)

    directory_name = str(g.month) + '-' + str(g.year)
    if not os.path.exists(directory_name):
        os.makedirs(directory_name)
    opers = g.get_opers(ws)
    grafik = g.generate_grafik(opers, ws)

    oper_names = []
    for oper in opers:
        oper_names += [oper[0][0] + '.' + oper[1]]

    csv_data = g.generate_csv(grafik)

    csv_once = {}
    for oper in oper_names:
        csv_once["%s.csv" % oper] = g.generate_csv(grafik, oper)

    for x in csv_once.keys():
        g.save_to_csv(directory_name + '/' + x, csv_once[x])

    g.save_to_csv(directory_name + '/all_opers.csv', csv_data)

    calendar_opers = g.generate_ics(grafik)
    with open(directory_name + '/my.ics', 'w', newline='') as my_file:
        my_file.writelines(calendar_opers)
        my_file.close()

    zip = "zip -r " + directory_name + ".zip " + directory_name + "/"
    os.system(zip)




except openpyxl.utils.exceptions.InvalidFileException:
#    convert = input('Zły format pliku, czy przekonwertować (T/N)?')
#    if convert == 'T':
        grafik_fn = grafik_filename.split('.')
        grafik_fn = grafik_fn[0] + '.xlsx'
        g.convert_excel(grafik_filename, grafik_fn)
        #sys.exit(0)
        print("przekonwertowano grafik na format xmls")
        cmd = "python3 main.py " + grafik_fn
        os.system(cmd)

#    else:
#        sys.exit(1)



