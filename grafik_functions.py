from string import ascii_lowercase
from calendar import monthrange
from datetime import datetime
from datetime import timedelta
from dateutil import tz
import csv
from ics import Calendar, Event
import pyexcel as p

collumn_of_hours = None
collumn_of_days = None
line_of_first_day = 0
line_of_last_day = 0
month = None
year = None
dateformat = None


# znalezienie numeru miesiąca
def month_converter(month):
    months = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
              "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"]
    return months.index(month) + 1


def convert_excel(file, new_file):
    p.save_book_as(file_name=file,
                   dest_file_name=new_file)


# pobranie listy operatorów wraz z ich kolumnami w grafiku
def get_opers(ws):
    operators = []
    for letter in ascii_lowercase:
        name_cell = letter + '2'
        surname_cell = letter + '3'
        if ws[name_cell].value is not None:
            if ws[name_cell].value == 'Obsadzony':
                break
            else:
                operator = [ws[name_cell].value, ws[surname_cell].value, letter]
                operators.append(operator)
    return operators


# pobranie daty z grafiku
def get_month_year(ws, cell):
    splitted = ws[cell].value.split()
    year = int(splitted[1])
    month = month_converter(splitted[0])
    return month, year


# pobranie liczby dni w miesiącu
def get_days_numbers(ws, cell):
    month = get_month_year(ws, cell)[0]
    year = get_month_year(ws, cell)[1]
    num_days_of_month = monthrange(year, month)[1]
    return num_days_of_month


# wskazanie wiersza z dniem miesiąca
def get_cell_with_day(hours, row):
    if hours == '8-16':
        return row
    elif hours == '16-22':
        return row - 1
    elif hours == '22-8':
        return row - 2


# pobranie godzin dla zmiany
def get_shifts(hours):
    timeformat = '%H'
    hours_splitted = hours.split('-')
    start = datetime.strptime(hours_splitted[0], timeformat)
    stop = datetime.strptime(hours_splitted[1], timeformat)
    if hours == '8-16':
        return 'Ranek', start, stop
    elif hours == '16-22':
        return 'Popoludniu', start, stop
    elif hours == '22-8':
        return 'Nocka', start, stop


# generowanie listy z grafikiem
def generate_grafik(opers_list, ws) -> object:
    grafik = []
    for oper in opers_list:
        # kolumna z grafikiem dla operatora
        oper_column = oper[2]
        # imie i nazwisko w formacie I.Nazwisko
        oper_name = oper[0][0] + '.' + oper[1]
        for i in range(line_of_first_day, line_of_last_day + 1):
            checked_cell = oper_column + str(i)
            checked_cell_value = str(ws[checked_cell].value)
            checked_cell_value = checked_cell_value.lower()
            # jeżeli w komórce jest 'x' czyli dyżur
            if checked_cell_value == 'x':
                hours_cell = collumn_of_hours + str(i)
                day_of_month = collumn_of_days + str(get_cell_with_day(ws[hours_cell].value, i))
                shift = get_shifts(ws[hours_cell].value)
                shift_name = shift[0]
                start_of_shift = shift[1]
                stop_of_shift = shift[2]
                date = str(ws[day_of_month].value) + "/" + str(month) + "/" + str(year)
                start_date = datetime.strptime(date, dateformat)
                start_date = start_date.date()
                if stop_of_shift.hour == 8:
                    stop_date = start_date + timedelta(days=1)
                else:
                    stop_date = start_date
                grafik += [(oper_name, start_date, start_of_shift, stop_date, stop_of_shift, shift_name)]
    return grafik


def generate_csv(grafik, oper=None) -> object:
    csv_data = []
    csv_header = ["Subject", "Start Date", "Start Time", "End Date", "End Time", "All Day Event", "Description",
                  "Location",
                  "Private"]
    all_day_event = "FALSE"
    location = 'CBPIO'
    private = 'FALSE'
    for i in range(len(grafik)):
        if oper:
            if grafik[i][0] == oper:
                subject = "PCSS: " + grafik[i][0]
                start_date = str(grafik[i][1])
                start_time = str(datetime.time(grafik[i][2]))
                end_date = str(grafik[i][3])
                end_time = str(datetime.time(grafik[i][4]))
                description = grafik[i][0]
                csv_data.append(
                    [subject, start_date, start_time, end_date, end_time, all_day_event, description, location,
                     private])
        else:
            subject = "PCSS: " + grafik[i][0]
            start_date = str(grafik[i][1])
            start_time = str(datetime.time(grafik[i][2]))
            end_date = str(grafik[i][3])
            end_time = str(datetime.time(grafik[i][4]))
            description = grafik[i][0]
            csv_data.append(
                [subject, start_date, start_time, end_date, end_time, all_day_event, description, location, private])

    csv_data.insert(0, csv_header)
    return csv_data


def save_to_csv(filename, data):
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',')
        writer.writerows(data)
    file.close()


# konwersja daty na polski timezone
def convert_date_ics(date):
    date_format = '%Y-%m-%d %H:%M:%S'
    due_date = datetime.strptime(date, date_format).replace(tzinfo=tz.gettz('Europe/Warsaw'))
    due_date = due_date.astimezone(tz.tzutc())
    return due_date.strftime(date_format)


def generate_ics(grafik):
    calendar = Calendar()
    for i in range(len(grafik)):
        dyzur = Event()
        dyzur.name = grafik[i][0]
        start = str(grafik[i][1]) + ' ' + str(datetime.time(grafik[i][2]))
        stop = str(grafik[i][3]) + ' ' + str(datetime.time(grafik[i][4]))
        dyzur.begin = convert_date_ics(start)
        dyzur.end = convert_date_ics(stop)
        calendar.events.add(dyzur)
    return calendar
