import pandas as pd
from datetime import datetime, timedelta

urer_input = input('Введите названия дистрибутивов через пробел: ')
distr = urer_input.upper().split()

distr_numbers = []
distr_name = []

# Загрузка Excel файла
workbook = 'base.xlsx'

# Считать все листы в словарь
# sheets_dict = pd.read_excel(workbook, sheet_name=None)
xz = pd.read_excel(workbook)

# Назначить каждый лист в отдельную переменную
#sheet_SKBO = sheets_dict['СКБО']
#sheet_SKUO = sheets_dict['СКУО']
#sheet_SKJO = sheets_dict['СКЮО']
#sheet_SBOO = sheets_dict['СБОО']
#sheet_MZ = sheets_dict['МЗ']
#sheet_DOP = sheets_dict['Доп']



def check_oldest_date(file_path, sheet_name, column_name):
    """
    Функция проверяет самую старую дату в указанном столбце Excel файла и возвращает True, если эта дата старше 96 часов от текущего времени.

    Параметры:
    file_path (str): Путь к Excel файлу.
    sheet_name (str): Название листа в Excel файле.
    column_name (str): Название столбца, в котором необходимо проверить даты.

    Возвращает:
    tuple: Если найдена дата старше 96 часов, возвращает кортеж (столбец, строка, True), в противном случае (None, None, False).
    """
    # Загрузка Excel файла в Pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Находим самую старую дату в заданном столбце
    oldest_date = df[column_name].min()

    # Проверяем, является ли эта дата старше 96 часов от текущего времени
    current_time = datetime.now()
    if (current_time - oldest_date) >= timedelta(hours=96):
        # Определение столбца и строки для старой даты
        column = column_name
        row = df[df[column_name] == oldest_date].index[0]
        return (column, row, True)
    else:
        return (None, None, False)


x = check_oldest_date("base.xlsx", distr[0], "Дата")

print(x[1])

distr_numbers.append(x[1])
distr_numbers.append(x[1])
distr_numbers.append(x[1])

value = xz.at[x[1], "Дистр"]

print(value)

print(distr[0])
print(distr_numbers)