# import numpy as np
import xlwings as xw


def world():
    wb = xw.Book.caller()
    wb.sheets[3]['A1'].value = 'Hello World! by Mirsoat'


# Добавляем кнопку на лист для активации макроса
def add_button():
    # Получаем активный лист
    sheet = xw.Book.caller().sheets.active

    # Указываем диапазон, где будет находиться кнопка
    button_range = sheet.range("A1")

    # Добавляем кнопку и связываем ее с макросом
    button = sheet.buttons.add(button_range.left, button_range.top, button_range.width, button_range.height)
    button.on_action = "highlight_column"
    button.text = "Выделить столбец"


# Запускаем макрос
if __name__ == "__main__":
    xw.Book("Твой_файл_с_таблицей.xlsx").set_mock_caller()
    add_button()