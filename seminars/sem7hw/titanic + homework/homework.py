import xlwings as xw
import math


def main():
    wb = xw.Book.caller()


@xw.func
def dif_module(num1, num2):
    if num1 or num2:
        return math.fabs(num1-num2)
    else:
        return f'input numbers, fasta!'


@xw.sub
def color_area():
    wb = xw.Book.caller()
    area = wb.selection
    area.color(64, 224, 208)





if __name__ == "__main__":
    xw.Book("titanic.xlsm").set_mock_caller()
    main()
