import xlrd
import re


def DataFromExcel(path):
    inputWorkbook = xlrd.open_workbook(path)
    inputWorksheet = inputWorkbook.sheet_by_index(0)

    row = inputWorksheet.nrows
    dataSupplier = []

    for i in range(1,row):
        dataSupplier.append(inputWorksheet.cell_value(i, 1))
    return dataSupplier
