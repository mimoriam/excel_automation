import openpyxl
from openpyxl.styles import Alignment
import os


def copy_range(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    for i in range(startRow, endRow + 1, 1):
        rowSelected = []
        for j in range(startCol, endCol + 1, 1):
            rowSelected.append(sheet.cell(row=i, column=j).value)
        rangeSelected.append(rowSelected)

    return rangeSelected


def paste_range(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
    countRow = 0
    for i in range(startRow, endRow + 1, 1):
        countCol = 0
        for j in range(startCol, endCol + 1, 1):
            sheetReceiving.cell(row=i, column=j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1


def main():
    file = "2019 Fluid dynamics Morning.xlsx"
    cleaned_file = "2019 Fluid dynamics Morning Cleaned Data.xlsx"

    if os.path.exists(cleaned_file):
        os.remove(cleaned_file)

    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb['Result Sheet']

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    sheet2 = wb2['Sheet']
    sheet2.title = 'Results Sheet'

    # for rowOfCellObjects in sheet['A25': 'A51']:
    #     for cellObj in rowOfCellObjects:
    #         registration_num_data.append(cellObj.value)
    #
    # print(registration_num_data)

    registration_range = copy_range(1, 25, 1, 51, sheet)
    paste_range(1, 2, 1, 26, sheet2, registration_range)

    name_of_students_range = copy_range(2, 25, 2, 51, sheet)
    paste_range(2, 2, 2, 26, sheet2, name_of_students_range)

    total_marks_range = copy_range(8, 25, 8, 51, sheet)
    paste_range(3, 2, 3, 26, sheet2, total_marks_range)

    grades_range = copy_range(9, 25, 9, 51, sheet)
    paste_range(4, 2, 4, 26, sheet2, grades_range)

    for row in ws2.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')

    sheet2['A1'] = "registration_num"
    sheet2['B1'] = "student_name"
    sheet2['C1'] = "total_num"
    sheet2["D1"] = "grade"

    wb2.save(filename=cleaned_file)


if __name__ == '__main__':
    main()
