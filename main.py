from openpyxl import load_workbook

def get_data():
    return [
            [11, 12, 13, 14, 15],
            [21, 22, 23, 24, 25],
            [31, 32, 33, 34, 35],
            [41, 42, 43, 44, 45],
            [51, 52, 53, 54, 55]
            ]

def print_data(big_list):
    wb = load_workbook("sheet.xlsx")
    wa = wb.active
    for item in big_list:
        wa.append(item)
    wb.save("sheet.xlsx")

data = get_data()
print_data(data)
