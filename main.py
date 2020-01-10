from openpyxl import load_workbook
from datetime import date

def get_data():
    print("Getting Data ..." , end  = " ")
    tags = ["TSLA", "MJNA", "PLUG"]
    big_list = [[f"date: {date.today()}","Company","Tag","Price","Number of Shares","total"]]
    for i, tag in enumerate(tags):
        inner_list = [str(i) + '1',str(i) + '2',str(i) + '3',str(i) + '4',]
        big_list.append(inner_list)
    print("Complete")
    return big_list

def print_data(big_list):
    print("Printing Data ...", end = " ")
    wb = load_workbook("sheet.xlsx")
    wa = wb.active
    for item in big_list:
        wa.append(item)
    wb.save("sheet.xlsx")
    print("Complete")

data = get_data()
print_data(data)
