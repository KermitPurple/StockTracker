from openpyxl import load_workbook
from datetime import date

def get_data():
    print("Getting Data ..." , end  = " ")
    tags = ["TSLA", "MJNA", "PLUG"]
    companies = ["Tesla, Inc.", "Medical Marijuana, Inc.", "Plug Power Inc."]
    num_of_shares = [8, 16851, 1000]
    big_list = [[f"date: {date.today()}","Company","Tag","Price","Number of Shares","total"]]
    gtotal = 0
    for i in range(len(tags)):
        tag = tags[i]
        company = companies[i]
        shares = num_of_shares[i]
        price = 0
        total = shares * price
        gtotal += total
        inner_list = ["", company, tag, price, shares, total]
        big_list.append(inner_list)
    big_list.append(["","","","","Grand Total:", gtotal])
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
