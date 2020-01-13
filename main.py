from yahoofinancials import YahooFinancials
from openpyxl import load_workbook
from datetime import date

def get_data():
    print("Getting Data ...")
    tags = ["TSLA", "MJNA", "PLUG"]
    companies = ["Tesla, Inc.", "Medical Marijuana, Inc.", "Plug Power Inc."]
    num_of_shares = [8, 16985, 261]
    big_list = [[f"Date: {date.today()}","Company","Tag","Price","Number of Shares","Total"]]
    gtotal = 0
    big_list.append([""])
    for i in range(len(tags)):
        tag = tags[i]
        print(f"\tGetting {tag}")
        company = companies[i]
        shares = num_of_shares[i]
        price = YahooFinancials(tag).get_current_price()
        total = shares * price
        gtotal += total
        inner_list = ["", company, tag, price, shares, total]
        big_list.append(inner_list)
    big_list.append(["","","","","Grand Total:", gtotal])
    print("Complete!")
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
