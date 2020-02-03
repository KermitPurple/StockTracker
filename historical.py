from openpyxl import load_workbook
from datetime import date
import yfinance as yf

def getData():
    startdate = date.fromisoformat("2020-01-06")
    enddate = date.fromisoformat("2020-01-31")
    datalist = [["Date", "Tag", "Company", "Shares", "Price", "Total",
        "Date", "Tag", "Company", "Shares", "Price", "Total",
        "Date", "Tag", "Company", "Shares", "Price", "Total", "Grand Total"]]
    tags = ["TSLA", "MJNA", "PLUG"]
    companies = ["Tesla, Inc.", "Medical Marijuana, Inc.", "Plug Power Inc."]
    num_of_shares = [8, 16985, 261]
    data = yf.download(tags[0] + " " + tags[1] + " " + tags[2],
            start=startdate.isoformat(),
            end=enddate.isoformat(),
            group_by="ticker")
    for i in range(len(data[tags[0]]["Open"])):
        grandtotal = 0
        datalist.append([])
        for j in range(len(tags)):
            price = data[tags[j]]["Open"][i]
            total = price * num_of_shares[j]
            grandtotal += total
            datalist[i + 1].append("insert date")
            datalist[i + 1].append(tags[j])
            datalist[i + 1].append(companies[j])
            datalist[i + 1].append(num_of_shares[j])
            datalist[i + 1].append(price)
            datalist[i + 1].append(total)
        datalist[i + 1].append(grandtotal)
    return datalist

def printdata(datalist):
    for i in range(len(datalist)):
        for j in range(len(datalist[i])):
            print(datalist[i][j], end =" ")
        print()
    
        

printdata(getData())
