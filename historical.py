from openpyxl import Workbook
from datetime import date, timedelta, datetime
import yfinance as yf
from os import system

def dateRange(start, end):
    dates = []
    for i in range((end-start).days + 1):
        dt = start + timedelta(days=i)
        data = yf.download("TSLA",
                start=dt.isoformat(),
                end=dt.isoformat(),
                group_by="ticker")
        if len(data) != 0:
            dates.append(dt)
    return dates

def getData():
    startdate = date.fromisoformat("2020-01-06")
    enddate = datetime.now().date()
    dates = dateRange(startdate, enddate)
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
    for i in range(len(data[tags[0]]["High"])):
        grandtotal = 0
        datalist.append([])
        for j in range(len(tags)):
            price = data[tags[j]]["High"][i]
            total = price * num_of_shares[j]
            grandtotal += total
            datalist[i + 1].append(dates[i])
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

def exportData(datalist):
    wb = Workbook()
    ws = wb.active
    for i in range(len(datalist)):
        ws.append(datalist[i])
    wb.save("History.xlsx")
    
        
dta = getData()
printdata(dta)
exportData(dta)
system("History.xlsx")
