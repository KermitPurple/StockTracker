from openpyxl import Workbook
from datetime import date, timedelta, datetime
import yfinance as yf
from os import system

def dateRange(start, end):
    dates = []
    for i in range((end-start).days + 1):
        dt = start + timedelta(days=i)
        data = yf.download("BNED",
                start=dt.isoformat(),
                end=dt.isoformat(),
                group_by="ticker")
        if len(data) != 0:
            dates.append(dt)
    return dates

def getData():
    startdate = date.fromisoformat("2020-02-10")
    enddate = datetime.now().date()
    dates = dateRange(startdate, enddate)
    datalist = [["Date", "Tag", "Company", "Shares", "Price", "Total",
         "Tag", "Company", "Shares", "Price", "Total",
         "Tag", "Company", "Shares", "Price", "Total",
         "Tag", "Company", "Shares", "Price", "Total",
         "Tag", "Company", "Shares", "Price", "Total", "Grand Total"]]
    tags = ["IMCX.CN", "REKR", "PLUG", "BNED", "TOMZ"]
    companies = ["IMCX company name", "REKR company name", "Plug Power Inc.", "Barnes and Noble", "TOMI Environmental Solutions, Inc."]
    num_of_shares = [948, 175, 150, 100, 2389]
    data = yf.download(tags[0] + " " + tags[1] + " " + tags[2] + " " + tags[3] + " " + tags[4],
            start=startdate.isoformat(),
            end=enddate.isoformat(),
            group_by="ticker")
    for i in range(len(data[tags[0]]["High"])):
        grandtotal = 0
        datalist.append([])
        datalist[i + 1].append(dates[i])
        for j in range(len(tags)):
            price = data[tags[j]]["High"][i]
            total = price * num_of_shares[j]
            grandtotal += total
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
#printdata(dta)
exportData(dta)
system("History.xlsx")
