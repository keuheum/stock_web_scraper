#The second line of the excel should contain the name of the stock and the second line should contain the stock code.

import win32com.client
from bs4 import BeautifulSoup 
import requests

x = 2
col = 2

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('path')
ws = wb.ActiveSheet

while ws.Cells(x,col).Value !=None:
    coad = (str(int(ws.Cells(x, col).Value)))
    link = ("http://vip.mk.co.kr/newSt/price/price.php?stCode=" + coad.zfill(6) )
    req = requests.get(link)
    soup = BeautifulSoup(req.text, "html.parser")

    ws.Cells(x, 3).Value = (soup.find("font" , class_="f5_r")).text #시가[market price]
    ws.Cells(x, 4).Value = (soup.find("span" , id="disArrD2[11]")).text#고가[high price]
    ws.Cells(x, 5).Value = (soup.find("span", id="disArrD3[1]")).text#저가[low price]
    closing_price = soup.find("font", class_="f3_r")#종가[closing price]
    title = soup.find("font", class_="f1").text

    if closing_price == None:
        ws.Cells(x, 6).Value = (soup.find("font", class_="f3_b")).text
        ws.Cells(x, 7).Value = "하락세[downtrend]"
    else:
        ws.Cells(x, 6).Value = closing_price.text
        ws.Cells(x, 7).Value = "상승세[uptrend]"

    x = x + 1

print(title)
wb.Save()
excel.Quit()
