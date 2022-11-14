"""All rights reserved by @shseam13
follow me on-
    Github:    https://github.com/shseam13
    Facebook:  https://www.facebook.com/shajjadhossains1/
    LinkedIn:  https://bd.linkedin.com/in/shajjad-hossain-seam-b6ba641b0
"""
#from random import randint
import openpyxl

workbook = openpyxl.Workbook()
sheet =workbook.active

data = (
    ("Days", "RD_for_TyofN", "Ty of Newsday", "RD_for_Demand", "Demand", "Rev._from_Sales", "Lost Profit", "Salv._fm_Sa.ofSc.", "Daily Profit"),)

for row in data:
    sheet.append(row)

# below column --> Day
for val in range(2,14):
    sheet[f"A{val}"].value = val-1

# below column --> Random Digit for Types of Newsday
#--> static code:
sheet["B2"].value = 94
sheet["B3"].value = 77
sheet["B4"].value = 49
sheet["B5"].value = 45
sheet["B6"].value = 43
sheet["B7"].value = 32
sheet["B8"].value = 49
sheet["B9"].value = 100
sheet["B10"].value = 16
sheet["B11"].value = 24
sheet["B12"].value = 31
sheet["B13"].value = 14

# below column --> Types of Newsdays
for val in range(2,14):
    if sheet[f"B{val}"].value < 36:
        sheet[f"C{val}"].value = "good"
    elif sheet[f"B{val}"].value < 81:
        sheet[f"C{val}"].value = "fair"
    elif sheet[f"B{val}"].value < 101:
        sheet[f"C{val}"].value = "poor"

# below column --> Random Digit for Demand
#--> static code:
sheet["D2"].value = 80
sheet["D3"].value = 20
sheet["D4"].value = 15
sheet["D5"].value = 88
sheet["D6"].value = 98
sheet["D7"].value = 65
sheet["D8"].value = 86
sheet["D9"].value = 73
sheet["D10"].value = 24
sheet["D11"].value = 60
sheet["D12"].value = 60
sheet["D13"].value = 29

# below column --> Demand
for val in range(2,14):
    if sheet[f"C{val}"].value == "good":
        if sheet[f"D{val}"].value < 4:
            sheet[f"E{val}"].value = 40
        elif sheet[f"D{val}"].value < 9:
            sheet[f"E{val}"].value = 50
        elif sheet[f"D{val}"].value < 24:
            sheet[f"E{val}"].value = 60
        elif sheet[f"D{val}"].value < 44:
            sheet[f"E{val}"].value = 70
        elif sheet[f"D{val}"].value < 79:
            sheet[f"E{val}"].value = 80
        elif sheet[f"D{val}"].value < 94:
            sheet[f"E{val}"].value = 90
        elif sheet[f"D{val}"].value < 101:
            sheet[f"E{val}"].value = 100

    elif sheet[f"C{val}"].value == "fair":
        if sheet[f"D{val}"].value < 11:
            sheet[f"E{val}"].value = 40
        elif sheet[f"D{val}"].value < 29:
            sheet[f"E{val}"].value = 50
        elif sheet[f"D{val}"].value < 69:
            sheet[f"E{val}"].value = 60
        elif sheet[f"D{val}"].value < 89:
            sheet[f"E{val}"].value = 70
        elif sheet[f"D{val}"].value < 97:
            sheet[f"E{val}"].value = 80
        elif sheet[f"D{val}"].value < 101:
            sheet[f"E{val}"].value = 90

    elif sheet[f"C{val}"].value == "poor":
        if sheet[f"D{val}"].value < 45:
            sheet[f"E{val}"].value = 40
        elif sheet[f"D{val}"].value < 67:
            sheet[f"E{val}"].value = 50
        elif sheet[f"D{val}"].value < 83:
            sheet[f"E{val}"].value = 60
        elif sheet[f"D{val}"].value < 95:
            sheet[f"E{val}"].value = 70
        elif sheet[f"D{val}"].value < 101:
            sheet[f"E{val}"].value = 80

# below column --> Revenew From Sales
for val in range(2,14):
    sheet[f"F{val}"].value = sheet[f"E{val}"].value * 0.5

# below columnn --> Lost Profit
for val in range(2,14):
    if sheet[f"E{val}"].value > 70:
        sheet[f"G{val}"].value = (sheet[f"E{val}"].value - 70) * 0.17
    else:
        sheet[f"G{val}"].value = 0

# below column --> Salvage from Sales of Scrap
for val in range(2,14):
    if sheet[f"E{val}"].value < 70:
        sheet[f"H{val}"].value = (70 - sheet[f"E{val}"].value) * 0.05
    else:
        sheet[f"H{val}"].value = 0

# below column --> Daily Profit
for val in range(2,14):
    sheet[f"I{val}"].value = sheet[f"F{val}"].value - (70*0.33) - sheet[f"G{val}"].value + sheet[f"H{val}"].value

# save workbook
workbook.save(filename="output.xlsx")