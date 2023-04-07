# pip install -r requirements
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import linecache


# input here .txt filename
filename = input("Enter .txt filename: ") + ".txt"
input_txt_file = open(filename, 'r')
# input here .xlsx filename
output_xlsx_file = load_workbook(input("Enter .xlsx filename: ") + ".xlsx")
ws = output_xlsx_file.active


row = 2
line_number = 0
blur_bid_word = "Blur Bid Profit: "
value_to_compare = float(input("Минимальное желаемое значение Blur Bid: "))


for line in input_txt_file:
    line_number += 1

    if blur_bid_word in line:
        blur_bid_profit = float(line.split()[-1])

        if float(blur_bid_profit) >= value_to_compare:
            line = linecache.getline(filename, line_number - 1)
            blur_bid_value = float(line.split()[-1])
            ws.cell(row=row, column=5).value = blur_bid_value
            ws.cell(row=row, column=5).fill = PatternFill("solid", fgColor="61f551")
        else:
            break
        row += 1

linecache.clearcache()
input_txt_file.close()
output_xlsx_file.save("2.xlsx")