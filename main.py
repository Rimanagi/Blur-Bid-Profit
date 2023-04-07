# pip install -r requirements
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# input here .txt filename
input_txt_file = open(input("Enter .txt filename: ") + ".txt", 'r')
print(input_txt_file)
# input here .xlsx filename
output_xlsx_file = load_workbook(input("Enter .xlsx filename: ") + ".xlsx")

if __name__ == '__main__':

    blur_bid_word = "Blur Bid: "
    ws = output_xlsx_file.active
    row, count = 2, 3
    value_to_compare = float(input("Минимальное желаемое значение Blur Bid: "))

    for line in input_txt_file:
        if count % 4 == 0 and blur_bid_word in line:
            line = line.split()
            blur_bid_value = float(line[-1])
            ws.cell(row=row, column=5).value = blur_bid_value
            # value to change                  ↓
            if float(blur_bid_value) > value_to_compare:
                ws.cell(row=row, column=5).fill = PatternFill("solid", fgColor="61f551")
            else:
                ws.cell(row=row, column=5).fill = PatternFill("solid", fgColor="FF4567")
            row += 1
        count += 1

    input_txt_file.close()
    output_xlsx_file.save("forOffers.xlsx")