# This program will parse the excel file called "SOWC 2014 Stat Tables_Table 9.xlsx" /
# in a way that Python can read
import xlrd

def main():
    book = xlrd.open_workbook('SOWC 2014 Stat Tables_Table 9.xlsx')

    sheet = book.sheet_by_name('Table 9 ')

    # Outputs the range 0-302 and each number is an "index number" we can use.
    row_num = sheet.nrows

    count = 0
    data = {}

    # For each object in the list (range of sheet.nrows)
    # In python 2, range and xrange were different functions. In python 3 they are both under the same class.
    for i in range(row_num):
        if count < 15:
            if i > 14:
                # Takes the list that is each row and saves it to the row variable. This makes our /
                # code more readable.
                row = sheet.row_values(i)
                row_num = sheet.nrows

                # This is an index of the first row, it pulls out the country from each row we iterate over.
                country = row[1]  # since the english name is in column B we need index 1

                data[i, country] = {}  # Adds the country as a key to the data dictionary in the format you want
                count += 1

    print (data)

main()
