# This program will parse the excel file called "SOWC 2014 Stat Tables_Table 9.xlsx" /
# in a way that Python can read
import xlrd
import pprint

book = xlrd.open_workbook('SOWC 2014 Stat Tables_Table 9.xlsx')

sheet = book.sheet_by_name('Table 9 ')

# Outputs the range 0-302 and each number is an "index number" we can use.
row_num = sheet.nrows

# Main list
list = []
# Dictionary
data = {}

# we no longer need a count statement to count our code because we know the data begins at line 14
for i in range(14, row_num):
    """
    Takes the list that is each row and saves it to the row variable. This makes our /
    code more readable.
    """
    row = sheet.row_values(i)
    country = row[1]  # for every iteration of i this prints the INDEX 1 of row, not the row 1
    if country == '':
        break
    elif country == 'Zimbabwe':
        break
    else:
        data[country] = {   # Adds the country as a key to the data dictionary in the format you want
            'Child_labor': {
                'total': [row[4],row[5]],
                'male': [row[6],row[7]],
                'female': [row[8],row[9]],
            },
            'child_marriage': {
                'married_by_15': [row[10],row[11]],
                'married_by_18': [row[12],row[13]],
            },
        }
        list.append(data)
        data = {} # This gives you the ability to index the dictionary entries within the list

pprint.pprint(list)


