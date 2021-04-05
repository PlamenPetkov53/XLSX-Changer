import xlrd
import xlsxwriter
print("Enter the directory of xlsx file: ")
path = input()
print("Enter name of xlsx file: ")
name_of_file = input()
#"C:\\Users\\Plamen\\Downloads\\Book.xlsx"
raw_string = r"{}".format(path)
current_path = f"{raw_string}\\{name_of_file}.xlsx"
workbook = xlrd.open_workbook(current_path)
worksheet = workbook.sheet_by_index(0)

all_rows = []
# Creating 2D Array and fill it with all xlsx information
for row in range(worksheet.nrows):
    current_row = []
    for col in range(worksheet.ncols):
        current_row.append(worksheet.cell_value(row, col))
    all_rows.append(current_row)

for row in range(0, len(all_rows)):
    # Check type of variable in current cell
    if type(all_rows[row][0]) and type(all_rows[row][4]) == str:
        # Fill Column A with concat string
        all_rows[row][0] = f"{all_rows[row][0]} {all_rows[row][4]}"

    # Check type of variable in current cell
    if type(all_rows[row][2]) and type(all_rows[row][3]) == float:
        # Fill Column B with summed values
        all_rows[row][1] = all_rows[row][2] + all_rows[row][3]

    # remove column E
    all_rows[row].remove(all_rows[row][4])

    # Remove Column C
    all_rows[row].remove(all_rows[row][2])

    # Remove Column D
    all_rows[row].remove(all_rows[row][2])

# Printing the whole array for test
print(all_rows)

path_new_file = raw_string = r"{}".format(path)
new_file = f"{path_new_file}\\Modified_{name_of_file}.xlsx"
new_workbook = xlsxwriter.Workbook(new_file)
new_sheet = new_workbook.add_worksheet()

# Filling the new sheet with information from all_rows 2D array
for row in range(len(all_rows)):
    for col in range(len(all_rows[0])):
        new_sheet.write(row, col, all_rows[row][col])
new_workbook.close()
