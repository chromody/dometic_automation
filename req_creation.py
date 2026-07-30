from datetime import datetime
import openpyxl
import pytz

# What needs to be automatically generated: Needs to automatically capture current date
# What needs to be input: Required Delivery Date (Can be blank), Source, GL Code, and Purpose
# Needs to prompt for item entry:
#	Item Description
#	Item Number
#	Quantity
#	Cost
# This can be repeated in a loop 24 times in total, afterwards the form is full

file_name = input("What would you like the name of the file to be: ")

print("Loading template form /mnt/c/Users/Jesvil/OneDrive - Dometic Group/Documents/Requisition Forms\n")

# 1. Open the existing template.xlsx workbook
workbook = openpyxl.load_workbook("/mnt/c/Users/Jesvil/OneDrive - Dometic Group/Documents/Requisition Forms/Template.xlsx")

# 2. Select the specific worksheet you want to edit
sheet = workbook["SHEET1"]

# 3. Pulling formatted date
print("Loading Current Date: ", end="")

current_date = datetime.now()
formatted_date = current_date.strftime("%-m/%-d/%Y")

print(formatted_date, "\n")

sheet["B3"] = formatted_date

# 4. Asking for information
req_deliver_date = input("Enter the required delivery date: ")
sheet["C5"] = req_deliver_date

source = input("Enter the source for the parts: ")
sheet["C42"] = source

gl_code = input("Enter the GL Code (often 4345): ")
sheet["B43"] = gl_code

purpose = input("Enter the purpose of all the parts: ")
sheet["C7"] = purpose

# 5. Entering loop for item fill out, we start at row 16 and end at row 39 (24 in total items possible), from column A to column D
starting_row = 16
if_continue = ""
for i in range(24):
    item_desc = input("Enter the item description: ")
    item_numb = input("Enter the item number: ")
    quantity = input("Enter the quantity: ")
    cost = input("Enter the cost: ")

    sheet["C" + str(starting_row + i)] = item_desc
    sheet["B" + str(starting_row + i)] = item_numb
    sheet["A" + str(starting_row + i)] = quantity
    sheet["D" + str(starting_row + i)] = cost

    if_continue = input("Would you like to make another part (Y for yes and N for no):")
    if (if_continue == "N"):
        break

# 4. Save your modifications back to the file
workbook.save("/mnt/c/Users/~/" + file_name + ".xlsx")
print("Workbook '" + file_name + "' successfully opened and updated!")

