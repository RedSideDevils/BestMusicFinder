import openpyxl
from colorama import Fore 
from banner import banner

print(Fore.MAGENTA + banner)

book = openpyxl.open('data.xlsx')

data =[]

sheet = book.active

row_count = sheet.max_row
column_count = sheet.max_column

for i in range(1,row_count):
    ll = {
        "id" : "",
        "Name":"",
        "Group_Name" : "",
        "Date" : int()
    }
    ll["id"] = sheet[i][0].value
    ll["Name"] = sheet[i][1].value
    ll["Group_Name"] = sheet[i][2].value
    ll["Date"] = sheet[i][3].value

    data.append(ll)

inpt1 = input(Fore.YELLOW + "[+]Enter First  Date: ")
inpt2 = input(Fore.YELLOW + "[+]Enter Second Date: ")

result = []

for i in range(len(data)):
    if data[i]["Date"] >= int(inpt1) and data[i]["Date"] <= int(inpt2):
        result.append(data[i])

print("\n")

for n in result:
    print(f"{Fore.RED}Title: {n['Name']} | Group: {n['Group_Name']} | Date: {str(n['Date'])}")