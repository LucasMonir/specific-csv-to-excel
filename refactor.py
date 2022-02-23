import os
import openpyxl
import csv

path = os.getcwd()
input_format = "csv"
output_format = "xlsx"
header = ["*","*", "*", "*", "*"]
csv_list = []

def chunks(L, n):
    return [L[x : x + n] for x in range(0, len(L), n)]

def get_directory_files():
    return os.listdir(path)

def format_ticket(ticket):
    while not ticket[0].isnumeric():
        ticket = ticket[1:]
    return ticket;

def clean_row(row):
    new_row = [];
    for item in row:
        if item != '-':
            for splitted in item.split(';'):
                new_row.append(splitted)
    new_row.__delitem__(len(new_row) - 1)
    
    return new_row;

def read_lines(file):
    with open(file) as to_convert:
        reader = csv.reader(to_convert)
        for row in reader:
            ticket = format_ticket(row[0].split(';')[0]);
            row[0] = row[0].split(';')[1].replace('"', '')
            row = clean_row(row);
            items =  chunks(row, 4)
            # Appendear os tickets :/

            if len(items) > 1:
                csv_list.append(items)


def get_file_extension(file):
    return file.split(".")[1] if len(file.split(".")) else '';

def csv_to_xlsx(file):
    read_lines(file)
    result = openpyxl.Workbook()
    sheet = result.active
    sheet.append(header)

    for row in csv_list:
        for group in row:
            sheet.append(group)

    new_file = f'{path}/{file.split(".")[0]}.{output_format}'
    result.save(new_file)
    print("File saved: " + file.split(".")[0] + ".xlsx")

def convert_to_xlsx():
    for file in get_directory_files():
        file_extension = get_file_extension(file)
        if os.path.isfile(file) and input_format in file_extension:
            csv_to_xlsx(file)

def main():
    convert_to_xlsx()


            # Appendear os tickets :/
# main();
