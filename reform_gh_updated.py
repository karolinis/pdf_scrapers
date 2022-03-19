import csv
import tabula as tb
import re
import glob
import pdfplumber
import time
import os
import pandas as pd

def scrape_text(pdfName):
    with pdfplumber.open(pdfName) as file: #open file
        for page in file.pages: #extract text from every page
            txt = page.extract_text()
        return txt

def get_regex(text):
    date = ''
    for line in text.split('\n'):
        date_row = re.compile(r'(Date)')
        if date_row.search(line):
            date_line = line.split(':')
            date = date_line[-1]
            date = date.split('.')
            date = f'{date[-1]}-{date[-2]}-{date[-3]}'

        total_row = re.compile(r'(Total EUR)')
        if total_row.search(line):
            total_line = line.split(':')
            total = total_line[-1]
            total = total.replace('.','')
            total = total.replace(',','.')

    return date, total

def get_table_data(file):
    main_body = tb.read_pdf_with_template(file,'Invoice_28953.tabula-template.json',pandas_options={'header':None, 'index':None})
    try:
        delivery_ad = main_body[2]
    except:
        delivery_ad = ''

    join_address = []
    try:
        for row in delivery_ad[0]:
            join_address.append(str(row))
    except:
        join_address.append('')

    try:
        if join_address[-1] == 'I' or join_address[-1] == 'nan':
            join_address = join_address[:-1]

        delivery_address = ' '.join(join_address).split('Delivery Address: ')[-1]
    except:
        pass
    heading = main_body[1]
    heading = heading[0][0]#0th column 0th row


    service_info = main_body[0]
    service_price = 0
    list_service_info = []

    for index, row in service_info.iterrows():
        if str(row[0]) != 'nan':
            service_number = re.sub('[a-zA-Z]','',str(row[0]))
            desc = row[1]
            price = row[3]
            price = str(price).replace('.','')
            price = str(price).replace(',','.')

            try:
                if service_number in list_service_info[-1][0] and desc in list_service_info[-1][1]:
                    service_price += float(price)
                    list_service_info.pop() #this is to remove the first element of the list
                    #which makes sure we only get unique values in our list. We can then loop through it and print them out.
                else:
                    service_price = 0
                    service_price += float(price) #when service number is not in the list, it means that it's a
                #transportational cost, so we seperate it.
                list_service_info.append([service_number, desc, service_price])
            except:
                service_price += float(price)
                list_service_info.append([service_number, desc, service_price])

    return delivery_address, heading, list_service_info

if __name__ == '__main__':
    files = glob.glob('*.pdf')
    excel_rows = []
    for file in files:
        try:
            date, total = get_regex(scrape_text(file))
            delivery_address, heading, collection = get_table_data(file)

            for row in collection:
                service_nr, desc, price = row
                row_to_write = [date, delivery_address, heading, service_nr,desc, price, total, file]
                print(row_to_write)
                excel_rows.append(row_to_write)
            print(f'File {file} scanned.')
        except:
            print(f'File {file} failed.')
            df.loc[len(df.index)] = 'Not parsed:'
            df.loc[len(df.index)] = file

    print('--------Done--------')
    time.sleep(2)

    columns = ['Date', 'Address', 'Heading','Product Number','Description',  'Item price', 'Total', 'File']

    with open('output.csv', 'w',newline='', encoding='utf-8') as o:
        write = csv.writer(o)
        write.writerow(columns)
        write.writerows(excel_rows)