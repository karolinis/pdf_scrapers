import csv
import tabula as tb
import re
import glob
import pdfplumber
import time

def scrape_text(pdfName):
    tmp = ''
    c = 0
    with pdfplumber.open(pdfName) as file: #open file
        for page in file.pages: #extract text from every page
            txt = page.extract_text()
            c += 1
            tmp += txt #if you need to parse all pages, push return back 4 spaces and: return tmp
            return txt


def get_regex(text):
    date = ''
    total = ''
    invoice = ''
    final_payment = ''
    for line in text.split('\n'):
        date_row = re.compile(r'(Dateofinvoice|'
                              r'Date of invoice|'
                              r'Date)')
        if date_row.search(line):
            date_line = line.split(':')
            date = date_line[-1]



        total_row = re.compile(r'(Total EUR|'
                               r'TotalEUR)')
        if total_row.search(line):
            total_line = line.split(':')
            total = total_line[-1]


        final_payment_row = re.compile(r'(Final payment date|'
                                       r'Finalpaymentdate)')
        if final_payment_row.search(line):
            final_payment_line = line.split(' ')
            final_payment = final_payment_line[-1]


        invoice_row = re.compile(r'(Invoiceno.|'
                                 r'Invoice no.)')
        if invoice_row.search(line):
            invoice_line = line.split(':')
            invoice = invoice_line[-1]
            print(line)


    print(date,total, final_payment, invoice)
    return date, total, final_payment, invoice

def get_table_data(file):

    main_body = tb.read_pdf_with_template(file,'tabula-template.json',pandas_options={'header':None, 'index':None},encoding='utf-8', pages=1)

    heading = main_body[0]
    heading = heading[0][0]#0th column 0th row

    service_info = main_body[1]
    numbers = []
    description = []
    services = []
    for nr,row in enumerate(service_info[0]):
        if str(row) != 'nan':
            service_main_row = service_info.loc[nr]
            service_nr = service_main_row[0]
            service_name = service_main_row[1]
            service_total = service_main_row[4]
            services.append(['','','','',service_nr,service_name,service_total])
        try:
            numbers.append(row[0])
        except:
            numbers.append(str(row))

    for row in service_info[1]:
        description.append(str(row))

    ND = list(zip(numbers, description))
    collection = []
    for order in ND:
        s_order = order[1].split(' ')
        if order[0] == 'nan' and len(s_order) == 2:
            order_id, price = s_order[0], s_order[1]
            collection.append([order_id,price])
        elif order[0] == 'nan' and len(s_order) == 3:
            price = s_order[1]+s_order[2]
            order_id = s_order[0]
            collection.append([order_id, price])

    return heading, collection, services


if __name__ == '__main__':
    try:
        files = glob.glob('*.pdf')
        excel_rows = []
        for file in files:
            try:
                date, total, final_payment, invoice = get_regex(scrape_text(file))
            except Exception as e:
                print(e,'\nError at get_regex')
                continue
            try:
                heading, collection, services = get_table_data(file)
            except Exception as e:
                print(e, 'TEST get_table_data failed.')
                continue
            try:
                excel_rows.append([date, final_payment, heading, invoice, total,'',''])

                for service in services:
                    excel_rows.append(service)

                for order in collection:
                    excel_rows.append(['','','','','',order[0],order[1]])
            except:
                print(f'Cannot append {file} to excel rows')
                excel_rows.append(f'{file} not parsed')
                continue

        columns = ['Date','Final payment date','Heading','Invoice Number','Total price','Order Number','Order cost']
        with open('output.csv','w',encoding='UTF-8',newline='') as out:
            writer = csv.writer(out)
            writer.writerow(columns)
            writer.writerows(excel_rows)

    except Exception as e:
        print(e, 'TEST main failed.')
        time.sleep(5)
