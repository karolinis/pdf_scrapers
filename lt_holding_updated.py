import tabula as tb
import re
import glob
import pdfplumber
import time
import pandas as pd

def scrape_text(pdfName):
    try:
        with pdfplumber.open(pdfName) as file:  # open file
            tmp = ''
            for page in file.pages:  # extract text from every page
                txt = page.extract_text()
                tmp = tmp + txt
            return tmp
    except Exception as e:
        print(e, 'TEST scrape_text failed.')
        time.sleep(5)


def get_regex(text):
    global date, total, invoice, due_date
    date = ''
    total = ''
    invoice = ''
    due_date = ''
    try:
        for line in text.split('\n'):
            date_row = re.compile(r'(Date)')
            if date_row.search(line):
                date_line = line.split(':')
                date = date_line[-1]

            ddate_row = re.compile(r'(due date|duedate)')
            if ddate_row.search(line):
                ddate_line = line.split(' ')
                due_date = ddate_line[-1]

            total_row = re.compile(r'(Total EUR|'
                                   r'TotalEUR)')
            if total_row.search(line):
                total_line = line.split(':')
                total = total_line[-1]

            invoice_row = re.compile(r'(Invoiceno.|'
                                     r'Invoice no.|'
                                     r'Nr. .......................)')
            if invoice_row.search(line):
                invoice_line = line.split(':')
                if len(invoice_line) == 2:
                    invoice = invoice_line[-1]
        return date, total, invoice, due_date

    except Exception as e:
        print(e, 'TEST get_regex failed, continuing.')
        time.sleep(5)


def get_table_data(file, flag):
    # credit note file: LT-Holding for credit notes.tabula-template.json
    if flag:
        main_body = tb.read_pdf_with_template(file, 'Invoice_28953.tabula-template.json',
                                       pandas_options={'header': None, 'index': None}, encoding='utf-8', pages=1)
    else:
        main_body = tb.read_pdf_with_template(file, 'Invoice_28953.tabula-template.json',
                                       pandas_options={'header': None, 'index': None}, encoding='utf-8', pages=1)

    service_info = main_body[0]
    #main rows don't have nulls in them, if row has no NaN's - if length of full_name is 0 = empty
    #create full_name from string row[1], if full_name is not empty, it means that full_name has ran a loop already
    #therefore it needs to be added to the name list
    service_price = 0
    list_service_info = []
    for index, row in service_info.iterrows():
        if str(row[0]) != 'nan':
            service_number = re.sub('[a-zA-Z]', '', str(row[0]))
            desc = row[1]
            price = row[3]
            price = str(price).replace('.', '')
            price = str(price).replace(',', '.')

            try:
                if service_number in list_service_info[-1][0] and desc in list_service_info[-1][1]:
                    service_price += float(price)
                    list_service_info.pop()  # this is to remove the first element of the list
                    # which makes sure we only get unique values in our list. We can then loop through it and print them out.
                else:
                    service_price = 0
                    service_price += float(price)  # when service number is not in the list, it means that it's a
                # transportational cost, so we seperate it.
                list_service_info.append([service_number, desc, service_price])
            except:
                service_price += float(price)
                list_service_info.append([service_number, desc, service_price])


    heading = main_body[1][0][0]

    contras = {
        1502: ['LT'],
        1503: ['LT'],  # transport
        1412: ['DK', 'DE', 'US', 'RE'],
        1416: ['DK', 'DE', 'US', 'RE'],  # transport
        1442: ['RCL'], 2824: ['PDE'],
        1447: ['SHO'], 1450: ['BAR'],
        1446: ['samples', 'SA']
    }

    acc = ''
    Kitchen = ['1502', '1412']
    Transport = ['1503', '1416']
    for service in list_service_info:
        if type(service) == list:
            if any(i in str(service[1]) for i in
                ['Kitchen', 'Reclamations', 'Market', 'goods', 'service', 'Fee', 'Service', 'Quality']):
                for key, values in contras.items():
                    for value in values:
                        if value in heading and str(key) != '1416':
                            acc = str(key)
                            vat_code = VAT(acc,Kitchen,Transport)
                            service.append(vat_code)
                service.append(acc)
                print(service)
            elif 'Transport' in str(service[1]):
                for key, values in contras.items():
                    for value in values:
                        if value in heading and str(key) != '1412':
                            acc = str(key)
                            vat_code = VAT(acc, Kitchen, Transport)
                            service.append(vat_code)
                service.append(acc)
                print(service)
    return list_service_info, heading

def VAT(acc, Kitchen, Transport):
    if acc in Transport:
        return 'IY25'
    elif acc in Kitchen:
        return 'IV25'
    else:
        return 'IV25/IY25'

def entry(date,total,invoice,due_date,heading,service_data):
    for service in service_data:

        df.insert = date
        df.iloc[-1] = due_date
        print(df)
        df.loc[row_index, 'Due date'] = due_date
        df.loc[row_index, 'Entry'] = file
        df.loc[row_index, 'Invoice Number'] = invoice
        df.loc[row_index, 'Total price'] = total
        df.loc[row_index, 'Text'] = f'Reform Supply & Logistics, UAB - {heading}'

        df.loc[row_index, 'Order 1 name'] = service[0]
        df.loc[row_index, 'Price 1'] = service[1]
        df.loc[row_index, 'Account 1'] = service[2]
        df.loc[row_index, 'Account VAT'] = service[3]
    print(df)
    return df

if __name__ == '__main__':
    files = glob.glob('*.pdf')

    columns = ['Date', 'Due date', 'Entry', 'Invoice Number', 'Total price','Service NR',
               'Order name', 'Price', 'Account', 'Account VAT', 'Text']

    df = pd.DataFrame(columns = columns)

    excel_rows = []
    row_index = 0
    for file in files:
        if 'Credit note' in file:
            try:
                date, total, invoice, due_date = get_regex(scrape_text(file))
                service_data, heading = get_table_data(file, 0)
            except Exception as e:
                print(e, f'\nFile {file} not parsed')
                df.loc[len(df.index)] = 'Not parsed:'
                df.loc[len(df.index)] = file
        else:
            try:
                date, total, invoice, due_date = get_regex(scrape_text(file))
                service_data, heading = get_table_data(file, 1)
                for i in service_data:
                    list_to_append = [date, due_date, file, invoice, total, i[0], i[1], i[2], i[3],i[4], f'Reform Supply & Logistics, UAB - {heading}']
                    df.loc[len(df.index)] = list_to_append
            except Exception as e:
                print(e,f'\nFile {file} not parsed')
                df.loc[len(df.index)] = 'Not parsed:'
                df.loc[len(df.index)] = file
    df.to_csv('output.csv', index=False)
    print('Have a nice day')
    time.sleep(1)