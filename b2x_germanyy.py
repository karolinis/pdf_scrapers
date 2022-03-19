import re, pdfplumber, csv, glob
import pandas as pd
import time

def scraper(pdfName):
    with pdfplumber.open(pdfName) as file: #open file
        all_txt = ''
        for page in file.pages: #extract text from every page
            txt = page.extract_text()
            all_txt = all_txt + '\n' + txt
        return all_txt #returning txt here will result in all page parse

def filter(text):
    data = [] #main list of elements

    for line in text.split('\n'):

        # finds invoices that start with:
        invoice_no_row = re.compile(r'(Invoice Number:|Rechnungsnummer:)')
        if invoice_no_row.search(line):
            # find invoice number
            invoice_line = line.split(' ')
            invoice = invoice_line[-1]
            if invoice_line[-1] not in data and invoice_line[-2] not in data:
                data.append(invoice)


        date_row = re.compile(r'(Invoice Date:|Rechnungsdatum:)')
        if date_row.search(line):
            date_line = line.split(' ')
            date = date_line[-1]
            if date_line[-1] not in data and date_line[-2] not in data:
                date = clean_date(date)
                if date not in data:
                    data.append(date)

        total_duedate_row = re.compile(r'(Fälligkeitsdatum:|Payment due date:)')
        if total_duedate_row.search(line):
            total_duedate_line = line.split(' ')
            if len(total_duedate_line) == 10: #this is due to document being in english; other lenght is 7
                due_date = clean_date(total_duedate_line[3])
                data.append(due_date) #due date
                total_credit = clean_numbers(total_duedate_line[-1])
                data.append(total_credit) #total
                #reikia prideti empty lines for correct output

            else:
                due_date = clean_date(total_duedate_line[1])
                data.append(due_date)
                total_credit = clean_numbers(total_duedate_line[-1])
                data.append(total_credit)  # total
                #same here


        service_row = re.compile(r'(Produkt Zwischensumme|Service Sub Total)')
        if service_row.search(line):
            service_line = line.split(' ')
            credit = clean_numbers(service_line[-1])
            data.append(['',data[1],'','','Ledger','','','',credit,'',data[0],'Ledger','','','','','','','EUR','','',data[2],'','','','','','DUM000','MD','PHS',''])

    list_format = ['',data[1],'','','Vendor','','','','',data[3],data[0],'Ledger','','','','','','','EUR','','',data[2],'','','','','','DUM000','MD','PHS','',data[4:]]

    yield list_format

def clean_date(date_line):
    if '-' in date_line:
        split_date = date_line.split('-')
        date = split_date[1], split_date[0], split_date[2]
        date = '-'.join(date)
    else:
        split_date = date_line.split('.')
        date = split_date[1], split_date[0], split_date[2]
        date = '-'.join(date)
    return date

def clean_numbers(number):
    original = number.split('.')
    number_list = ' '.join(original)
    commas = re.sub(',', '.', number_list)
    final = re.sub(' ', '', commas)
    return float(final)

def working_file_parse(doc):
    with open(doc, newline='') as file:
        read = csv.reader(file, delimiter=',')
        tmp = []
        for row in read:
            tmp.append(row)
        return tmp

def add_name_working_file(file):
    project_file = []
    for nr,line in enumerate(file[1:]):
        if line[0] == '':
            line[0] = file[nr-1][0]
            project_file.append(line)
    return project_file



if __name__ == '__main__':
    start = time.time()
    invoices = glob.glob('*.pdf')
    excel_rows = []
    project_file = working_file_parse('working_file_b2xg.csv')
    project_file = add_name_working_file(project_file)

    columns = ['Note', 'Date','Document Date','Voucher','Account type','Account','Enumerated Text', 'Text'
                ,'Debit', 'Credit', 'Invoice','Offset account type','Offset account','Ref to physical voucher'
                ,'Physical voucher','HasOffsetAccounts','VAT (Account)','VAT (Offset account)','Currency'
                ,'Debit(Currency)','Credit(Currency)', 'Due Date', 'Settlement type','Settlements','Force settlement'
                ,'Project','Cost center','Location','Industry','Services provid','On hold']


    DD = glob.glob('*DD.pdf')

    global prj_name
    for invoice in invoices:
        data = next(filter(scraper(invoice)))
        for row in data[-1]:
            excel_rows.append(row)
        excel_rows.append(data[:-1])



    for nr,row in enumerate(excel_rows):
        for project in project_file[1:]:
            prj_name_org = project[1]
            prj_name = re.sub('/RC95', '', prj_name_org)
            prj_name = re.sub('/RC02', '', prj_name)
            prj_name = re.sub('Total', '', prj_name)
            if row[10] == project[0] and excel_rows[nr][-6] == '' and 'Total' in prj_name_org and excel_rows[nr-1][-6] != prj_name:
                print(project[5])
                if excel_rows[nr][4] == 'Ledger':
                    excel_rows[nr][8] = project[5]
                    excel_rows[nr][-6] = prj_name
                else:
                    excel_rows[nr][-6] = prj_name
                break

    with open('output.csv', 'w', encoding='UTF-8', newline='') as o:
        write = csv.writer(o)
        write.writerow(columns)
        write.writerows(excel_rows)

    end = time.time()
    print(end - start)
