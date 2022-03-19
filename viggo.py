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
        print(e, 'Failure at scrape_text.')
        time.sleep(5)


def get_regex(text):
    global date, total, invoice, due_date
    date = ''
    cvr = ''
    amount = ''
    invoice = ''
    uge = ''
    sidste = False
    # from regex:
    # CVR:, dato, nr:(invoice), til udbetaling
    try:

        txt = text.split('\n')[0].split(' ')[:-3]

        for line in text.split('\n'):
            uge_row = re.compile(r'(Afregningsnota uge)')
            if uge_row.search(line):
                uge = line.split(' ')[-2]

            date_row = re.compile(r'(Dato)')
            if date_row.search(line):
                date_line = line.split(' ')[-1]
                date = date_line.replace('/','-')

            cvr_row = re.compile(r'(CVR:)')
            if cvr_row.search(line):
                cvr = line.split(' ')[1]

            overfort_sidste = re.compile(r'(Overført fra sidste afregningsnota)')
            if overfort_sidste.search(line):
                overfort_sidste_line = line.split(' ')[-2]
                sidste = round(float(overfort_sidste_line.replace(',', '')),2)

            udbetaling_row = re.compile(r'(Til udbetaling)')
            overfort_naeste = re.compile(r'(Overføres til næste afregningsnota)')
            if udbetaling_row.search(line):
                udbetaling = round(float(line.split(' ')[-2].replace(',','')),2)
                if sidste:
                    amount = udbetaling - sidste
                else:
                    amount = udbetaling
            elif overfort_naeste.search(line):
                naeste = round(float(line.split(' ')[-2].replace(',', '')), 2)
                if sidste:
                    amount = naeste - sidste
                else:
                    amount = naeste

            invoice_row = re.compile(r'(Nr:)')
            if invoice_row.search(line):
                invoice_line = line.split(' ')
                invoice = invoice_line[-1]
    except Exception as e:
        print(e, 'Regex failure')

    return uge, date, cvr, amount, invoice, ' '.join(txt)

if __name__ == '__main__':
    files = glob.glob('*.pdf')

    columns = ['Approval','Type','Date','Entry','Invoice','Text','Amount',
               '','Account','VAT','Contra account','VAT','Currency','Department']

    df = pd.DataFrame(columns=columns)

    excel_rows = []
    row_index = 0
    for file in files:
        uge, date, cvr, amount, invoice, txt = get_regex(scrape_text(file))

        date = '-'.join(date.split('-')[::-1])
        df.loc[row_index, 'Date'] = date
        df.loc[row_index, 'Invoice'] = invoice
        df.loc[row_index, 'Amount'] = amount
        df.loc[row_index, 'Account'] = cvr
        txt = txt + f' (week {uge})'
        df.loc[row_index, 'Text'] = txt
        row_index += 1

        print(f'File {file} done.')
    df.to_csv('output.csv', index=False)