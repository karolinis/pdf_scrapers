import re, pdfplumber, csv, glob, xlsxwriter

def scraper(pdfName):
    with pdfplumber.open(pdfName) as file: #open file
        all_txt = ''
        for page in file.pages: #extract text from every page
            txt = page.extract_text()
            all_txt = all_txt + '\n' + txt
        return all_txt #returning txt here will result in all page parse

def filter(text):
    data = []
    try:
        for nr,line in enumerate(text.split('\n')):
            # finds invoice dates:
            d_line = line.split(' ')
            pNumber_row = re.compile(r'(Project number)')
            if pNumber_row.search(line):
                data.append(d_line[-1]) #project number

            partnerName_row = re.compile(r'(Partner full name)')
            if partnerName_row.search(line):
                full_name = ' '.join(d_line[4:d_line.index('Events:')])
                data.append(full_name.strip())  # project number

            InvoiceProposalNr_row = re.compile(r'(Invoice Proposal Nr)')
            if InvoiceProposalNr_row.search(line):
                prop_nr = '-'+d_line[6]
                data.append(d_line[5]+prop_nr) #invoice proposal nr

            partnerInvNr_row = re.compile(r'(Partner invoice number:)')
            if partnerInvNr_row.search(line):
                data.append(d_line[4:]) #partner invoice number

            service_line = re.compile(r'(59\d\d)')
            if service_line.search(line):
                #ax glcd, sp code, service price in one list to be able to visualise easy on excel
                industry = ['MD', 'OT', 'CE', 'HA', 'DH', 'PS', 'UM']
                if 'USD' in d_line:
                    if d_line[2] in industry:
                        data.append([d_line[0],'DUM000',d_line[1],d_line[2],d_line[d_line.index('USD')-1],'','USD'])
                    else:
                        data.append([d_line[0],'DUM000', d_line[1], '', d_line[d_line.index('USD') - 1], '', 'USD'])
                elif 'TRM' in d_line:
                    if d_line[2] in industry:
                        data.append([d_line[0],'DUM000',d_line[1],d_line[2],d_line[d_line.index('TRM')-1],'','TRM'])
                    else:
                        data.append([d_line[0],'DUM000', d_line[1], '', d_line[d_line.index('TRM') - 1], '', 'TRM'])
                else:
                    if d_line[2] in industry:
                        data.append([d_line[0],'DUM000',d_line[1],d_line[2],d_line[d_line.index('EUR')-1],'','EUR'])
                    else:
                        data.append([d_line[0],'DUM000', d_line[1], '', d_line[d_line.index('EUR') - 1], '', 'EUR'])

            CNV_row = re.compile(r'(Currency Net Value)')
            if CNV_row.search(line):
                data.append(d_line[-2])  # partner invoice number

    except Exception as ex:
        print(ex)

    return data

if __name__ == '__main__':
    invoices = glob.glob('IP*.pdf')
    excel_rows = []
    for invoice in invoices:
        filtered_data = filter(scraper(invoice))
        if len(filtered_data) != 0:
            services = filtered_data[4:-1]
            ledger = ['Ledger']
            net_value = ['Vendor','','','','','',filtered_data[-1],filtered_data[-2][-1]]
            partner_inv = filtered_data[3]

            for service in services:
                excel_rows.append(filtered_data[0:3]+ledger+service+partner_inv)
            excel_rows.append(filtered_data[0:3]+net_value+partner_inv)

    columns = ['Project number','Partner name','Invoice proposal number','Account type','AX GLCD','Location','SP CODE','Industry','Debit','Credit','Currency','Partner invoice number']
    with open('IP_output.csv', 'w', encoding='UTF-8', newline='') as o:
        write = csv.writer(o)
        write.writerow(columns)
        write.writerows(excel_rows)