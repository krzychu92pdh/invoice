from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
import datetime
import csv
import generator
import webbrowser
import subprocess
import os
import sys
import shutil
import config

# os.chdir(os.path.dirname(sys.argv[0]))


def get_invoice_month_from_user():
    months = ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień",
              "październik", "listopad", "grudzień"]

    while True:
        month_of_invoice = input("Podaj miesiąc w którym wykonana była usługa: ")
        month_of_invoice = month_of_invoice.lower()
        if month_of_invoice in months:
            break
        else:
            print("Nie ma takiego miesiąca spróbuj jeszcze raz:")


def get_invoice_amount_from_user():
    while True:
        amount_of_invoice = input("Podaj kwotę pieniędzy na jaką ma być wystawiona faktura: ")
        amount_of_invoice = amount_of_invoice.replace(",", ".")
        try:
            float(amount_of_invoice)
        except ValueError:
            continue
        break


def generate_invoice_document():
    invoice_document = Document('template_inv.docx')
    print(data)
    findandreplace("Faktura nr FV", "" + data + "/" + lastdayofmonth2)
    findandreplace("Data wystawienia:", lastdayofmonth)
    findandreplace("Miejsce wystawienia", config.place)

    invoice_document.tables[0].cell(1, 0).text = config.name_1
    invoice_document.tables[0].cell(2, 0).text = config.adress_1
    invoice_document.tables[0].cell(3, 0).text = config.additional_adress_1
    invoice_document.tables[0].cell(4, 0).text = "NIP:" + config.nip_1
    invoice_document.tables[0].cell(5, 0).text = "E-mail:" + config.mail_1
    invoice_document.tables[0].cell(6, 0).text = "Tel.:" + config.tel_1
    invoice_document.tables[0].cell(1, 0).paragraphs[0].runs[0].font.bold = True

    invoice_document.tables[0].cell(1, 1).text = config.name_2
    invoice_document.tables[0].cell(2, 1).text = config.adress_2
    invoice_document.tables[0].cell(3, 1).text = config.additional_adress_2
    invoice_document.tables[0].cell(4, 1).text = "NIP:" + config.nip_2
    invoice_document.tables[0].cell(1, 1).paragraphs[0].runs[0].font.bold = True

    invoice_document.tables[1].cell(1, 1).text = name
    invoice_document.tables[1].cell(1, 5).text = netto
    invoice_document.tables[1].cell(1, 7).text = netto
    invoice_document.tables[1].cell(1, 9).text = netto
    invoice_document.tables[1].cell(1, 5).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    invoice_document.tables[1].cell(1, 7).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    invoice_document.tables[1].cell(1, 9).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

    invoice_document.tables[2].cell(1, 1).text = netto
    invoice_document.tables[2].cell(1, 3).text = netto
    invoice_document.tables[2].cell(2, 1).text = netto
    invoice_document.tables[2].cell(2, 3).text = netto
    invoice_document.tables[2].cell(1, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    invoice_document.tables[2].cell(1, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    invoice_document.tables[2].cell(2, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    invoice_document.tables[2].cell(2, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

    invoice_document.tables[3].cell(0, 3).text = value + " PLN"
    invoice_document.tables[3].cell(2, 3).text = value + " PLN"
    invoice_document.tables[3].cell(3, 3).text = word_value
    invoice_document.tables[3].cell(1, 1).text = termofpayment
    invoice_document.tables[3].cell(2, 1).text = config.bank_name
    invoice_document.tables[3].cell(3, 1).text = config.bank_account

    invoice_document.tables[4].cell(0, 0).text = config.name_of_issuing
    invoice_document.tables[4].cell(0, 0).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    invoice_document.tables[4].cell(0, 0).paragraphs[0].runs[0].font.bold = True


def open_preview():
    pass


def confirm_invoice():
    pass


def save_files():
    pass


def delete_files():
    pass


########################################################################################

monthnr = months.index(month) + 1

today = datetime.date.today()
year = today.strftime("%Y")
if monthnr != 12:
    lastday = today.replace(day=1, month=int(monthnr+1), year=int(year)) - datetime.timedelta(days=1)
elif monthnr == 12:
    lastday = today.replace(day=1, month=1, year=int(year)+1) - datetime.timedelta(days=1)
lastdayofmonth = lastday.strftime("%d/%m/%Y")
lastdayofmonth2 = lastday.strftime("%m/%Y")
term = lastday + datetime.timedelta(days=14)
termofpayment = term.strftime("%d/%m/%Y")

name = "usługi medyczne wykonane w miesiącu - "+ month + " "+year+""
netto = value

with open('inv_nr.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    list = list(reader)
    for i in list:
        if i.get("month") == str(monthnr):
            y = int(i['nr_inv'])
            i['nr_inv'] = str(y+1)
            data = i['nr_inv']

with open('inv_nr.csv', "w", newline='') as csvfile:
    writer = csv.DictWriter(csvfile, ['month', 'nr_inv'])
    writer.writeheader()
    writer.writerows(list)

word_value = str(generator.generator(value))
#############################################################################

def findandreplace(find, replace):
    for paragraph in document.paragraphs:
        if find in paragraph.text:
            paragraph.add_run(replace).bold = True


document = Document('template_inv.docx')
print(data)
findandreplace("Faktura nr FV", ""+data+"/" + lastdayofmonth2)
findandreplace("Data wystawienia:", lastdayofmonth)
findandreplace("Miejsce wystawienia", config.place)

###################### TABLE1 #####################
document.tables[0].cell(1, 0).text = config.name_1
document.tables[0].cell(2, 0).text = config.adress_1
document.tables[0].cell(3, 0).text = config.additional_adress_1
document.tables[0].cell(4, 0).text = "NIP:" + config.nip_1
document.tables[0].cell(5, 0).text = "E-mail:" + config.mail_1
document.tables[0].cell(6, 0).text = "Tel.:" + config.tel_1
document.tables[0].cell(1, 0).paragraphs[0].runs[0].font.bold = True

document.tables[0].cell(1, 1).text = config.name_2
document.tables[0].cell(2, 1).text = config.adress_2
document.tables[0].cell(3, 1).text = config.additional_adress_2
document.tables[0].cell(4, 1).text = "NIP:" + config. nip_2
document.tables[0].cell(1, 1).paragraphs[0].runs[0].font.bold = True


###################### TABLE2 #####################
document.tables[1].cell(1, 1).text = name
document.tables[1].cell(1, 5).text = netto
document.tables[1].cell(1, 7).text = netto
document.tables[1].cell(1, 9).text = netto
document.tables[1].cell(1, 5).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
document.tables[1].cell(1, 7).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
document.tables[1].cell(1, 9).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

###################### TABLE3 #####################
document.tables[2].cell(1, 1).text = netto
document.tables[2].cell(1, 3).text = netto
document.tables[2].cell(2, 1).text = netto
document.tables[2].cell(2, 3).text = netto
document.tables[2].cell(1, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
document.tables[2].cell(1, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
document.tables[2].cell(2, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
document.tables[2].cell(2, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

document.tables[3].cell(0, 3).text = value+" PLN"
document.tables[3].cell(2, 3).text = value+" PLN"
document.tables[3].cell(3, 3).text = word_value
document.tables[3].cell(1 ,1).text = termofpayment
document.tables[3].cell(2 ,1).text = config.bank_name
document.tables[3].cell(3 ,1).text = config.bank_account


document.tables[4].cell(0 ,0).text = config.name_of_issuing
document.tables[4].cell(0, 0).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
document.tables[4].cell(0, 0).paragraphs[0].runs[0].font.bold = True

nameoffile = "faktura rehasport nr "+data+" - "+ month +" "+ year
document.save(nameoffile+".docx")
output = subprocess.check_output(['libreoffice', '--convert-to', 'pdf', nameoffile+".docx"])
webbrowser.open(nameoffile+".pdf", new=1)


dest_fpath = config.path

while True:
    approve = input("Czy faktura jest ok? y/n: ")
    approve = approve.lower()
    if approve == "y":
        os.makedirs(os.path.dirname(dest_fpath+month+"/"), exist_ok=True)
        shutil.move(nameoffile+".docx", dest_fpath+month+"/"+nameoffile+".docx")
        shutil.move(nameoffile+".pdf", dest_fpath+month+"/"+nameoffile+".pdf")
        break
    elif approve == "n":
        while True:
            x = input("Wybierz opcję: \ndel - usuń plik \nedit - zostaw plik .docx: ")
            if x == "del":
                os.remove(nameoffile+".docx")
                os.remove(nameoffile+".pdf")
                print("Usunięto wszystkie pliki")
                break
            if x == "edit":
                os.remove(nameoffile+".pdf")
                print("Popraw błędy w pliku .docx")
                break
            else:
                continue
        break
    else:
        continue
