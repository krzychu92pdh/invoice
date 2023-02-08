from functions import *

set_data_of_invoice()
insert_seller_details()
insert_buyer_details()
add_item_to_invoice(1)

n = 1
while True:
    print("\n x - wyjście \n 1 - dodaj następną pozycję")
    your_choice = input("Podaj wartość: ")
    if your_choice != "x":
        n += 1
        add_item_to_invoice(n)

    else:
        break

insert_total_amount()
payment_information()
save_as_docx()
convert_to_pdf()

