from openpyxl import load_workbook
from date import now_day, now_month, now_year
from sell import product_sell
from view import view_baza

wb = load_workbook("products_baza.xlsx")
sheet = wb.active

products = []
def add_func():
    wb.save(f"Add products_baza {now_day()}-{now_month()}-{now_year()}.xlsx")
    wb.save("products_baza.xlsx")
    
def add_product():
    product_name = input("Mahsulotni nomini kiriting: ")
    product_count = input("Mahsulot soni: ")
    quantity_type = input("Qanday sotiladi (kg, dona, litr): ")
    product_quantity = input("Bir donasini miqdori: ")
    product_price = input("Mahsulot narxi: ")
    product_ish_sanasi = input("Mahsulotni ishlab chiqarilgan sanasi: ")
    product_saqlash_muddati = input("Mahsulotni saqlash muddati: ")
    all_price = int(product_count) * int(product_price)
    #### User kiritgan mahsulotlarni dict ga olib qoshdik
    product_dict = {
        "name": product_name.title(),
        "count": product_count,
        "quantity": product_quantity,
        "quantity_type": quantity_type,
        "price": product_price,
        "ish_sanasi": product_ish_sanasi,
        "saqlash_muddati": product_saqlash_muddati,
        "all_price": all_price,
    }
    products.append(product_dict)
    
    #### Soradik yana qoshiladimi yoki yetadimi
    yana_qosh = input("Yana qoshamizmi (1/0): ")
    if yana_qosh == "1":
        add_product()

    #### User kiritgan mahsulotlarni excel ga yozdik
    number = 0
    id = sheet.max_row + 1
    print(id)
    for i in range(id, len(products)+id):
        sheet[f"A{i}"].value = id - 1
        sheet[f"B{i}"].value = products[number]["name"]
        sheet[f"C{i}"].value = products[number]["count"]
        sheet[f"D{i}"].value = products[number]["quantity"] + ' ' +  products[number]["quantity_type"]
        sheet[f"E{i}"].value = products[number]["price"]
        sheet[f"F{i}"].value = products[number]["ish_sanasi"]
        sheet[f"G{i}"].value = products[number]["saqlash_muddati"]
        sheet[f"H{i}"].value = products[number]["all_price"]
        number += 1

def choice():
    savol = input("Mahsulot qo'shish (1): \nMahsulot sotish (2): \nMahsulotlarni korish (3): ")
    if savol == "1":
        add_product()
        add_func()
    elif savol == "2":
        product_sell()
    elif savol == "3":
        view_baza()
        
choice()
