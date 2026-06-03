import tkinter as tk
from tkinter import messagebox, ttk
import openpyexcel as op
import os

# File 
file_name = "Guevarra_Database.xlsx"

if not os.path.exists(file_name):
    work = op.Workbook()
    sheet = work.active
    sheet['a1'] = "Product Name"
    sheet['b1'] = "Price"
    sheet['c1'] = "Brand"
    sheet['d1'] = "Size"
    sheet['e1'] = "Color"

    # raw data of product
    sheet['a2'] = "Nike Air Max Dn Roam"
    sheet['b2'] = "₱6,692"
    sheet['c2'] = "Nike"
    sheet['a3'] = "Dunk Low Retro Premium"
    sheet['b3'] = "₱4,897"
    sheet['c3'] = "Nike"
    sheet['a4'] = "Air Force 1 '07"
    sheet['b4'] = "₱5,317"
    sheet['c4'] = "Nike"
    sheet['a5'] = "Terra Manta Suede"
    sheet['b5'] = "₱3,357"
    sheet['c5'] = "Nike"
    sheet['a6'] = "Ultimashow Shoes"
    sheet['b6'] = "₱2,560"
    sheet['c6'] = "Adidas"
    sheet['a7'] = "COURTSHOT"
    sheet['b7'] = "₱2,610"
    sheet['c7'] = "Adidas"
    sheet['a8'] = "Harden V9"
    sheet['b8'] = "₱5,700"
    sheet['c8'] = "Adidas"
    sheet['a9'] = "Tokyo"
    sheet['b9'] = "₱4,950"
    sheet['c9'] = "Adidas"
    sheet['a10'] = "Racer TR23"
    sheet['b10'] = "₱3,430"
    sheet['c10'] = "Adidas"
    sheet['a11'] = "kai 1"
    sheet['b11'] = "₱5,596"
    sheet['c11'] = "ANTA"
    sheet['a12'] = "KAI 3"
    sheet['b12'] = "₱8,595"
    sheet['c12'] = "ANTA"
    sheet['a13'] = "PG7 3"
    sheet['b13'] = "₱4,595"
    sheet['c13'] = "ANTA"
    sheet['a14'] = "G21 4"
    sheet['b14'] = "₱5,995"
    sheet['c14'] = "ANTA"
    sheet['a15'] = "Shock Wave 7"
    sheet['b15'] = "₱5,596"
    sheet['c15'] = "ANTA"
    sheet['a16'] = "chuck 70"
    sheet['b16'] = "₱4,295"
    sheet['c16'] = "CONVERSE"
    sheet['a17'] = "old skool"
    sheet['b17'] = "₱3,198"
    sheet['c17'] = "VANS"
    sheet['a18'] = "BOOK 2"
    sheet['b18'] = "₱8,595"
    sheet['c18'] = "NIKE"
    sheet['a19'] = "GEL-KAYANO 14"
    sheet['b19'] = "₱9,590"
    sheet['c19'] = "ADIDAS"
    sheet['a20'] = "SK8-Low"
    sheet['b20'] = "₱2,470"
    sheet['c20'] = "VANS"


    # create sheet2
    sheet2 = work.create_sheet(title="Purchase")
    sheet2['a1'] = "ID"
    sheet2['b1'] = "Product Name"
    sheet2['c1'] = "Brand"
    sheet2['d1'] = "Price"
    sheet2['e1'] = "Size"

    work.save(file_name)


def display():
    wrkk = op.load_workbook("Guevarra_Database.xlsx")
    shet = wrkk.active
    
    for row in tree.get_children():
        tree.delete(row)
        
    for row in shet.iter_rows(min_row=2):
        row_values = [cell.value for cell in row]
        if any(row_values):
            tree.insert("", tk.END, values=row_values)

def select(event):
    selection = tree.focus()
    values = tree.item(selection, "values")

    if values:
        pro_entry.delete(0, tk.END)
        brand_entry.delete(0, tk.END)
        Price_entry.delete(0, tk.END)
        size_entry.delete(0, tk.END)

        pro_entry.insert(0, values[0])
        Price_entry.insert(0, values[1])
        brand_entry.insert(0, values[2])
        size_entry.insert(0, values[3])

def get_in():
    name = pro_entry.get()
    price = Price_entry.get()
    brand = brand_entry.get()
    size = size_entry.get()
    wb = op.load_workbook(file_name)

    sheet2 = wb["Purchase"]
    
    newId = sheet2.max_row 

    sheet2.append([newId, name, price, brand, size])
    wb.save(file_name)


def total():
    price_input = Price_entry.get()
    try:
      # para mawala yung sign at dot
        new_price = int(price_input.replace("₱", "").replace(",", "").split(".")[0]) 
        current_text = money["text"]
        clean_text = current_text.replace("₱", "").replace(",", "").split(".")[0]
        current_balance = int(clean_text)
            
        updated_total = current_balance + new_price
        money["text"] = f"₱{updated_total:,}.00"


        current_items = int(prod_count["text"]) 
        updated_items = current_items + 1
        prod_count["text"] = str(updated_items)
        
    except ValueError:
        money["text"] = "Enter valid price & quantity!"
    
    print(new_price)
    get_in()

def display_product():
    wrkk = op.load_workbook("Guevarra_Database.xlsx")
    shet = wrkk["Purchase"]
    
    for row in sales_tree.get_children():
        sales_tree.delete(row)
        
    for row in shet.iter_rows(min_row=2):
        row_values = [cell.value for cell in row]
        if any(row_values):
            sales_tree.insert("", tk.END, values=row_values)


    
def calculate_excel_total():
    wb = op.load_workbook(file_name)
    sheet2 = wb["Purchase"]
    
    total_price = 0
    total_items = 0

    for row in sheet2.iter_rows(min_row=2):
        row_values = [cell.value for cell in row]

        if any(row_values) and len(row_values) >= 4:
            price_val = row_values[2]  # to locate the price
            # print(price_val)
            if price_val is not None:
                try:
                    clean_price = int(price_val.replace("₱", "").replace(",", "").split(".")[0])
                    total_price += int(clean_price)
                    total_items += 1
                except ValueError:
                    pass  

    money1["text"] = f"₱{total_price:,}.00"
    prod_count1["text"] = str(total_items)


def sales():
    left_panel.grid_forget()
    right_panel.grid_forget()
    left_wing.grid(row=1, column=0, sticky="nw")
    right_wing.grid(row=1, column=1, sticky="nsew")
    menu.grid(row=0,column=0, sticky="nsew")
    display_product()
    calculate_excel_total()

    
def product():
    left_wing.grid_forget()
    right_wing.grid_forget()
    menu.grid_forget()
    left_panel.grid(row=0, column=0, sticky="nw")
    right_panel.grid(row=0, column=1, sticky="nsew")

def select_refund(event):
    selection = sales_tree.focus()
    values = sales_tree.item(selection, "values")

    if values:
        pro_entry1.delete(0, tk.END)
        brand_entry1.delete(0, tk.END)
        Price_entry1.delete(0, tk.END)
        size_entry1.delete(0, tk.END)

        pro_entry1.insert(0, values[1])
        Price_entry1.insert(0, values[2])
        brand_entry1.insert(0, values[3])
        size_entry1.insert(0, values[4])

def subtract():
    price_input1 = Price_entry1.get()
    # reund_input1 = refund1.get()
    try:
        
        new_price1 = int(price_input1.replace("₱", "").replace(",", "").split(".")[0])
        
        current_text1 = money1["text"]
        
        clean_text = current_text1.replace("₱", "").replace(",", "").split(".")[0]
        current_balance1 = int(clean_text)
        
        updated_total1 = current_balance1 - new_price1
        
        money1["text"] = f"₱{updated_total1:,}.00"

        current_items1 = int(prod_count1["text"]) 
        updated_items1 = max(0, current_items1 - 1) 
        prod_count1["text"] = str(updated_items1)

        current_refund1 = refund1["text"]
        clean_refund = current_refund1.replace("₱", "").replace(",", "").split(".")[0]
        goods_refund1 = int(clean_refund)
        updated_refund1 = new_price1 + goods_refund1
        refund1 ["text"] = f"₱{updated_refund1:,}.00"

    
        
    except ValueError:
        money1["text"] = "Enter valid price & quantity!"
        
    
    

    

def refund():
    selectedd = sales_tree.focus()

    if not selectedd:
        messagebox.showerror("Error","Select a Record first")
        return
    
    

    value = sales_tree.item(selectedd,"values")
    recordId = value[0]

    confirm = messagebox.askyesnocancel("Confirm","Are you sure you want to delete?")
    if not confirm:
        return

    workb = op.load_workbook("Guevarra_Database.xlsx")
    sheet2 = workb["Purchase"]
    for i,row in enumerate(sheet2.iter_rows(min_row=2),start=2):
        if int(row[0].value) == int(recordId):
            sheet2.delete_rows(i)
            break
    workb.save("Guevarra_Database.xlsx")

    messagebox.showinfo("Success","Record deleted successfully")
    
    subtract()
    pro_entry1.delete(0, tk.END)
    brand_entry1.delete(0, tk.END)
    Price_entry1.delete(0, tk.END)
    size_entry1.delete(0, tk.END)
    display_product()


window = tk.Tk()
window.title("Sales Monitoring System")
# window.geometry("850x450") 
window.resizable(True, True)

left_wing = tk.Frame(window, padx=15, pady=10)
right_wing = tk.Frame(window, padx=15, pady=10)

menu = tk.Frame(window, padx=5, pady=5)

# View Switcher Buttons
option = tk.Frame(menu)
option.grid(row=0, column=0, sticky="w", pady=(0, 5),padx=(1,20))
sell_btn = tk.Button(option, text="Products", width=10, command=product)
sales_btn = tk.Button(option, text="Sales List", width=10,command=sales)
sell_btn.grid(row=0, column=0, padx=(0, 5))
sales_btn.grid(row=0, column=1)

upper = tk.Frame(menu)
upper.grid(row=0, column=1, padx=30, sticky="w")

total_sales1 = tk.Frame(upper, bg="green", padx=15, pady=8)
total_sales1.grid(row=0, column=0, padx=(0, 10))
money_label1 = tk.Label(total_sales1, text="TOTAL PURCHASE", bg="green", fg="white", font=("Arial", 9, "bold"))
money_label1.grid(row=0, column=0)
money1 = tk.Label(total_sales1, text="₱0.00", bg="green", fg="white", font=("Arial", 11))
money1.grid(row=1, column=0)

product_sales1 = tk.Frame(upper, bg="green", padx=15, pady=8)
product_sales1.grid(row=0, column=1, padx=(0, 10))
prod_label1 = tk.Label(product_sales1, text="TOTAL ITEM", bg="green", fg="white", font=("Arial", 9, "bold"))
prod_label1.grid(row=0, column=0)
prod_count1 = tk.Label(product_sales1, text="0", bg="green", fg="white", font=("Arial", 11))
prod_count1.grid(row=1, column=0)

refund_sales1 = tk.Frame(upper, bg="green", padx=25, pady=8)
refund_sales1.grid(row=0, column=2)
refund_label1 = tk.Label(refund_sales1, text="REFUND", bg="green", fg="white", font=("Arial", 9, "bold"))
refund_label1.grid(row=0, column=0)
refund1 = tk.Label(refund_sales1, text="₱0.00", bg="green", fg="white", font=("Arial", 11))
refund1.grid(row=1, column=0)

# right




# Treeview 
cols = ("ID","Product Name", "Price", "Brand", "Size", "Quantity")
sales_tree = ttk.Treeview(left_wing, columns=cols, show="headings")

for col in cols:
    sales_tree.heading(col, text=col)
    sales_tree.column(col, width=110, anchor="center")
    
sales_tree.grid(row=2, column=0, sticky="nsew",padx=(1,10))
sales_tree.bind("<<TreeviewSelect>>", select_refund)


base = tk.Frame(left_wing, bd=1, relief="groove", padx=15, pady=15)
base.grid(row=2, column=1, sticky="w")

form1 = tk.Frame(base)
form1.grid(row=1, column=0, columnspan=2)


pro_name1 = tk.Label(form1, text="Product Name")
pro_name1.grid(row=0, column=0, sticky="w", pady=4, padx=(0, 5))
pro_entry1 = tk.Entry(form1, width=22)
pro_entry1.grid(row=0, column=1, pady=4)

Price1 = tk.Label(form1, text="Price")
Price1.grid(row=1, column=0, sticky="w", pady=4, padx=(0, 5))
Price_entry1 = tk.Entry(form1, width=22)
Price_entry1.grid(row=1, column=1, pady=4)

brand1 = tk.Label(form1, text="Brand")
brand1.grid(row=2, column=0, sticky="w", pady=4, padx=(0, 5))
brand_entry1 = tk.Entry(form1, width=22)
brand_entry1.grid(row=2, column=1, pady=4)

size1 = tk.Label(form1, text="Size")
size1.grid(row=3, column=0, sticky="w", pady=4, padx=(0, 5))
size_entry1 = tk.Entry(form1, width=22)
size_entry1.grid(row=3, column=1, pady=4)


btn_frame1 = tk.Frame(base)
btn_frame1.grid(row=2, column=0, columnspan=2, pady=(15, 0))

refund_btn1 = tk.Button(btn_frame1, text="Refund", width=9, command=refund)
# Update1 = tk.Button(btn_frame1, text="Update", width=9)
# delete1 = tk.Button(btn_frame1, text="Remove", width=9)

refund_btn1.grid(row=0, column=0, padx=2)
# Update1.grid(row=0, column=1, padx=2)
# delete1.grid(row=0, column=2, padx=2)



# ==================== LEFT PANEL (COLUMN 0) ====================
left_panel = tk.Frame(window, padx=15, pady=10)
left_panel.grid(row=0, column=0, sticky="nw")



option = tk.Frame(left_panel)
option.grid(row=0, column=0, sticky="w", pady=(0, 5))
sell_btn = tk.Button(option, text="Products", width=10)
sales_btn = tk.Button(option, text="Sales List", width=10,command=sales)
sell_btn.grid(row=0, column=0, padx=(0, 5))
sales_btn.grid(row=0, column=1)




base = tk.Frame(left_panel, bd=1, relief="groove", padx=15, pady=15)
base.grid(row=1, column=0, sticky="w")

title = tk.Label(base, text="Add / Edit Product", font=("Arial", 11, "bold"))
title.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")

form = tk.Frame(base)
form.grid(row=1, column=0, columnspan=2)

pro_name = tk.Label(form, text="Product Name")
pro_name.grid(row=0, column=0, sticky="w", pady=4, padx=(0, 5))
pro_entry = tk.Entry(form, width=22)
pro_entry.grid(row=0, column=1, pady=4)



Price = tk.Label(form, text="Price")
Price.grid(row=1, column=0, sticky="w", pady=4, padx=(0, 5))
Price_entry = tk.Entry(form, width=22)
Price_entry.grid(row=1, column=1, pady=4)

brand = tk.Label(form, text="Brand")
brand.grid(row=2, column=0, sticky="w", pady=4, padx=(0, 5))
brand_entry = tk.Entry(form, width=22)
brand_entry.grid(row=2, column=1, pady=4)

size = tk.Label(form, text="Size")
size.grid(row=3, column=0, sticky="w", pady=4, padx=(0, 5))
size_entry = tk.Entry(form, width=22)
size_entry.grid(row=3, column=1, pady=4)


btn_frame = tk.Frame(base)
btn_frame.grid(row=2, column=0, columnspan=2, pady=(15, 0))

add = tk.Button(btn_frame, text="Add Item", width=9, command=total)
# Update = tk.Button(btn_frame, text="Update", width=9)
# delete = tk.Button(btn_frame, text="Remove", width=9)

add.grid(row=0, column=0, padx=2)
# Update.grid(row=0, column=1, padx=2)
# delete.grid(row=0, column=2, padx=2)


# ==================== RIGHT PANEL (COLUMN 1) ====================
right_panel = tk.Frame(window, padx=15, pady=10)
right_panel.grid(row=0, column=1, sticky="nsew")


upper = tk.Frame(right_panel)
upper.grid(row=0, column=0, pady=(0, 15), sticky="w")

total_sales = tk.Frame(upper, bg="green", padx=15, pady=8)
total_sales.grid(row=0, column=0, padx=(0, 10))
money_label = tk.Label(total_sales, text="DAILY PURCHASE", bg="green", fg="white", font=("Arial", 9, "bold"))
money_label.grid(row=0, column=0)
money = tk.Label(total_sales, text="₱0.00", bg="green", fg="white", font=("Arial", 11))
money.grid(row=1, column=0)

product_sales = tk.Frame(upper, bg="green", padx=15, pady=8)
product_sales.grid(row=0, column=1)
prod_label = tk.Label(product_sales, text="DAILY ITEM", bg="green", fg="white", font=("Arial", 9, "bold"))
prod_label.grid(row=0, column=0)
prod_count = tk.Label(product_sales, text="0", bg="green", fg="white", font=("Arial", 11))
prod_count.grid(row=1, column=0)


info = tk.Label(right_panel, text="Select the item you want to buy:", font=("Arial", 9, "italic"), fg="gray30")
info.grid(row=1, column=0, sticky="w", pady=(5, 5))

cols = ("Product Name", "Price", "Brand", "Size", "Color")
tree = ttk.Treeview(right_panel, columns=cols, show="headings")

for col in cols:
    tree.heading(col, text=col)
    tree.column(col, width=110, anchor="center")
    
tree.grid(row=2, column=0, sticky="nsew")
tree.bind("<<TreeviewSelect>>", select)





display() 



window.mainloop()
