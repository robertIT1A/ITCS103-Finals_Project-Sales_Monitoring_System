import tkinter as tk
from tkinter import messagebox, ttk
import openpyexcel as op
import os

# File para di na paulit ulit hirap itype eh
file_name = "Guevarra_Database.xlsx"

if not os.path.exists(file_name):
    work = op.Workbook()
    sheet = work.active
    sheet['a1'] = "Product Name"
    sheet['b1'] = "Price"
    sheet['c1'] = "Brand"
    sheet['d1'] = "Size"
    sheet['e1'] = "Color"

    # product record
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


    # create sheet2 para sa mga na bili na
    sheet2 = work.create_sheet(title="Purchase")
    sheet2['a1'] = "ID"
    sheet2['b1'] = "Product Name"
    sheet2['c1'] = "Brand"
    sheet2['d1'] = "Price"
    sheet2['e1'] = "Size"
    sheet2['f1'] = "Color"

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
        # size_radoibtn.delete(0, tk.END)

        pro_entry.insert(0, values[0])
        Price_entry.insert(0, values[1])
        brand_entry.insert(0, values[2])
        # size_radoibtn.insert(0, values[3])
        size_radoibtn.set(values[3] if values[3] is not None else "")
        color_radiobtn.set(values[4] if values[4] is not None else "")

def inputvalidation():
    name = pro_entry.get()
    price = Price_entry.get()
    brand = brand_entry.get()
    size = size_radoibtn.get()
    color = color_radiobtn.get()

    if not name or not price or not brand:
        messagebox.showerror("Input Error","All input must not be empty")
        return False
    
    if size == "None" and color == "None":
        confirm = messagebox.askyesnocancel("Confirm", "Are you sure you don't want to select BOTH size and color?")
        if not confirm: 
            return False
    elif size == "None":
        confirm = messagebox.askyesnocancel("Confirm", "Are you sure you don't want to select a SIZE?")
        if not confirm:
            return False
    elif color == "None":
        confirm = messagebox.askyesnocancel("Confirm", "Are you sure you don't want to select a COLOR?")
        if not confirm:
            return False   
    return True




def total():
    price_input = Price_entry.get()
    try:
        new_price = int(price_input.replace("₱", "").replace(",", "").split(".")[0])
        # para mawala yung sign at dot 
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
    
    # print(new_price)
    # inputvalidation()
def update():
    selected = tree.focus()

    if not selected:
        messagebox.showerror("Error","Select a Record first")
        return
    
    if not inputvalidation():
        return
    
    # values = tree.item(selected,"values")
    # recordId = values[0]

    # name = pro_entry.get()
    # price = Price_entry.get()
    # brand = brand_entry.get()
    # size = size_radoibtn.get()
    # color = color_radiobtn.get()

    # wrkbook = op.load_workbook("Guevarra_Database.xlsx")
    # act = wrkbook.active

    # for row in act.iter_rows(min_row=2):
    #     if int(row[0].value) == int(recordId):
    #         row[1].value = name
    #         row[2].value = price
    #         row[3].value = brand
    #         row[4].value = size
    #         row[5].value = color
    
    # wrkbook.save("Guevarra_Database.xlsx")

    # messagebox.showinfo("Success","Record updated successfully")
    values = tree.item(selected, "values")
    old_name = values[0]  

    new_name = pro_entry.get()
    new_price = Price_entry.get()
    new_brand = brand_entry.get()
    new_size = size_radoibtn.get()
    new_color = color_radiobtn.get()

    try:
        wrkbook = op.load_workbook(file_name)
        act = wrkbook.active

        record_found = False
        # check if match sa data
        for row in act.iter_rows(min_row=2):
            if row[0].value == old_name:
                row[0].value = new_name
                row[1].value = new_price
                row[2].value = new_brand
                row[3].value = new_size
                row[4].value = new_color
                record_found = True
                break
        
        if record_found:
            wrkbook.save(file_name)
            messagebox.showinfo("Success", "Record updated successfully in Database!")
            display()  
        else:
            messagebox.showerror("Error", "Could not find the original record matching that name.")

    except Exception as e:
        messagebox.showerror("Database Error", f"An error occurred: {str(e)}")

def get_in():
    if not inputvalidation():
        return False
    name = pro_entry.get()
    price = Price_entry.get()
    brand = brand_entry.get()
    size = size_radoibtn.get()
    color = color_radiobtn.get()

    wb = op.load_workbook(file_name)

    sheet2 = wb["Purchase"]
    
    newId = sheet2.max_row 

    sheet2.append([newId, name, price, brand, size, color])
    wb.save(file_name)
    messagebox.showinfo("Success","Record Purchased Successfully!")
    total()

def display_product():
    wrkk = op.load_workbook("Guevarra_Database.xlsx")
    shet = wrkk["Purchase"]
    brands_bought = []
    
    for row in sales_tree.get_children():
        sales_tree.delete(row)
        
    for row in shet.iter_rows(min_row=2):
        row_values = [cell.value for cell in row]
        
        if any(row_values):
            sales_tree.insert("", tk.END, values=row_values)

        # para sa most brand
        if len(row_values) >= 3:
            #  kukunin yung brand sa lahat ng nasa treeview
            brand_val = row_values[3] 
            if brand_val is not None:   
                brands_bought.append(str(brand_val).strip().upper())

    # print(brands_bought)
    # kukunin yung pinaka madami
    high_item1 = max(brands_bought, key=brands_bought.count)

    # print(f"The most repeated brand is: {high_item1}")
    top_brand1["text"] = high_item1


    
def calculate_excel_total():
    wb = op.load_workbook(file_name)
    sheet2 = wb["Purchase"]
    
    total_price = 0
    total_items = 0

    for row in sheet2.iter_rows(min_row=2):
        row_values = [cell.value for cell in row]

        if any(row_values) and len(row_values) >= 4:
            price_val = row_values[2]  # assign for price
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

# mga pantago para sa product and sales list
def sales():
    left_panel.grid_forget()
    right_panel.grid_forget()
    left_wing.grid(row=1, column=0, sticky="nw")
    right_wing.grid(row=1, column=1, sticky="nsew")
    menu.grid(row=0,column=0, sticky="nsew")
    display_product()
    calculate_excel_total()
    sales_btn1 ["fg"] = "white"
    sales_btn1 ["background"] = "#305CDE"


def product():
    left_wing.grid_forget()
    right_wing.grid_forget()
    menu.grid_forget()
    left_panel.grid(row=0, column=0, sticky="nw")
    right_panel.grid(row=0, column=1, sticky="nsew", pady=10)
    sell_btn ["fg"] = "white"
    sell_btn ["background"] = "#305CDE"

# select for refund sa sales list
def select_refund(event):
    selection = sales_tree.focus()
    values = sales_tree.item(selection, "values")

    if values:
        pro_entry1.delete(0, tk.END)
        brand_entry1.delete(0, tk.END)
        Price_entry1.delete(0, tk.END)
        size_entry1.delete(0, tk.END)
        color_entry1.delete(0, tk.END)

        pro_entry1.insert(0, values[1])
        Price_entry1.insert(0, values[2])
        brand_entry1.insert(0, values[3])
        size_entry1.insert(0, values[4])
        color_entry1.insert(0, values[5])

# def top_brand():

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
    color_entry1.delete(0, tk.END)
    display_product()


window = tk.Tk()
window.title("Sales Monitoring System")
window.resizable(False, False)
window.configure(background="#121214")
# ==================== para sa select list ====================
left_wing = tk.Frame(window, padx=15, pady=10, background="#121214")
right_wing = tk.Frame(window, padx=15, pady=10, background="#121214")

menu = tk.Frame(window, padx=5, pady=5, background="#121214")

option = tk.Frame(menu, background="#121214")
option.grid(row=0, column=0, sticky="w", pady=(0, 5),padx=(1,20))
sell_btn1 = tk.Button(option, text="Products",background="#D51A37", width=10, command=product,font=("Palatino Linotype",10,"bold"))
sales_btn1 = tk.Button(option, text="Sales List",background="#305CDE",fg="white", width=10,font=("Palatino Linotype",10,"bold"))
sell_btn1.grid(row=0, column=0, padx=(5, 5), ipadx=20, ipady=10)
sales_btn1.grid(row=0, column=1, ipadx=20, ipady=10)

high_brand1 = tk.Frame(option, bg="#FFB703",bd=0, relief="flat", padx=40, pady=15)
high_brand1.grid(row=1, column=0, columnspan=2, padx=40, pady=10, )
brand_label1 = tk.Label(high_brand1, text="TOP BRAND", bg="#FFB703", fg="black", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
top_brand1 = tk.Label(high_brand1, text="NONE", bg="#FFB703", fg="black", font=("Segoe UI", 15, "bold"))
top_brand1.grid(row=1, column=0, sticky="w")

upper = tk.Frame(menu, background="#121214")
upper.grid(row=0, column=1, padx=30, sticky="w")

total_sales1 = tk.Frame(upper, bg="#10B981", padx=40, pady=25, bd=0, relief="flat")
total_sales1.grid(row=0, column=0, padx=(0, 15))
money_label1 = tk.Label(total_sales1, text="TOTAL PURCHASE", bg="#10B981", fg="black", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
money1 = tk.Label(total_sales1, text="₱0.00", bg="#10B981", fg="black", font=("Segoe UI", 28, "bold"))
money1.grid(row=1, column=0, sticky="w")

product_sales1 = tk.Frame(upper, bg="#3B82F6", padx=40, pady=25, bd=0, relief="flat")
product_sales1.grid(row=0, column=1, padx=(0, 10))
prod_label1 = tk.Label(product_sales1, text="TOTAL ITEM", bg="#3B82F6", fg="white", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
prod_count1 = tk.Label(product_sales1, text="0", bg="#3B82F6", fg="white", font=("Segoe UI", 28, "bold"))
prod_count1.grid(row=1, column=0)

refund_sales1 = tk.Frame(upper, bg="#F63B3B", padx=40, pady=25, bd=0, relief="flat")
refund_sales1.grid(row=0, column=2)
refund_label1 = tk.Label(refund_sales1, text="REFUND", bg="#F63B3B", fg="black", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
refund1 = tk.Label(refund_sales1, text="₱0.00", bg="#F63B3B", fg="black", font=("Segoe UI", 28, "bold"))
refund1.grid(row=1, column=0)



sales_frame = tk.Frame(left_wing)
sales_frame.grid(row=2, column=0, sticky="nsew", padx=(1,10))
cols = ("ID","Product Name", "Price", "Brand", "Size", "Color")
sales_tree = ttk.Treeview(sales_frame, columns=cols, show="headings")
# para magkascroll
sales_scroll = ttk.Scrollbar(sales_frame, orient="vertical", command=sales_tree.yview)
sales_tree.configure(yscrollcommand=sales_scroll.set)
sales_tree.grid(row=0, column=0, sticky="nsew")
sales_scroll.grid(row=0, column=1, sticky="ns")
sales_frame.rowconfigure(0, weight=1)
sales_frame.columnconfigure(0, weight=1)

for col in cols:
    sales_tree.heading(col, text=col)
    sales_tree.column(col, width=110, anchor="center")
    
sales_tree.bind("<<TreeviewSelect>>", select_refund)


base = tk.Frame(left_wing, bd=1, relief="groove", padx=15, pady=15,background="#FFB703")
base.grid(row=2, column=1, sticky="w")

form1 = tk.Frame(base,background="#FFB703")
form1.grid(row=1, column=0, columnspan=2)

# Form Inputs Layout
pro_name1 = tk.Label(form1, text="Product Name",background="#FFB703",font=("Palatino Linotype",10))
pro_name1.grid(row=0, column=0, sticky="w", pady=4, padx=(0, 5))
pro_entry1 = tk.Entry(form1, width=22)
pro_entry1.grid(row=1, column=0, columnspan=2, pady=4,ipadx=50)

Price1 = tk.Label(form1, text="Price",background="#FFB703",font=("Palatino Linotype",10))
Price1.grid(row=2, column=0, sticky="w", pady=4, padx=(0, 5))
Price_entry1 = tk.Entry(form1, width=22)
Price_entry1.grid(row=3, column=0, columnspan=2, pady=4,ipadx=30)

brand1 = tk.Label(form1, text="Brand",background="#FFB703",font=("Palatino Linotype",10))
brand1.grid(row=4, column=0, sticky="w", pady=4, padx=(0, 5))
brand_entry1 = tk.Entry(form1, width=22)
brand_entry1.grid(row=5, column=0, columnspan=2, pady=4,ipadx=25)

size1 = tk.Label(form1, text="Size",background="#FFB703",font=("Palatino Linotype",10))
size1.grid(row=6, column=0, sticky="w", pady=4, padx=(0, 5))
size_entry1 = tk.Entry(form1, width=22)
size_entry1.grid(row=7, column=0, columnspan=2, pady=4,ipadx=15)

color10 = tk.Label(form1, text="Color",background="#FFB703",font=("Palatino Linotype",10))
color10.grid(row=8, column=0, sticky="w", pady=4, padx=(0, 5))
color_entry1 = tk.Entry(form1, width=22)
color_entry1.grid(row=9, column=0, columnspan=2, pady=4,ipadx=10)

btn_frame1 = tk.Frame(base, background="#FFB703")
btn_frame1.grid(row=2, column=0, columnspan=2, pady=(15, 0))

refund_btn1 = tk.Button(btn_frame1, text="Refund", width=9, command=refund, background="#FFB703", activebackground="#D49B0B", activeforeground="#000000", font=("Palatino Linotype",10))
# Update1 = tk.Button(btn_frame1, text="Update", width=9)
# delete1 = tk.Button(btn_frame1, text="Remove", width=9)

refund_btn1.grid(row=0, column=0, padx=2)
# Update1.grid(row=0, column=1, padx=2)
# delete1.grid(row=0, column=2, padx=2)



# ==================== Product ====================
left_panel = tk.Frame(window, padx=15, pady=10, background="#121214")
left_panel.grid(row=0, column=0, sticky="nw")

option = tk.Frame(left_panel,background="#121214")
option.grid(row=0, column=0, sticky="w", pady=(0, 5))
sell_btn = tk.Button(option, text="Products",background="#305CDE",fg="white",font=("Palatino Linotype",10,"bold"),width=10)
sales_btn = tk.Button(option, text="Sales List",background="#D51A37",command=sales,font=("Palatino Linotype",10,"bold"),width=10)
sell_btn.grid(row=0, column=0, padx=(5, 5), ipadx=20, ipady=10)
sales_btn.grid(row=0, column=1, ipadx=20, ipady=10)

base = tk.Frame(left_panel, bd=1, relief="groove", padx=15, pady=15,background="#FFB703")
base.grid(row=1, column=0, sticky="w")

title = tk.Label(base, text="Add / Edit Product", font=("Bookman Old Style", 15, "bold"),background="#FFB703")
title.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")

form = tk.Frame(base,background="#FFB703")
form.grid(row=1, column=0, columnspan=2)


pro_name = tk.Label(form, text="Product Name",background="#FFB703",font=("Palatino Linotype",10))
pro_name.grid(row=0, column=0, sticky="w", pady=4, padx=(0, 5))
pro_entry = tk.Entry(form, width=22)
pro_entry.grid(row=1, column=0, columnspan=2, pady=4,ipadx=50)

Price = tk.Label(form, text="Price",background="#FFB703",font=("Palatino Linotype",10))
Price.grid(row=2, column=0, sticky="w", pady=4, padx=(0, 5))
Price_entry = tk.Entry(form, width=22)
Price_entry.grid(row=3, column=0, columnspan=2, pady=4,ipadx=30)

brand = tk.Label(form, text="Brand",background="#FFB703",font=("Palatino Linotype",10))
brand.grid(row=4, column=0, sticky="w", pady=4, padx=(0, 5))
brand_entry = tk.Entry(form, width=22)
brand_entry.grid(row=5, column=0, columnspan=2, pady=4)

size = tk.Label(form, text="Size",background="#FFB703",font=("Palatino Linotype",10))
size.grid(row=6, column=0, sticky="w", pady=4, padx=(0, 5))
# size_entry = tk.Entry(form, width=22)
# size_entry.grid(row=3, column=1, pady=4)
sizes_option = tk.Frame(form, bd=3, relief="ridge",padx=30)
sizes_option.grid(row=7, column=0, columnspan=2, pady=4)
size_radoibtn = tk.StringVar(window)
radio1 = tk.Radiobutton(sizes_option,text="3.5",variable=size_radoibtn,value="3.5",font="Garamond")
radio1.grid(row=0, column=0,padx=(0,10))
radio2 = tk.Radiobutton(sizes_option,text="4",variable=size_radoibtn,value="4",font="Garamond")
radio2.grid(row=0, column=1,padx=(0,10))
radio3 = tk.Radiobutton(sizes_option,text="4.5",variable=size_radoibtn,value="4.5",font="Garamond")
radio3.grid(row=0, column=2)
radio4 = tk.Radiobutton(sizes_option,text="5",variable=size_radoibtn,value="5",font="Garamond")
radio4.grid(row=1, column=0,padx=(0,20))
radio5 = tk.Radiobutton(sizes_option,text="5.5",variable=size_radoibtn,value="5.5",font="Garamond")
radio5.grid(row=1, column=1,padx=(0,3))
radio6 = tk.Radiobutton(sizes_option,text="6",variable=size_radoibtn,value="6",font="Garamond")
radio6.grid(row=1, column=2,padx=(0,10))
radio7 = tk.Radiobutton(sizes_option,text="6.5",variable=size_radoibtn,value="6.5",font="Garamond")
radio7.grid(row=2, column=0,padx=(0,10))
radio8 = tk.Radiobutton(sizes_option,text="7",variable=size_radoibtn,value="7",font="Garamond")
radio8.grid(row=2, column=1,padx=(0,10))
radio9 = tk.Radiobutton(sizes_option,text="7.5",variable=size_radoibtn,value="7.5",font="Garamond")
radio9.grid(row=2, column=2)


color = tk.Label(form, text="Color",background="#FFB703",font=("Palatino Linotype",10))
color.grid(row=8, column=0, sticky="w", pady=4, padx=(0, 5))
color_option = tk.Frame(form, bd=3, relief="raised")
color_option.grid(row=9, column=0,columnspan=2, pady=3)
color_radiobtn = tk.StringVar(window)
color1 = tk.Radiobutton(color_option,text="White",variable=color_radiobtn,value="White",font="Perpetua")
color1.grid(row=0, column=0)
color2 = tk.Radiobutton(color_option,text="Black",variable=color_radiobtn,value="Black",font="Perpetua")
color2.grid(row=0, column=1)
color3 = tk.Radiobutton(color_option,text="Blue",variable=color_radiobtn,value="Blue",font="Perpetua")
color3.grid(row=1, column=0)
color4 = tk.Radiobutton(color_option,text="Red",variable=color_radiobtn,value="Red",font="Perpetua")
color4.grid(row=1, column=1)
color5 = tk.Radiobutton(color_option,text="Brown",variable=color_radiobtn,value="Brown",font="Perpetua")
color5.grid(row=2, column=0)
color6 = tk.Radiobutton(color_option,text="Green",variable=color_radiobtn,value="Green",font="Perpetua")
color6.grid(row=2, column=1)


btn_frame = tk.Frame(base,background="#FFB703")
btn_frame.grid(row=2, column=0, columnspan=2, pady=(15, 0))

add = tk.Button(btn_frame, text="Add Item", background="#FFB703", activebackground="#D49B0B", activeforeground="#000000",width=9, command=get_in,font=("Palatino Linotype",10))
Update = tk.Button(btn_frame, text="Edit", background="#FFB703", activebackground="#D49B0B", activeforeground="#000000",width=9, command=update,font=("Palatino Linotype",10))
# delete = tk.Button(btn_frame, text="Remove", width=9)

add.grid(row=0, column=0, padx=2)
Update.grid(row=0, column=1, padx=2)
# delete.grid(row=0, column=2, padx=2)



right_panel = tk.Frame(window, bg="#121214")
right_panel.grid(row=0, column=1, sticky="nsew", padx=15, pady=10)
right_panel.columnconfigure(0, weight=1)
right_panel.rowconfigure(2, weight=1)


upper = tk.Frame(right_panel, bg="#121214")
upper.grid(row=0, column=0, pady=(0, 15), sticky="ew") # Changed sticky to "ew" to stretch horizontally


upper.columnconfigure(0, weight=1)
upper.columnconfigure(1, weight=1)

total_sales = tk.Frame(upper, bg="#10B981", padx=40, pady=25, bd=0, relief="flat")
total_sales.grid(row=0, column=0, padx=(0, 15), sticky="nsew") # sticky="nsew" fills the weighted column
total_sales.columnconfigure(0, weight=1)

money_label = tk.Label(total_sales, text="DAILY PURCHASE", bg="#10B981", fg="black", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
money = tk.Label(total_sales, text="₱0.00", bg="#10B981", fg="black", font=("Segoe UI", 28, "bold")) 
money.grid(row=1, column=0, sticky="w")


product_sales = tk.Frame(upper, bg="#3B82F6", padx=40, pady=25, bd=0, relief="flat")
product_sales.grid(row=0, column=1, sticky="nsew") # sticky="nsew" fills the weighted column
product_sales.columnconfigure(0, weight=1)

prod_label = tk.Label(product_sales, text="DAILY ITEM", bg="#3B82F6", fg="white", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
prod_count = tk.Label(product_sales, text="0", bg="#3B82F6", fg="white", font=("Segoe UI", 28, "bold"))
prod_count.grid(row=1, column=0, sticky="w")

info = tk.Label(right_panel, text="Select the item you want to buy:", font=("Arial", 10, "italic"), fg="#CFCFCE",background="#121214")
info.grid(row=1, column=0, sticky="w", pady=(5, 5))

tree_frame = tk.Frame(right_panel)
tree_frame.grid(row=2, column=0, sticky="nsew")

cols = ("Product Name", "Price", "Brand", "Size", "Color")
tree = ttk.Treeview(tree_frame, columns=cols, show="headings")

tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=tree_scroll.set)

tree.grid(row=0, column=0, sticky="nsew")
tree_scroll.grid(row=0, column=1, sticky="ns")

tree_frame.rowconfigure(0, weight=1)
tree_frame.columnconfigure(0, weight=1)

# para maging maganda yung table
style = ttk.Style()
style.theme_use('alt')
# para sa size
style.configure("Treeview", rowheight=33.5)
for col in cols:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor="center")
    
tree.bind("<<TreeviewSelect>>", select)


display() 



window.mainloop()
