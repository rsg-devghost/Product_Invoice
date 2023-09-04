import tkinter
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
import datetime


# Root - Widget/ Outer / Parent Frame
window = tkinter.Tk()
window.title("Invoice Generation")


def clear_row():
    quantity_entry.delete(0, tkinter.END)
    quantity_entry.insert(0, "0")
    description_entry.delete(0, tkinter.END)
    price_entry.delete(0, tkinter.END)
    price_entry.insert(0, "0.0")


def error_message():
    messagebox.showerror('Logical Error', "Cannot enter Zero Items!! Please enter at least one item.")


invoice_list = []


def add_item():
    # In the if condition check for other non-entry values as well!!
    if int(quantity_entry.get()) != 0:
        qty = int(quantity_entry.get())
        desc = description_entry.get()
        price = float(price_entry.get())
        total = qty * price
        items_list = [qty, desc, price, total]
        tree.insert('', 0, values=items_list)
        clear_row()
        invoice_list.append(items_list)

    else:
        error_message()


def new_invoice():
    fname_entry.delete(0, tkinter.END)
    lname_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_row()
    tree.delete(*tree.get_children())
    invoice_list.clear()


def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = fname_entry.get()+''+lname_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal + (salestax * subtotal)

    doc.render({"name": name,
                "phone": phone,
                "invoice_list": invoice_list,
                "subtotal": subtotal,
                "salestax": str(salestax*100)+'%',
                "total": total})
    # For some unknown reason u cannot add ":" below name of doc.
    doc_name = "customer name" + name + ".docx"
    doc.save(doc_name)


# The widget inside window
frame = tkinter.Frame(window)
# this function helps in positioning the widgets
frame.pack(padx=20, pady=20)

fname_label = tkinter.Label(frame, text="First Name")
fname_label.grid(row=0, column=0)
lname_label = tkinter.Label(frame, text="Last Name")
lname_label.grid(row=0, column=1)
phone_label = tkinter.Label(frame, text="Phone No")
phone_label.grid(row=0, column=2)
quantity_label = tkinter.Label(frame, text="Quantity")
quantity_label.grid(row=2, column=0)
description_label = tkinter.Label(frame, text="Description")
description_label.grid(row=2, column=1)
price_label = tkinter.Label(frame, text="Unit Price")
price_label.grid(row=2, column=2)
add_item_button = tkinter.Button(frame, text="Add Item", command=add_item)
add_item_button.grid(row=4, column=1, pady=10)
generate_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
generate_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)


# Remember we imported ttk from tkinter.That's to create tree view here
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text="Quantity")
tree.heading('desc', text="Description")
tree.heading('price', text="Unit Price")
tree.heading('total', text="Grand Total")
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

fname_entry = tkinter.Entry(frame)
lname_entry = tkinter.Entry(frame)
phone_entry = tkinter.Entry(frame)
# Later input the default value = "Null"
quantity_entry = tkinter.Spinbox(frame, from_=1, to=100)
description_entry = tkinter.Entry(frame)
price_entry = tkinter.Spinbox(frame, from_=0.0, to=1000, increment=0.5)


fname_entry.grid(row=1, column=0)
lname_entry.grid(row=1, column=1)
phone_entry.grid(row=1, column=2)
quantity_entry.grid(row=3, column=0)
description_entry.grid(row=3, column=1)
price_entry.grid(row=3, column=2)


# Basically to keep the gui running until the close
window.mainloop()
