import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
from datetime import datetime

def clear_item():
    Description_entry.delete(0, tkinter.END)
    Quantity_spinbox.delete(0, tkinter.END)
    Quantity_spinbox.insert(0, "1")
    unit_price_spinbox.delete(0, tkinter.END)
    unit_price_spinbox.insert(0,"0.0")

invoice_list = []
def add_item():
    Description = Description_entry.get()
    Quantity = int(Quantity_spinbox.get())
    unit_price = float(unit_price_spinbox.get())
    line_total = Quantity * unit_price
    invoice_item = [Description, Quantity, unit_price, line_total]

    tree.insert('',0, values=invoice_item)
    clear_item()

    invoice_list.append(invoice_item)

def new_invoice():
    Invoice_No_spinbox.delete(0, tkinter.END)
    Invoice_No_spinbox.insert(0, "1")
    clear_item()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def generate_invoice():
    doc = DocxTemplate("INVOICE1264.docx_020919 - Copy.docx")
    Invoice_No = Invoice_No_spinbox.get()
    Sub_Total = sum(item[3] for item in invoice_list)
    Total = Sub_Total
    doc.render({"Invoice_No": Invoice_No,
                "Date": datetime.today().strftime('%d-%B-%Y') ,
                "invoice_list": invoice_list,
                "Sub_Total": Sub_Total,
                "Total": Total})
    doc_name = "INVOICE" + Invoice_No +".docx"
    doc.save(doc_name)



window = tkinter.Tk()
window.title("Invoice Generator")

frame =  tkinter.Frame(window)
frame.pack()

Invoice_No_label = tkinter.Label(frame, text="Invoice No")
Invoice_No_label.grid(row=0, column=0) 
Invoice_No_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
Invoice_No_spinbox.grid(row=1, column=0)

Description_label = tkinter.Label(frame, text="Description")
Description_label.grid(row=0, column=1)
Description_entry = tkinter.Entry(frame)
Description_entry.grid(row=1, column=1)

Quantity_label = tkinter.Label(frame, text="Quantity")
Quantity_label.grid(row=0, column=2) 
Quantity_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
Quantity_spinbox.grid(row=1, column=2)

unit_price_label = tkinter.Label(frame, text="Unit Price")
unit_price_label.grid(row=2, column=0) 
unit_price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=100, increment=0.5)
unit_price_spinbox.grid(row=3, column=0)

add_item_button = tkinter.Button(frame, text="Add Item", command = add_item)
add_item_button.grid(row=3, column=1)

columns = ('Description', 'Quantity', 'Unit Price', 'Total')
tree = ttk.Treeview(frame, columns=columns, show="headings")

columns = ('Description', 'Quantity', 'Unit Price', 'Total')
tree.heading('Description', text='Description')
tree.heading('Quantity', text='Quantity')
tree.heading('Unit Price', text='Unit Price')
tree.heading('Total', text='Total')

tree.grid(row=4, column=0, columnspan=3, padx=20, pady=10)


Generate_invoice_button = tkinter.Button(frame, text="Generate Invoice", command = generate_invoice)
Generate_invoice_button.grid(row=5, column=0, columnspan=3, sticky="news", padx=20, pady=10)
New_invoice_button = tkinter.Button(frame, text="New Invoice", command = new_invoice)
New_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=10)
window.mainloop()