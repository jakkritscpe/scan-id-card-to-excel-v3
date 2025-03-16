import os
import binascii
import sys
import xlwings as xw
import customtkinter as ctk
from tkinter import filedialog
from smartcard.System import readers
from smartcard.util import toHexString



def thai2unicode(data):
    return bytes(data).decode('tis-620').strip()

def get_data(connection, cmd):
    req = [0x00, 0xc0, 0x00, 0x00, cmd[-1]]
    data, sw1, sw2 = connection.transmit(cmd)
    data, sw1, sw2 = connection.transmit(req)
    return thai2unicode(data)

def read_id_card():
    try:
        reader_list = readers()
        if not reader_list:
            return "No reader found"
        
        reader = reader_list[0]
        connection = reader.createConnection()
        connection.connect()
        
        # Commands
        SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
        THAI_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
        COMMANDS = {
            "CID": [0x80, 0xb0, 0x00, 0x04, 0x02, 0x00, 0x0d],
            "TH Fullname": [0x80, 0xb0, 0x00, 0x11, 0x02, 0x00, 0x64],
            "EN Fullname": [0x80, 0xb0, 0x00, 0x75, 0x02, 0x00, 0x64],
            "DOB": [0x80, 0xb0, 0x00, 0xD9, 0x02, 0x00, 0x08],
            "Gender": [0x80, 0xb0, 0x00, 0xE1, 0x02, 0x00, 0x01],
            "Address": [0x80, 0xb0, 0x15, 0x79, 0x02, 0x00, 0x64]
        }
        
        connection.transmit(SELECT + THAI_CARD)
        
        card_data = {key: get_data(connection, cmd) for key, cmd in COMMANDS.items()}
        return card_data
    
    except Exception as e:
        return str(e)

def save_to_excel(file_path, sheet_name, data):
    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row + 1
        sheet.range(f"A{last_row}").value = list(data.values())
        wb.save()
        return "Data saved successfully"
    except Exception as e:
        return str(e)

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_var.set(file_path)
        wb = xw.Book(file_path)
        sheets = [sheet.name for sheet in wb.sheets]
        sheet_var.set(sheets[0])
        sheet_menu.configure(values=sheets)

def process_card():
    file_path = file_var.get()
    sheet_name = sheet_var.get()
    if not file_path or not sheet_name:
        output_var.set("Please select an Excel file and sheet first")
        return
    
    card_data = read_id_card()
    if isinstance(card_data, dict):
        output_var.set("\n".join(f"{k}: {v}" for k, v in card_data.items()))
        save_status = save_to_excel(file_path, sheet_name, card_data)
        output_var.set(output_var.get() + "\n" + save_status)
    else:
        output_var.set(card_data)

# UI Setup
ctk.set_appearance_mode("System")
root = ctk.CTk()
root.title("ID Card Reader")
root.geometry("500x400")

file_var = ctk.StringVar()
sheet_var = ctk.StringVar()
output_var = ctk.StringVar()

ctk.CTkLabel(root, text="Select Excel File:").pack(pady=5)
ctk.CTkButton(root, text="Browse", command=select_file).pack()
ctk.CTkLabel(root, textvariable=file_var).pack(pady=5)

ctk.CTkLabel(root, text="Select Sheet:").pack(pady=5)
sheet_menu = ctk.CTkOptionMenu(root, variable=sheet_var)
sheet_menu.pack()

ctk.CTkButton(root, text="Read ID Card", command=process_card).pack(pady=10)
ctk.CTkLabel(root, textvariable=output_var, wraplength=400, justify="left").pack(pady=10)

root.mainloop()
