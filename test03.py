import os
import sys
import xlwings as xw
import customtkinter as ctk
from tkinter import filedialog
from smartcard.System import readers

# Constants for APDU commands
SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
THAI_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
COMMANDS = {
    "CID": [0x80, 0xB0, 0x00, 0x04, 0x02, 0x00, 0x0D],
    "TH Fullname": [0x80, 0xB0, 0x00, 0x11, 0x02, 0x00, 0x64],
    "EN Fullname": [0x80, 0xB0, 0x00, 0x75, 0x02, 0x00, 0x64],
    "DOB": [0x80, 0xB0, 0x00, 0xD9, 0x02, 0x00, 0x08],
    "Gender": [0x80, 0xB0, 0x00, 0xE1, 0x02, 0x00, 0x01],
    "Address": [0x80, 0xB0, 0x15, 0x79, 0x02, 0x00, 0x64]
}

def decode_thai(data):
    return bytes(data).decode('tis-620').strip()

def transmit_command(connection, command):
    try:
        data, sw1, sw2 = connection.transmit(command)
        response_command = [0x00, 0xC0, 0x00, 0x00, command[-1]]
        data, sw1, sw2 = connection.transmit(response_command)
        return decode_thai(data)
    except Exception as e:
        raise RuntimeError(f"Failed to transmit command {command}: {e}")

def read_id_card():
    try:
        available_readers = readers()
        if not available_readers:
            raise RuntimeError("No smartcard readers found.")

        reader = available_readers[0]
        connection = reader.createConnection()
        connection.connect()
        connection.transmit(SELECT + THAI_CARD)

        return {key: transmit_command(connection, cmd) for key, cmd in COMMANDS.items()}
    except Exception as e:
        raise RuntimeError(f"Failed to read ID card: {e}")

def save_to_excel(file_path, sheet_name, data):
    try:
        workbook = xw.Book(file_path)
        sheet = workbook.sheets[sheet_name]
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row + 1
        sheet.range(f"A{last_row}").value = list(data.values())
        workbook.save()
    except Exception as e:
        raise RuntimeError(f"Failed to save data to Excel: {e}")

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_var.set(file_path)
        try:
            workbook = xw.Book(file_path)
            sheet_names = [sheet.name for sheet in workbook.sheets]
            sheet_var.set(sheet_names[0])
            sheet_menu.configure(values=sheet_names)
        except Exception as e:
            output_var.set(f"Error loading Excel file: {e}")

def process_id_card():
    file_path = file_var.get()
    sheet_name = sheet_var.get()

    if not file_path or not sheet_name:
        output_var.set("Please select an Excel file and sheet first.")
        return

    try:
        card_data = read_id_card()
        output = "\n".join(f"{key}: {value}" for key, value in card_data.items())
        output_var.set(output)
        save_to_excel(file_path, sheet_name, card_data)
        output_var.set(output_var.get() + "\nData saved successfully.")
    except Exception as e:
        output_var.set(f"Error: {e}")

    root.after(5000, process_id_card)

def update_output_display(*args):
    output_textbox.delete("1.0", "end")
    output_textbox.insert("1.0", output_var.get())

ctk.set_appearance_mode("System")
root = ctk.CTk()
root.title("ID Card Reader")
root.geometry("500x450")

file_var = ctk.StringVar()
sheet_var = ctk.StringVar()
output_var = ctk.StringVar()

file_frame = ctk.CTkFrame(root)
file_frame.pack(pady=10, padx=20, fill="x")
label_file = ctk.CTkLabel(file_frame, text="Select Excel File:", anchor="w")
label_file.pack(side="left", padx=10)
button_browse = ctk.CTkButton(file_frame, text="Browse", command=select_excel_file, fg_color="#EEEEEE", text_color="#000000")
button_browse.pack(side="right", padx=10)

sheet_frame = ctk.CTkFrame(root)
sheet_frame.pack(pady=10, padx=20, fill="x")
label_sheet = ctk.CTkLabel(sheet_frame, text="Select Sheet:", anchor="w")
label_sheet.pack(side="left", padx=10)
sheet_menu = ctk.CTkOptionMenu(sheet_frame, values=[], variable=sheet_var, fg_color="#EEEEEE", text_color="#000000")
sheet_menu.pack(side="right", padx=10)

output_textbox = ctk.CTkTextbox(root, wrap="word", width=400, height=150)
output_textbox.pack(pady=10, padx=20, fill="both", expand=True)

output_var.trace_add("write", update_output_display)

root.after(0, process_id_card)
root.mainloop()
