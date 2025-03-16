import os
import sys
import xlwings as xw
import customtkinter as ctk
from tkinter import filedialog
from smartcard.System import readers

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
        
        return {key: get_data(connection, cmd) for key, cmd in COMMANDS.items()}
    except Exception as e:
        return str(e)

def toggle_language(event=None):
    global lang
    lang = "TH" if lang == "EN" else "EN"
    update_texts()

def update_texts():
    texts = {
        "EN": {
            "select_file": "Select Excel File:",
            "browse": "Browse",
            "select_sheet": "Select Sheet:",
            "select_language": "Select Language:",
            "read_card": "Read ID Card"
        },
        "TH": {
            "select_file": "เลือกไฟล์ Excel:",
            "browse": "เรียกดู",
            "select_sheet": "เลือกชีต:",
            "select_language": "เลือกภาษา:",
            "read_card": "อ่านบัตรประชาชน"
        }
    }
    lang_data = texts[lang]
    label_file.configure(text=lang_data["select_file"])
    button_browse.configure(text=lang_data["browse"])
    label_sheet.configure(text=lang_data["select_sheet"])
    label_language.configure(text=lang_data["select_language"])
    button_read.configure(text=lang_data["read_card"])
    button_toggle_language.configure(text=f"Switch to {'English' if lang == 'TH' else 'Thai'}")

# UI Setup
ctk.set_appearance_mode("System")
root = ctk.CTk()
root.title("ID Card Reader")
root.geometry("500x500")

file_var = ctk.StringVar()
sheet_var = ctk.StringVar()
output_var = ctk.StringVar()
lang = "EN"

frame_file = ctk.CTkFrame(root)
frame_file.pack(pady=10, padx=20, fill="x")
label_file = ctk.CTkLabel(frame_file, text="Select Excel File:", anchor="w")
label_file.pack(side="left", padx=10)
# button_browse = ctk.CTkButton(frame_file, text="Browse", command=select_file, fg_color="#EEEEEE", text_color="#000000")
# button_browse.pack(side="right", padx=10)

frame_sheet = ctk.CTkFrame(root)
frame_sheet.pack(pady=10, padx=20, fill="x")
label_sheet = ctk.CTkLabel(frame_sheet, text="Select Sheet:", anchor="w")
label_sheet.pack(side="left", padx=10)
sheet_menu = ctk.CTkOptionMenu(frame_sheet, values=[], variable=sheet_var, fg_color="#EEEEEE", text_color="#000000")
sheet_menu.pack(side="right", padx=10)

frame_lang = ctk.CTkFrame(root)
frame_lang.pack(pady=10, padx=20, fill="x")
label_language = ctk.CTkLabel(frame_lang, text="Select Language:", anchor="w")
label_language.pack(side="left", padx=10)
lang_menu = ctk.CTkOptionMenu(frame_lang, values=["EN", "TH"], command=toggle_language, fg_color="#EEEEEE", text_color="#000000")
lang_menu.pack(side="right", padx=10)

button_read = ctk.CTkButton(root, text="Read ID Card", command=read_id_card, fg_color="#28A745", text_color="#FFFFFF", corner_radius=8, width=200)
button_read.pack(pady=20)

button_toggle_language = ctk.CTkButton(root, text="Switch to Thai", command=toggle_language, fg_color="#AAAAAA", text_color="#FFFFFF", corner_radius=8, width=200)
button_toggle_language.pack(pady=10)

output_frame = ctk.CTkFrame(root)
output_frame.pack(pady=10, padx=20, fill="both", expand=True)
ctk.CTkLabel(output_frame, textvariable=output_var, wraplength=400, justify="left", anchor="w").pack(pady=10, padx=10)

update_texts()
root.mainloop()
