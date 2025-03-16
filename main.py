import re
import xlwings as xw
import customtkinter as ctk
from tkinter import filedialog, messagebox
from smartcard.System import readers
import time
import threading

# คำสั่งอ่านข้อมูลจากบัตรประชาชน
SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
THAI_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
COMMANDS = {
    "CID":         [0x80, 0xb0, 0x00, 0x04, 0x02, 0x00, 0x0d],
    "TH_Fullname": [0x80, 0xb0, 0x00, 0x11, 0x02, 0x00, 0x64],
    "EN_Fullname": [0x80, 0xb0, 0x00, 0x75, 0x02, 0x00, 0x64],
    "DOB":         [0x80, 0xb0, 0x00, 0xD9, 0x02, 0x00, 0x08],
    "Gender":      [0x80, 0xb0, 0x00, 0xE1, 0x02, 0x00, 0x01],
    "Address":     [0x80, 0xb0, 0x15, 0x79, 0x02, 0x00, 0x64]
}

def thai2unicode(data):
    """แปลงข้อมูลจาก TIS-620 เป็น Unicode"""
    return bytes(data).decode('tis-620').strip()

def clean_text(text):
    """ล้างข้อความจากอักขระพิเศษ"""
    return re.sub(r'[#\x00]+', ' ', text).strip()

def get_data(connection, cmd):
    """อ่านข้อมูลจากบัตร"""
    req = [0x00, 0xc0, 0x00, 0x00, cmd[-1]]
    data, sw1, sw2 = connection.transmit(cmd)
    data, sw1, sw2 = connection.transmit(req)
    return thai2unicode(data)

def read_id_card():
    """อ่านข้อมูลจากบัตรประชาชน"""
    try:
        reader_list = readers()
        if not reader_list:
            return None, "No reader found"
        
        reader = reader_list[0]
        connection = reader.createConnection()
        connection.connect()

        connection.transmit(SELECT + THAI_CARD)

        data = {key: get_data(connection, cmd) for key, cmd in COMMANDS.items()}
        return data, None
    except Exception as e:
        return None, str(e)

def save_to_excel(file_path, sheet_name, data):
    """บันทึกข้อมูลลงไฟล์ Excel"""
    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row + 1    

        card_data = {key: clean_text(value) for key, value in data.items()}
        column_value = sheet.range(f"E1:E{last_row}").value
        column_check = [str(value).replace(".0", "") for value in column_value]

        if card_data["CID"] not in column_check:
            order = f"A{last_row}"
            new_data = [order, card_data["TH_Fullname"], card_data["Address"], "-", card_data["CID"], ""]
            sheet.range(f"A{last_row}").value = new_data
            wb.save()
            messagebox.showinfo("Success", "Data saved successfully")
        else:
            messagebox.showwarning("Unsuccess", "This data already exists in the file.")
    except Exception as e:
        return str(e)

def select_file():
    """เลือกไฟล์ Excel"""
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_var.set(file_path)
        wb = xw.Book(file_path)
        sheets = [sheet.name for sheet in wb.sheets]

        if "Cus" in sheets:
            sheet_var.set("Cus")
        elif sheets:
            sheet_var.set(sheets[0])
        
        sheet_menu.configure(values=sheets)

def process_card(file_path, sheet_name):
    """อ่านบัตรประชาชนและบันทึกข้อมูลลง Excel"""
    card_data, error = read_id_card()
    if card_data:
        output_text = "\n".join(f"{k}: {clean_text(v)}" for k, v in card_data.items())
        output_var.set(output_text)
        save_to_excel(file_path, sheet_name, card_data)
    elif error:
        output_var.set(error)
    else:
        output_var.set("No ID card detected.")

def update_output(*args):
    """อัปเดตผลลัพธ์ใน GUI"""
    output_textbox.delete("1.0", "end")
    output_textbox.insert("1.0", output_var.get())

def card_monitor_loop():
    """ตรวจจับการเสียบและถอดบัตร"""
    prev_status = False  

    while True:
        try:
            file_path = file_var.get()
            sheet_name = sheet_var.get()
            
            if not file_path or not sheet_name:
                output_var.set("Please select an Excel file and sheet first")
            else:
                reader_list = readers()
                if reader_list:
                    reader = reader_list[0]
                    connection = reader.createConnection()

                    try:
                        connection.connect()
                        # output_var.set("Conneted..")
                        current_status = True
                    except:
                        # output_var.set("Can not conneted..")
                        current_status = False
                else:
                    current_status = False

                if current_status and not prev_status:
                    output_var.set("ID card inserted. Processing...")
                    process_card(file_path, sheet_name)

                if not current_status and prev_status:
                    output_var.set("ID card removed.")

                prev_status = current_status
        except Exception as e:
            output_var.set(f"Error: {str(e)}")

        time.sleep(3)  # ตรวจสอบทุก 1 วินาที

#######################
#     UI Setup        #
#######################
ctk.set_appearance_mode("Dark")
root = ctk.CTk()
root.title("ID Card Reader")
root.geometry("300x300")
root.resizable(False, False)

file_var = ctk.StringVar()
sheet_var = ctk.StringVar()
output_var = ctk.StringVar()

frame_file = ctk.CTkFrame(root)
frame_file.pack(pady=10, padx=20, fill="both")
label_file = ctk.CTkLabel(frame_file, text="Select Excel File", anchor="w")
label_file.pack(side="left", padx=10)
button_browse = ctk.CTkButton(frame_file, text="Browse", command=select_file)
button_browse.pack(side="right", padx=10)

frame_sheet = ctk.CTkFrame(root)
frame_sheet.pack(pady=10, padx=20, fill="both")
label_sheet = ctk.CTkLabel(frame_sheet, text="Select Sheet", anchor="w")
label_sheet.pack(side="left", padx=10)
sheet_menu = ctk.CTkOptionMenu(frame_sheet, values=[], variable=sheet_var)
sheet_menu.pack(side="right", padx=10)

output_textbox = ctk.CTkTextbox(root, wrap="word", width=400, height=150)
output_textbox.pack(pady=10, padx=20, fill="both", expand=True)

output_var.trace_add("write", update_output)

# ใช้ threading เพื่อให้ UI ไม่ค้าง
threading.Thread(target=card_monitor_loop, daemon=True).start()

root.mainloop()
