import re
import xlwings as xw
import customtkinter as ctk
from tkinter import filedialog, messagebox
from smartcard.System import readers
import time
import threading
import logging

# Logging configuration
logging.basicConfig(level=logging.DEBUG)

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

# 
LANG = "TH"
CODE_MSG = {
    "EN": {
        "001": "No reader found",
        "002": "Success",
        "003": "Data saved successfully",
        "004": "Unsuccess",
        "005": "This data already exists in the file.",
        "006": "No ID card detected.",
        "007": "Please select an Excel file and sheet first",
        "008": "ID card inserted. Processing...",
        "009": "ID card removed.",
        "010": "Man",
        "011": "Women"
    },
    "TH": {
        "001": "ไม่พบเครื่องอ่าน",
        "002": "สำเร็จ",
        "003": "บันทึกข้อมูลเรียบร้อยแล้ว",
        "004": "ไม่สำเร็จ",
        "005": "ข้อมูลนี้มีอยู่ในไฟล์แล้ว",
        "006": "ไม่พบบัตรประชาชน",
        "007": "กรุณาเลือกไฟล์ Excel และแผ่นงานก่อน",
        "008": "บัตรประชาชนถูกใส่ กำลังประมวลผล...",
        "009": "บัตรประชาชนถูกถอดออก",
        "010": "ชาย",
        "011": "หญิง"
    }
}


def thai2unicode(data):
    """แปลงข้อมูลจาก TIS-620 เป็น Unicode"""
    return bytes(data).decode('tis-620').strip()

def clean_text(text):
    """ล้างข้อความจากอักขระพิเศษ"""
    return re.sub(r'[#\x00]+', ' ', text).strip()

def gender(code):
    if code == str(1):
        return CODE_MSG[LANG]["010"]
    else:
        return CODE_MSG[LANG]["011"]

def birth_day(data):
    if data:
        return f"{data[6:8]}/{data[4:6]}/{data[0:4]}"
    else:
        return data

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
            return None, CODE_MSG[LANG]["001"]
        
        reader = reader_list[0]
        connection = reader.createConnection()
        connection.connect()

        connection.transmit(SELECT + THAI_CARD)

        data = {key: get_data(connection, cmd) for key, cmd in COMMANDS.items()}
        return data, None
    except Exception as e:
        return None, str("In fuction : read_id_card " + e)

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
            messagebox.showinfo(CODE_MSG[LANG]["002"], CODE_MSG[LANG]["003"])
        else:
            messagebox.showwarning(CODE_MSG[LANG]["004"], CODE_MSG[LANG]["005"])
    except Exception as e:
        logging.error(f"เกิดข้อผิดพลาดในหารบันทึกข้อมูล: {e}")

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
        # สร้างรายการข้อความที่จัดรูปแบบแล้ว
        formatted_lines = []
        for key, value in card_data.items():
            if key == "Gender":
                value = gender(value)
            
            if key == "DOB":
                value = birth_day(value)
                
            cleaned_value = clean_text(value)
            formatted_line = f"{key} : {cleaned_value}"
            formatted_lines.append(formatted_line)

        # รวมข้อความทั้งหมดเข้าด้วยกัน โดยแต่ละข้อความอยู่ในบรรทัดใหม่
        output_text = "\n".join(formatted_lines)
        output_var.set(output_text)
        save_to_excel(file_path, sheet_name, card_data)
    elif error:
        output_var.set(error)
    else:
        output_var.set(CODE_MSG[LANG]["006"])

def update_output(*args):
    """อัปเดตผลลัพธ์ใน GUI"""
    output_textbox.delete("1.0", "end")
    output_textbox.insert("1.0", output_var.get())

def card_monitor_loop():
    """ตรวจจับการเสียบและถอดบัตร"""
    prev_status = False  
    
    title_status = True

    while True:
        try:
            file_path = file_var.get()
            sheet_name = sheet_var.get()
            
            if not file_path or not sheet_name:
                output_var.set(CODE_MSG[LANG]["007"])
            else:
                reader_list = readers()
                if reader_list:
                    reader = reader_list[0]
                    connection = reader.createConnection()
                    
                    if title_status:
                        output_var.set(f"กำลังเชื่อมต่อกับเครื่องอ่านบัตร {reader} \n\n กรุณาเสียบบัตร / ถอดแล้วเสียบบัตรอีกครั้ง...")

                    try:
                        connection.connect()
                        current_status = True
                        if current_status and not prev_status:
                            output_var.set(CODE_MSG[LANG]["008"])
                            title_status = False
                            process_card(file_path, sheet_name)
                    except:
                        title_status = True
                        current_status = False
                        # if not current_status and prev_status:
                        #     output_var.set(CODE_MSG[LANG]["009"])
                        
                else:
                    current_status = False
                    output_var.set("⚠️ การเชื่อมต่อเครื่องอ่านบัตรไม่สำเร็จ...")

                prev_status = current_status
        except Exception as e:
            output_var.set(f"เกิดข้อผิดพลาดใน Monitor: {str(e)}")

        time.sleep(1)  # ตรวจสอบทุก 1 วินาที

#######################
#     UI Setup        #
#######################
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")
root = ctk.CTk()
root.title("ID Card Reader")
root.iconbitmap("asset/3d.ico")
root.geometry("300x300")
root.resizable(False, False)

file_var = ctk.StringVar()
sheet_var = ctk.StringVar()
output_var = ctk.StringVar()

frame_file = ctk.CTkFrame(root)
frame_file.pack(pady=(20,5), padx=20, fill="both", expand=True)
label_file = ctk.CTkLabel(frame_file, text="Select Excel File", anchor="w")
label_file.pack(side="left", padx=10)
button_browse = ctk.CTkButton(frame_file, text="Browse", command=select_file, width=128)
button_browse.pack(side="right", padx=10)

frame_sheet = ctk.CTkFrame(root)
frame_sheet.pack(pady=5, padx=20, fill="both", expand=True)
label_sheet = ctk.CTkLabel(frame_sheet, text="Select Sheet", anchor="w")
label_sheet.pack(side="left", padx=10)
sheet_menu = ctk.CTkOptionMenu(frame_sheet, values=[], variable=sheet_var, width=128)
sheet_menu.pack(side="right", padx=10)

output_textbox = ctk.CTkTextbox(root, wrap="word", width=400, height=100)
output_textbox.pack(pady=(10,20), padx=20, fill="both", expand=True)

output_var.trace_add("write", update_output)

# ใช้ threading เพื่อให้ UI ไม่ค้าง
threading.Thread(target=card_monitor_loop, daemon=True).start()

root.mainloop()
