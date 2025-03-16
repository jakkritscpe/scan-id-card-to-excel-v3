import re
import xlwings as xw
import customtkinter as ctk
from tkinter import filedialog, messagebox
from smartcard.System import readers

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

def thai2unicode(data):
    return bytes(data).decode('tis-620').strip()

def clean_text(text):
    """ล้างข้อความจากอักขระพิเศษ"""
    return re.sub(r'[#\x00]+', ' ', text).strip()

def get_data(connection, cmd):
    req = [0x00, 0xc0, 0x00, 0x00, cmd[-1]]
    data, sw1, sw2 = connection.transmit(cmd)
    data, sw1, sw2 = connection.transmit(req)
    return thai2unicode(data)

def read_id_card():
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
    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row + 1    
        
        # get last row value -> []    
        last_row_values = sheet.range(f"A{last_row - 1}").expand('right').value
        
        # map key value result.
        card_data = dict(zip(list(COMMANDS.keys()), list(data.values())))
        
        # Read values from the selected column from row 1 to the last row
        column_value = sheet.range(f"E1:E{last_row}").value
        column_check = [str(value).replace(".0", "") for value in column_value]
            
        if not card_data["CID"] in column_check:
            # create new order.
            order = "A" + str(int(last_row_values[0].replace("A", "")) + 1)
            new_data = [order, card_data["TH Fullname"], card_data["Address"],"-",card_data["CID"],""]
            
            sheet.range(f"A{last_row}").value = new_data
            wb.save()
            messagebox.showinfo("Success", "Data saved successfully")
            
        else:
            messagebox.showwarning("Unscess", "This data already exists in the file.")
            
        return
            
    except Exception as e:
        return str(e)

def select_file():
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

def process_card():
    file_path = file_var.get()
    sheet_name = sheet_var.get()
    
    if not file_path or not sheet_name:
        output_var.set("Please select an Excel file and sheet first")
    else:
        card_data, error = read_id_card()
        if card_data:
            output_text = "\n".join(f"{k}: {clean_text(v)}" for k, v in card_data.items())
            output_var.set(output_text)
            
            # clean the text before saving to excel file
            cleaned_data = {k: clean_text(v) for k, v in card_data.items()}
            
            save_to_excel(file_path, sheet_name, cleaned_data)
            # output_var.set(output_var.get() + "\n" + save_status)
            
        elif error:
            output_var.set(error)
        else:
            output_var.set("No ID card detected.")
    
    # Schedule the next check
    root.after(5000, process_card)  # Check every 5000 milliseconds (5 seconds)

def update_output(*args):
    output_textbox.delete("1.0", "end")
    output_textbox.insert("1.0", output_var.get())
    

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
button_browse = ctk.CTkButton(frame_file, text="Browse", command=select_file, fg_color="#EEEEEE", text_color="#000000")
button_browse.pack(side="right", padx=10)

frame_sheet = ctk.CTkFrame(root)
frame_sheet.pack(pady=10, padx=20, fill="both")
label_sheet = ctk.CTkLabel(frame_sheet, text="Select Sheet", anchor="w")
label_sheet.pack(side="left", padx=10)
sheet_menu = ctk.CTkOptionMenu(frame_sheet, values=[], variable=sheet_var, fg_color="#EEEEEE", text_color="#000000")
sheet_menu.pack(side="right", padx=10)

output_textbox = ctk.CTkTextbox(root, wrap="word", width=400, height=150)
output_textbox.pack(pady=10, padx=20, fill="both", expand=True)

output_var.trace_add("write", update_output)

# Start the periodic card processing
root.after(000, process_card)  # Initial delay before the first check

root.mainloop()
