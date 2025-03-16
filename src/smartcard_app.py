import customtkinter as ctk
import time
import threading
import logging
import psutil
import tkinter.messagebox as messagebox
import queue
from tkinter import filedialog
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import NoCardException, CardConnectionException
from .excel_manager import ExcelManager
from .config import CONFIG, AID, COMMANDS
from .utils import clean_text, thai2unicode

# Logging configuration
logging.basicConfig(level=logging.DEBUG)

# CustomTkinter theme configuration
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

class SmartCardApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("โปรแกรมอ่านบัตรประชาชน")
        self.geometry("500x480")
        self.resizable(False, False)
        
        # ตัวแปรควบคุม
        self.running = True
        self.card_processing = False

        # ExcelManager เริ่มต้นเป็น None (รอให้ผู้ใช้เลือกไฟล์และแผ่นงาน)
        self.excel = None
        self.sheet_names = []
        self.selected_sheet = ""
        
        # สร้าง Queue สำหรับงาน Excel (ให้ Main Thread จัดการ COM operations)
        self.excel_queue = queue.Queue()

        self.status = ""
        self.status_color = ""
        
        # สร้าง GUI
        self.create_widgets()
        
        # เริ่มเธรดอ่านบัตร (เธรดนี้จะตรวจสอบก่อนว่ามีการเลือก Excel แล้วหรือไม่)
        self.reader_thread = threading.Thread(target=self.card_reader_handler, daemon=True)
        self.reader_thread.start()

        # ตรวจสอบ Excel ทุก 1 วินาที
        self.after(1000, self.check_excel_running)
        # ประมวลผล Queue งาน Excel ใน Main Thread
        self.after(500, self.process_excel_queue)
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def create_widgets(self):
        """สร้างส่วนประกอบ GUI"""
        # ส่วนแสดงสถานะระบบ
        self.status_label = ctk.CTkLabel(
            self, 
            text="🔄 กำลังตรวจสอบเครื่องอ่าน...", 
            font=("Arial", 16),
            wraplength=450
        )
        self.status_label.pack(pady=10)
        
        # --- ส่วนเลือกไฟล์ Excel และเลือกแผ่นงาน --- 
        self.excel_file_label = ctk.CTkLabel(self, text="เลือกไฟล์ Excel:")
        self.excel_file_label.pack(pady=5)
        
        self.file_button = ctk.CTkButton(self, text="เลือกไฟล์", command=self.select_file)
        self.file_button.pack(pady=5)
        
        # Entry สำหรับแสดงพาธไฟล์ (Read-Only)
        self.path_entry = ctk.CTkEntry(self, width=300, state="disabled")
        self.path_entry.pack(pady=5)
        
        self.sheet_label = ctk.CTkLabel(self, text="เลือกชื่อแผ่น:")
        self.sheet_label.pack(pady=5)
        
        self.sheet_combobox = ctk.CTkComboBox(self, values=[], command=self.select_sheet)
        self.sheet_combobox.pack(pady=5)
        # ----------------------------------------------------
        
        # Textbox สำหรับแสดงผลการอ่านบัตร
        self.text_area = ctk.CTkTextbox(
            self, 
            width=450, 
            height=220, 
            font=("Arial", 14),
            state="disabled"
        )
        self.text_area.pack(pady=10)
    
    def select_file(self):
        """เลือกไฟล์ Excelและโหลดชื่อแผ่นงาน"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if file_path:
            # แสดงพาธใน Entry (Read-Only)
            self.path_entry.configure(state="normal")
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, file_path)
            self.path_entry.configure(state="disabled")
            
            # โหลดไฟล์ Excelและดึงชื่อแผ่น
            self.load_excel(file_path)
    
    def load_excel(self, file_path):
        """โหลดไฟล์ Excelและดึงชื่อแผ่นงาน"""
        try:
            # สร้าง ExcelManager แบบชั่วคราวเพื่อดึงชื่อแผ่น
            temp_excel = ExcelManager(file_path, "")
            self.sheet_names = temp_excel.get_sheet_names()
            
            if len(self.sheet_names) == 1:
                # ถ้ามีแผ่นเดียว ให้เลือกอัตโนมัติและปิด ComboBox
                self.selected_sheet = self.sheet_names[0]
                self.sheet_combobox.set(self.selected_sheet)
                self.sheet_combobox.configure(state="disabled")
            else:
                # ถ้ามีหลายแผ่น ให้เปิด ComboBoxให้เลือก
                self.sheet_combobox.configure(values=self.sheet_names, state="normal")
                self.sheet_combobox.set("")

            # หากมีแผ่นถูกเลือกแล้ว ให้สร้าง ExcelManager จริง
            if self.selected_sheet:
                try:
                    self.excel = ExcelManager(file_path, self.selected_sheet)
                except Exception as e:
                    logging.error("เกิดข้อผิดพลาดในการสร้าง ExcelManager: %s", e, exc_info=True)
                    messagebox.showerror("Excel Error", f"เกิดข้อผิดพลาดในการเปิดไฟล์ Excel: {e}")
                    self.excel = None
            else:
                self.excel = None  # รอให้ผู้ใช้เลือกแผ่นใน ComboBox
        except Exception as e:
            logging.error("ไม่สามารถเปิดไฟล์ Excel: %s", e, exc_info=True)
            messagebox.showerror("Excel Error", f"ไม่สามารถเปิดไฟล์ Excel: {e}")
    
    def select_sheet(self, sheet_name):
        """เมื่อเลือกแผ่นงานใน ComboBox ให้บันทึกและสร้าง ExcelManager ใหม่"""
        self.selected_sheet = sheet_name
        logging.info("เลือกแผ่น: %s", self.selected_sheet)
        file_path = self.path_entry.get()
        if file_path:
            try:
                self.excel = ExcelManager(file_path, self.selected_sheet)
            except Exception as e:
                logging.error("โหลดแผ่นงานไม่สำเร็จ: %s", e, exc_info=True)
                messagebox.showerror("Excel Error", f"โหลดแผ่นงานไม่สำเร็จ: {e}")
    
    def check_excel_running(self):
        """ตรวจสอบว่า Excel ยังทำงานอยู่หรือไม่"""
        if self.excel and self.excel.app:
            try:
                if not psutil.pid_exists(self.excel.app.pid):
                    logging.info("Excel ปิดไปแล้ว กำลังปิดโปรแกรม...")
                    self.on_close()
                    return
            except Exception as e:
                logging.error("เกิดข้อผิดพลาดในการตรวจสอบ Excel: %s", e, exc_info=True)
                self.on_close()
                return
        self.after(1000, self.check_excel_running)
    
    def card_reader_handler(self):
        """เธรดหลักสำหรับอ่านบัตร"""
        while self.running:
            # ตรวจสอบว่ามีการเลือก Excel แล้วหรือไม่
            if self.excel is None:
                self.update_gui("โปรดเลือกไฟล์ Excel และแผ่นงานก่อน", "error")
                time.sleep(1)
                continue

            try:
                reader = self.get_reader()
                if reader:
                    self.monitor_card(reader)
                else:
                    self.update_gui("❌ ไม่พบเครื่องอ่านบัตร!", "error")
                    time.sleep(2)
            except Exception as e:
                logging.error("ข้อผิดพลาดระบบ: %s", e, exc_info=True)
                self.update_gui("⚠️ เกิดข้อผิดพลาดระบบ!", "error")
                time.sleep(1)
    
    def get_reader(self):
        """ตรวจสอบและคืนค่าเครื่องอ่านบัตรที่พร้อมใช้งาน"""
        try:
            readers_list = readers()
            if not readers_list:
                logging.warning("ไม่พบเครื่องอ่านบัตร")
                return None
            logging.info("พบเครื่องอ่านบัตร: %s", [str(r) for r in readers_list])
            return readers_list[0]
        except Exception as e:
            logging.error("การตรวจสอบเครื่องอ่านล้มเหลว: %s", e, exc_info=True)
            return None
    
    def monitor_card(self, reader):
        """ตรวจสอบการเสียบบัตร"""
        while self.running and not self.card_processing:
            try:
                connection = reader.createConnection()
                connection.connect()
                self.process_card(connection)
            except (NoCardException, CardConnectionException):
                self.update_gui("💳 กรุณาเสียบบัตรประชาชน...", "info")
                time.sleep(1)
            except Exception as e:
                logging.error("การเชื่อมต่อล้มเหลว: %s", e, exc_info=True)
                self.update_gui("⚠️ การเชื่อมต่อผิดพลาด!", "error")
                time.sleep(1)
    
    def process_card(self, connection):
        """ประมวลผลข้อมูลบัตร"""
        self.card_processing = True
        try:
            self.update_gui("📖 กำลังอ่านข้อมูล...", "info")
            self.clear_text_area()
            
            # อ่าน ATR ของบัตร
            atr = connection.getATR()
            logging.info("ATR: %s", toHexString(atr))
            
            # เลือกแอปพลิเคชัน
            if not self.select_application(connection, atr):
                raise Exception("ไม่พบแอปพลิเคชันที่รองรับ")
            
            # อ่านข้อมูลจากบัตรตาม COMMANDS
            for key in COMMANDS:
                data = self.execute_command(connection, key)
                if data:
                    self.append_text(f"{key}: {clean_text(thai2unicode(data))}\n")
            
            order_code = self.excel.get_last_row()  # ดึงแถวสุดท้ายใน Excel
            order_code = str(int(order_code) + 1)
            cid = clean_text(thai2unicode(self.execute_command(connection, "CID")))

            is_duplicated = self.excel.check_duplicate("E", cid)
            self.update_gui("อ่านข้อมูลสำเร็จ", "success")
                    
            # ถ้าไม่ซ้ำ ให้เพิ่มข้อมูลลง Excel (ส่งงานผ่าน Queue ให้ Main Thread จัดการ)
            if not is_duplicated:
                new_data = [
                    f"A{order_code}",
                    clean_text(thai2unicode(self.execute_command(connection, "THFULLNAME"))),
                    clean_text(thai2unicode(self.execute_command(connection, "ADDRESS"))),
                    "-",
                    cid,
                    "",
                ]
                self.excel_queue.put(new_data)
                messagebox.showinfo("สำเร็จ", "บันทึกข้อมูลสำเร็จ")
                logging.info("เพิ่มข้อมูลลง Excel: %s", new_data)
                time.sleep(1)
            else:
                messagebox.showwarning("แจ้งเตือน", "มีข้อมูลนี้อยู่ในไฟล์แล้ว")
                logging.info("เจอข้อมูลซ้ำ: %s", cid)
                time.sleep(1)
            
            # รอจนกว่าบัตรจะถูกถอดออก
            self.wait_for_card_removal(connection)
            
        except Exception as e:
            logging.error("การประมวลผลล้มเหลว: %s", e, exc_info=True)
            self.update_gui(f"ข้อผิดพลาด: {e}", "error")
        finally:
            self.card_processing = False
            connection.disconnect()
    
    def select_application(self, connection, atr):
        """เลือกแอปพลิเคชันด้วย AID ที่เหมาะสม"""
        try:
            select_cmd = [0x00, 0xA4, 0x04, 0x0C, len(AID)] + AID
            connection.transmit(select_cmd)
            return True
        except Exception as e:
            logging.error("เลือก AID ล้มเหลว: %s", e, exc_info=True)
        return False
    
    def execute_command(self, connection, command_key):
        """ดำเนินการคำสั่ง APDU"""
        try:
            cmd = COMMANDS[command_key]
            data, sw1, sw2 = connection.transmit(cmd)
            if sw1 == 0x61:
                get_response = [0x00, 0xC0, 0x00, 0x00, sw2]
                data, sw1, sw2 = connection.transmit(get_response)
            if sw1 != 0x90:
                raise Exception(f"คำสั่ง {command_key} ล้มเหลว (SW={hex(sw1)}{hex(sw2)})")
            return data
        except Exception as e:
            logging.error("คำสั่ง %s ผิดพลาด: %s", command_key, e, exc_info=True)
            return None
    
    def wait_for_card_removal(self, connection):
        """รอจนกว่าบัตรจะถูกถอดออก"""
        while self.running:
            try:
                connection.transmit([0x00, 0x84, 0x00, 0x00, 0x00])
                time.sleep(0.5)
            except Exception:
                self.update_gui("บัตรถูกถอดออกแล้ว", "info")
                break
    
    def update_gui(self, status, color):
        """อัปเดตสถานะ GUI"""
        self.status_label.configure(text=status)
        if color == "success":
            self.status_label.configure(text_color="green")
        elif color == "error":
            self.status_label.configure(text_color="red")
        else:
            self.status_label.configure(text_color="orange")
    
    def clear_text_area(self):
        """เคลียร์ข้อความใน Textbox"""
        self.text_area.configure(state="normal")
        self.text_area.delete(1.0, "end")
        self.text_area.configure(state="disabled")
    
    def append_text(self, text):
        """เพิ่มข้อความลงใน Textbox"""
        self.text_area.configure(state="normal")
        self.text_area.insert("end", text)
        self.text_area.configure(state="disabled")
    
    def process_excel_queue(self):
        """ประมวลผลงาน Excel จาก Queue"""
        while not self.excel_queue.empty():
            data = self.excel_queue.get()
            try:
                self.excel.add_data(data)
            except Exception as e:
                logging.error("บันทึกข้อมูลลง Excel ผิดพลาด: %s", e)
        self.after(500, self.process_excel_queue)

    def on_close(self):
        """เมื่อโปรแกรมปิด"""
        self.running = False
        self.destroy()
