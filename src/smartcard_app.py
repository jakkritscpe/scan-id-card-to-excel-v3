import customtkinter as ctk
import time
import threading
import logging
import psutil
import tkinter.messagebox as messagebox
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
ctk.set_default_color_theme("blue")

class SmartCardApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("โปรแกรมอ่านบัตรประชาชน")
        self.geometry("500x360")
        self.resizable(False, False)
        
        # ตัวแปรควบคุม
        self.running = True
        self.card_processing = False
        self.excel = ExcelManager(CONFIG["excel_path"], CONFIG["sheet_name"])

        self.status = ""
        self.status_color = ""
        
        # สร้าง GUI
        self.create_widgets()
        
        # เริ่มเธรดอ่านบัตร
        self.reader_thread = threading.Thread(target=self.card_reader_handler, daemon=True)
        self.reader_thread.start()

        # ตรวจสอบ Excel ทุก 1 วินาที
        self.after(1000, self.check_excel_running)
                        

    def create_widgets(self):
        """สร้างส่วนประกอบ GUI"""
        # สถานะระบบ
        self.status_label = ctk.CTkLabel(
            self, 
            text="🔄 กำลังตรวจสอบเครื่องอ่าน...", 
            font=("Arial", 16),
            wraplength=450
        )
        self.status_label.pack(pady=10)

        # พื้นที่แสดงผลข้อมูล
        self.text_area = ctk.CTkTextbox(
            self, 
            width=450, 
            height=250, 
            font=("Arial", 14),
            state="disabled",
            corner_radius=10  # เพิ่ม border radius
        )
        self.text_area.pack(pady=10)

        # ปุ่มออก
        # self.exit_button = ctk.CTkButton(
        #     self, 
        #     text="ออก", 
        #     command=self.on_close,
        #     fg_color="#d9534f",
        #     hover_color="#c9302c"
        # )
        # self.exit_button.pack(pady=10)

    def check_excel_running(self):
        """ตรวจสอบว่า Excel ยังทำงานอยู่หรือไม่"""
        if self.excel and self.excel.app:
            try:
                # ตรวจสอบว่า Process Excel ยังทำงานอยู่ไหม
                if not psutil.pid_exists(self.excel.app.pid):  
                    logging.info("Excel ปิดไปแล้ว กำลังปิดโปรแกรม...")
                    self.on_close()
                    return
            except Exception as e:
                logging.error(f"เกิดข้อผิดพลาดในการตรวจสอบ Excel: {str(e)}")
                self.on_close()
                return
        
        # เรียกตัวเองใหม่ทุก 1 วินาที
        self.after(1000, self.check_excel_running)

    def card_reader_handler(self):
        """จัดการเธรดหลักสำหรับอ่านบัตร"""
        while self.running:
            try:
                reader = self.get_reader()
                if reader:
                    self.monitor_card(reader)
                else:
                    self.update_gui("❌ ไม่พบเครื่องอ่านบัตร!", "error")
                    time.sleep(2)
            except Exception as e:
                logging.error(f"ข้อผิดพลาดระบบ: {str(e)}")
                self.update_gui("⚠️ เกิดข้อผิดพลาดระบบ!", "error")
                time.sleep(1)

    def get_reader(self):
        """
        ตรวจสอบและคืนค่าเครื่องอ่านบัตรที่พร้อมใช้งาน

        Returns:
            object: เครื่องอ่านบัตรตัวแรกที่พบ ถ้าไม่พบจะคืนค่า None
        """
        try:
            readers_list = readers()
            if not readers_list:
                logging.warning("ไม่พบเครื่องอ่านบัตร")
                return None

            # แสดง log รายการเครื่องอ่านที่พบ (ถ้ามีมากกว่าหนึ่ง)
            logging.info(f"พบเครื่องอ่านบัตร: {[str(r) for r in readers_list]}")
            return readers_list[0]
        except Exception as e:
            logging.error(f"การตรวจสอบเครื่องอ่านล้มเหลว: {str(e)}")
            return None


    def monitor_card(self, reader):
        """ตรวจสอบการเสียบบัตร"""
        while self.running and not self.card_processing:
            try:
                # สร้าง connection ใหม่ในแต่ละรอบ เพื่อให้แน่ใจว่าใช้ connection ที่ใช้งานได้อยู่
                connection = reader.createConnection()
                connection.connect()
                self.process_card(connection)
            except (NoCardException, CardConnectionException):
                self.update_gui("💳 กรุณาเสียบบัตรประชาชน...", "info")
                time.sleep(1)
            except Exception as e:
                logging.error(f"การเชื่อมต่อล้มเหลว: {str(e)}")
                self.update_gui("⚠️ การเชื่อมต่อผิดพลาด!", "error")
                time.sleep(1)


    def process_card(self, connection):
        """ประมวลผลข้อมูลบัตร"""
        self.card_processing = True
        try:
            self.update_gui("📖 กำลังอ่านข้อมูล...", "info")
            self.clear_text_area()
            
            # อ่านข้อมูล ATR
            atr = connection.getATR()
            logging.info(f"ATR: {toHexString(atr)}")
            
            # เลือกแอปพลิเคชัน
            if not self.select_application(connection, atr):
                raise Exception("ไม่พบแอปพลิเคชันที่รองรับ")
            
            # อ่านข้อมูลทั้งหมด
            for key in COMMANDS:
                data = self.execute_command(connection, key)
                if data:
                    self.append_text(f"{key}: {clean_text(thai2unicode(data))}\n")
            
            order_code = self.excel.get_last_row()
            order_code = str(int(order_code) + 1)

            cid = clean_text(thai2unicode(self.execute_command(connection, "CID")))
            is_duplicated = self.excel.check_duplicate("E", cid)

            self.update_gui(f"อ่านข้อมูลสำเร็จ", "success")
                    
            # ตรวจสอบข้อมูลซ้ำ
            if not is_duplicated:
                # อัปเดตไฟล์ Excel
                new_data = [
                    f"A{order_code}",
                    clean_text(thai2unicode(self.execute_command(connection, "THFULLNAME"))),
                    clean_text(thai2unicode(self.execute_command(connection, "ADDRESS"))),
                    "-",
                    clean_text(thai2unicode(self.execute_command(connection, "CID"))),
                    "",
                ]
                self.excel.add_data(new_data)
                messagebox.showinfo("สำเร็จ", "🟢 บันทึกข้อมูลสำเร็จ")
                time.sleep(1)
            else:
                messagebox.showwarning("แจ้งเตือน", "⚠️ มีข้อมูลนี้อยู่ในไฟล์แล้ว")
                time.sleep(1)
            
            # รอการถอดบัตร
            self.wait_for_card_removal(connection)
            
        except Exception as e:
            logging.error(f"การประมวลผลล้มเหลว: {str(e)}")
            self.update_gui(f"⚠️ ข้อผิดพลาด: {str(e)}", "error")
        finally:
            self.card_processing = False
            connection.disconnect()

    def select_application(self, connection, atr):
        """เลือกแอปพลิเคชันด้วย AID ที่เหมาะสม"""
        try:
            # ส่งคำสั่ง SELECT
            select_cmd = [0x00, 0xA4, 0x04, 0x0C, len(AID)] + AID
            connection.transmit(select_cmd)
            return True
            
        except Exception as e:
            logging.error(f"เลือก AID ล้มเหลว: {str(e)}")
        
        return False

    def execute_command(self, connection, command_key):
        """ดำเนินการคำสั่ง APDU"""
        try:
            cmd = COMMANDS[command_key]
            data, sw1, sw2 = connection.transmit(cmd)
            
            # จัดการกรณีต้องการอ่านข้อมูลเพิ่มเติม
            if sw1 == 0x61:
                get_response = [0x00, 0xC0, 0x00, 0x00, sw2]
                data, sw1, sw2 = connection.transmit(get_response)
            
            if sw1 != 0x90:
                raise Exception(f"คำสั่ง {command_key} ล้มเหลว (SW={hex(sw1)}{hex(sw2)})")
            
            return data
        except Exception as e:
            logging.error(f"คำสั่ง {command_key} ผิดพลาด: {str(e)}")
            return None

    def wait_for_card_removal(self, connection):
        """รอจนกว่าบัตรจะถูกถอดออก"""
        # self.update_gui(f"{self.status}\n\n กรุณาถอดบัตร...", self.status_color)
        while self.running:
            try:
                connection.transmit([0x00, 0x84, 0x00, 0x00, 0x00])
                time.sleep(0.5)
            except:
                self.update_gui("บัตรถูกถอดออกแล้ว", "info")
                return

    def update_gui(self, message, status_type="info"):
        """อัปเดตสถานะ GUI อย่างปลอดภัย"""
        color_map = {
            "info": "gray",
            "success": "#5cb85c",
            "error": "#d9534f",
            "warning": "#f0ad4e"
        }
        
        def _update():
            self.status_label.configure(
                text=message,
                text_color=color_map.get(status_type, "gray")
            )
        self.after(0, _update)

    def append_text(self, text):
        """เพิ่มข้อความลงใน Textbox"""
        def _append():
            self.text_area.configure(state="normal")
            self.text_area.insert("end", text)
            self.text_area.configure(state="disabled")
            self.text_area.see("end")
        self.after(0, _append)

    def clear_text_area(self):
        """ล้างข้อมูลใน Textbox"""
        def _clear():
            self.text_area.configure(state="normal")
            self.text_area.delete("1.0", "end")
            self.text_area.configure(state="disabled")
        self.after(0, _clear)

    def on_close(self):
        """จัดการการปิดโปรแกรม"""
        self.running = False
        if self.excel:
            try:
                self.excel.wb.save()
                self.excel.app.quit()
            except Exception as e:
                logging.error(f"ปิด Excel ไม่สำเร็จ: {str(e)}")

        self.destroy()

