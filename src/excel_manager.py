import xlwings as xw
import logging
import pythoncom
import os

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG)

class ExcelManager:
    """จัดการการทำงานกับไฟล์ Excel โดยใช้ xlwings"""

    def __init__(self, file_path, sheet_name):
        # เรียกใช้ COM initialization ใน Thread นี้
        pythoncom.CoInitialize()
        self.file_path = file_path
        self.sheet_name = sheet_name

        # ตรวจสอบสิทธิ์การเขียนไฟล์ (อ่าน/เขียนได้)
        if not os.access(self.file_path, os.W_OK):
            logger.error(f"ไฟล์ {self.file_path} ถูกล็อกหรือเป็น Read-only")
            raise PermissionError(f"ไฟล์ {self.file_path} เป็นแบบอ่านอย่างเดียว")

        # สร้าง Excel Application (ไม่แสดงหน้าต่าง) และป้องกันการเปิด workbook ใหม่โดยอัตโนมัติ
        self.app = xw.App(visible=True, add_book=False)
        try:
            self.wb = self.app.books.open(self.file_path)
            self.sheet = self.wb.sheets[self.sheet_name]
        except Exception as e:
            logger.error(f"เปิดไฟล์ Excel ไม่สำเร็จ: {str(e)}")
            self.app.quit()
            raise

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.wb.save(self.file_path)
        except Exception as e:
            logger.error(f"บันทึกไฟล์ Excel ไม่สำเร็จ: {str(e)}")
        finally:
            self.app.quit()
            pythoncom.CoUninitialize()  # ปิดการใช้งาน COM เมื่อเสร็จงาน
        if exc_type is not None:
            raise exc_val

    def get_last_row(self):
        """หาหมายเลขแถวสุดท้ายในคอลัมน์ A"""
        try:
            return self.sheet.range('A1').end('down').row
        except Exception:
            return 1  # หากไม่มีข้อมูลเลย

    def add_data(self, data: list):
        """เพิ่มข้อมูลใหม่ลงใน Excel"""
        try:
            last_row = self.get_last_row() + 1
            self.sheet.range(f'A{last_row}').value = data
            self.wb.save(self.file_path)
        except Exception as e:
            logger.error(f"เพิ่มข้อมูลไม่สำเร็จ: {str(e)}")
            raise

    def check_duplicate(self, column: str, value: str) -> bool:
        """
        ตรวจสอบว่ามีข้อมูลที่ซ้ำกันในคอลัมน์ที่ระบุหรือไม่
        """
        return value in self.get_column_values(column)

    def get_column_values(self, column: str) -> list:
        """
        ดึงค่าจากคอลัมน์ที่ระบุใน Excel ทั้งหมดแล้วเก็บไว้ใน list
        โดยแปลงทุกค่าเป็นสตริงและตัด ".0" ออกหากมี

        Args:
            column (str): ชื่อคอลัมน์ (เช่น "A", "B", "E" เป็นต้น)

        Returns:
            list: รายการของค่าจากเซลล์ในคอลัมน์ที่ระบุ
        """
        try:
            last_row = self.get_last_row()
            if last_row < 1:
                return []
            cells = self.sheet.range(f"{column}1:{column}{last_row}")
            values = []
            for cell in cells:
                val = cell.value
                s = str(val) if val is not None else ""
                if s.endswith(".0"):
                    s = s[:-2]
                values.append(s.strip())
            return values
        except Exception as e:
            logger.error(f"ไม่สามารถดึงข้อมูลจากคอลัมน์ {column} ได้: {e}")
            return []
