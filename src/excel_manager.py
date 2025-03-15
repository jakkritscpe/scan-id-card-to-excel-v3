import xlwings as xw
import logging

logger = logging.getLogger(__name__)

class ExcelManager:
    """จัดการการทำงานกับไฟล์ Excel โดยใช้ xlwings"""
    
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.app = xw.App(visible=False)
        try:
            self.wb = self.app.books.open(self.file_path)
            self.sheet = self.wb.sheets[self.sheet_name]
        except Exception as e:
            logger.error(f"เปิดไฟล์ Excel ไม่สำเร็จ: {str(e)}")
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

            # ระบุ path ของไฟล์ให้กับเมธอด save เพื่อหลีกเลี่ยงข้อผิดพลาด timeout
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
                # แปลงค่าเป็นสตริง ถ้าเป็น None ให้เป็นค่าว่าง
                s = str(val) if val is not None else ""
                # หากค่าสิ้นสุดด้วย ".0" ให้ตัดออก
                if s.endswith(".0"):
                    s = s[:-2]
                values.append(s.strip())
            return values
        except Exception as e:
            logger.error(f"ไม่สามารถดึงข้อมูลจากคอลัมน์ {column} ได้: {e}")
            return []



