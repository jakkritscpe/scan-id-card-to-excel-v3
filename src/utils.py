import logging
import re
import xlwings as xw
import pythoncom
import os

logger = logging.getLogger(__name__)

def thai2unicode(data):
    """แปลงข้อมูล TIS-620 เป็น Unicode พร้อมจัดการข้อผิดพลาด"""
    try:
        return bytes(data).decode('tis-620').strip()
    except UnicodeDecodeError as e:
        logging.error(f"การแปลงอักขระล้มเหลว: {str(e)}")
        return bytes(data).decode('tis-620', errors='replace').strip()
    
def clean_text(text):
    """ล้างข้อความจากอักขระพิเศษ"""
    return re.sub(r'[#\x00]+', ' ', text).strip()


def open_excel(file_path, sheet_name, is_visible=True):
    """เปิดไฟล์ Excel และคืนค่า (app, workbook, sheet)"""
    pythoncom.CoInitialize()  # แก้ปัญหา COM Access
    app = xw.App(visible=is_visible)
    try:
        wb = app.books.open(file_path)
        sheet = wb.sheets[sheet_name]
        return app, wb, sheet
    except Exception as e:
        app.quit()
        logger.error(f"เปิดไฟล์ Excel ไม่สำเร็จ: {str(e)}")
        raise


def close_excel(app, wb, file_path):
    """ปิดไฟล์ Excel และบันทึกการเปลี่ยนแปลง"""
    try:
        wb.save(file_path)
    except Exception as e:
        logger.error(f"บันทึกไฟล์ Excel ไม่สำเร็จ: {str(e)}")
    finally:
        app.quit()
        pythoncom.CoUninitialize()  # ปิดการใช้งาน COM


def get_last_row(sheet, column="A"):
    """หาหมายเลขแถวสุดท้ายในคอลัมน์ที่กำหนด"""
    try:
        return sheet.range(f"{column}1").end('down').row
    except Exception:
        return 1  # หากไม่มีข้อมูลเลย


def add_data(file_path, sheet_name, data, column="A"):
    """เพิ่มข้อมูลใหม่ลงใน Excel"""
    app, wb, sheet = open_excel(file_path, sheet_name)
    try:
        last_row = get_last_row(sheet, column) + 1
        sheet.range(f"{column}{last_row}").value = data
        wb.save(file_path)
    except Exception as e:
        logger.error(f"เพิ่มข้อมูลไม่สำเร็จ: {str(e)}")
        raise
    finally:
        close_excel(app, wb, file_path)


def check_duplicate(file_path, sheet_name, column, value):
    """ตรวจสอบว่ามีข้อมูลที่ซ้ำกันในคอลัมน์ที่ระบุหรือไม่"""
    values = get_column_values(file_path, sheet_name, column)
    return value in values


def get_column_values(file_path, sheet_name, column):
    """ดึงค่าจากคอลัมน์ที่ระบุใน Excel ทั้งหมดและเก็บไว้ใน list"""
    app, wb, sheet = open_excel(file_path, sheet_name)
    try:
        last_row = get_last_row(sheet, column)
        if last_row < 1:
            return []
        cells = sheet.range(f"{column}1:{column}{last_row}")
        values = []
        for cell in cells:
            val = cell.value
            s = str(val) if val is not None else ""
            if s.endswith(".0"):  # ตัด ".0" ถ้าเป็นตัวเลข
                s = s[:-2]
            values.append(s.strip())
        return values
    except Exception as e:
        logger.error(f"ไม่สามารถดึงข้อมูลจากคอลัมน์ {column} ได้: {e}")
        return []
    finally:
        close_excel(app, wb, file_path)