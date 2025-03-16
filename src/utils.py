import logging
import re

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
