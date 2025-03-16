import customtkinter as ctk
from tkinter import filedialog
from PIL import Image  # ใช้สำหรับโหลดไอคอน

ctk.set_appearance_mode("Dark")  # ตั้งค่าโหมดการแสดงผลเป็น Dark Mode
ctk.set_default_color_theme("blue")  # ใช้ธีมสีเขียว

class PathSelectorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Path Selector")
        self.geometry("450x200")
        self.grid_columnconfigure(1, weight=1)  # ให้คอลัมน์ที่ 1 ขยายตัวตามขนาดหน้าต่าง
        
        # Label และปุ่มเลือกไฟล์
        self.file_label = ctk.CTkLabel(self,text="เลือกไฟล์ :")
        self.file_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Entry สำหรับแสดงพาธ
        self.path_entry = ctk.CTkEntry(self, width=200, state="disabled")
        self.path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.file_button = ctk.CTkButton(self, text="เลือกไฟล์", command=self.select_file)
        self.file_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Label และ ComboBox สำหรับเลือกตัวเลือก
        self.combo_label = ctk.CTkLabel(self, text="เลือกตัวเลือก :")
        self.combo_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.combo_box = ctk.CTkComboBox(self,width=200, values=["A", "B", "C"], command=self.select_option)
        self.combo_box.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")

    def update_entry(self, path):
        """ อัปเดตค่าใน Entry และปิดการแก้ไข """
        self.path_entry.configure(state="normal")  # เปิดให้แก้ไขได้ชั่วคราว
        self.path_entry.delete(0, "end")  # ล้างค่าเก่า
        self.path_entry.insert(0, path)  # ใส่ค่าพาธใหม่
        self.path_entry.configure(state="disabled")  # ปิดการแก้ไขอีกครั้ง

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")]  # จำกัดให้เลือกไฟล์ Excel เท่านั้น
        )
        if file_path:
            self.update_entry(file_path)

    def select_option(self, choice):
        """ ฟังก์ชันที่ทำงานเมื่อเลือกค่าใน ComboBox """
        print(f"เลือกค่า: {choice}")

if __name__ == "__main__":
    app = PathSelectorApp()
    app.mainloop()
