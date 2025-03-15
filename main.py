from src.smartcard_app import SmartCardApp

if __name__ == "__main__":
    app = SmartCardApp()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
