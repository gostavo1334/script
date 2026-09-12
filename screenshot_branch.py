#!/usr/bin/env python3
"""
SIMPLE: Screenshot Branch Sheet and Send to Telegram
- Opens your Excel file
- Goes to Branch sheet
- Takes screenshot of Excel window
- Sends to Telegram
"""

import os
import sys
import time
import requests
import win32com.client
import win32gui
from PIL import ImageGrab
from datetime import datetime

# ====================
# CONFIGURATION
# ====================
BOT_TOKEN = "8992293068:AAEWvcFouV8xc0v9PDfnBsNwaxGgAViYlUQ"
CHAT_ID = "-5221976122"

def get_caption():
    ts = datetime.now().strftime("(%d/%m):%H:%M")
    return f"""នេះជាបញ្ជីឈ្មោះ អតិថិជនថ្មីក្នុងខែ កញ្ញា ដែលមិនទាន់មានពត៌មានរបស់បុគ្គលិកអ្នកណែនាំ
សូមបញ្ចូលព័ត៌មានអ្នកណែនាំ

*សាខាមានលទ្ធផលភ្ញៀវថ្មីតិចជាងគេ

*សាខាមិនទាន់មានលទ្ធផលភ្ញៀវថ្មី{ts}

*សាខាមានលទ្ធផលភ្ញៀវថ្មីច្រើនដែលមិនទាន់មានពត៌មានអ្នកណែនាំ

PKD kính gửi danh sách khách hàng mới trong tháng 9 chưa có thông tin nhân viên .

* Chi nhánh hoàn thành  KH mới thấp nhất:
+

*Chi nhánh không có kết quả KH mới ngày {ts}:

*Chi nhánh có nhiều KH mới chưa nhâp mã giới thiệu: 
+PNPP012

Thanks."""
FOLDER = os.path.join(os.path.expanduser("~"), "Excel_Screenshots")
os.makedirs(FOLDER, exist_ok=True)

TELEGRAM_URL = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"


def main():
    # Your file and sheet
    file_path = r"C:\Users\DELL\Desktop\report\sendtotele\New_Customer Sept  Final_11.09l.xlsx"
    sheet_name = "Branch"
    
    print("Starting...")
    
    # Open Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Must be visible for screenshot
    excel.DisplayAlerts = False
    
    try:
        print(f"Opening: {file_path}")
        workbook = excel.Workbooks.Open(file_path)
        
        print(f"Going to sheet: {sheet_name}")
        try:
            sheet = workbook.Worksheets(sheet_name)
            sheet.Activate()
        except:
            sheet = workbook.ActiveSheet
            print(f"Using active sheet instead")
        
        # Maximize window
        excel.WindowState = -4137  # xlMaximized
        
        # Wait for Excel to render
        time.sleep(2)
        
        # Find Excel window
        def find_excel_window():
            def callback(hwnd, result):
                class_name = win32gui.GetClassName(hwnd)
                if "XL" in class_name:
                    result.append(hwnd)
                return True
            
            windows = []
            win32gui.EnumWindows(callback, windows)
            return windows[-1] if windows else None
        
        hwnd = find_excel_window()
        if not hwnd:
            print("ERROR: Excel window not found")
            return
        
        # Get window position
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        
        print(f"Capturing: {left}x{top} to {right}x{bottom}")
        
        # Take screenshot
        screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
        
        # Save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        image_path = os.path.join(FOLDER, f"branch_{timestamp}.png")
        screenshot.save(image_path)
        print(f"Saved: {image_path}")
        
        # Send to Telegram
        print("Sending to Telegram...")
        with open(image_path, 'rb') as f:
            response = requests.post(
                TELEGRAM_URL,
                files={'photo': f},
                data={'chat_id': CHAT_ID, 'caption': get_caption()}
            )
        
        result = response.json()
        if result.get('ok'):
            print("SUCCESS: Screenshot sent to Telegram!")
        else:
            print(f"ERROR: {result.get('description', 'Unknown error')}")
        
    finally:
        # Close Excel
        try:
            workbook.Close(False)
            excel.Quit()
        except:
            pass
        print("Done!")


if __name__ == "__main__":
    try:
        import requests, win32com.client, win32gui, PIL
    except ImportError:
        print("Install: pip install requests pywin32 pillow")
        sys.exit(1)
    main()
