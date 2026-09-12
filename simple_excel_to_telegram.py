#!/usr/bin/env python3
"""
SIMPLE Excel to Telegram Bot
- Opens Excel file
- Goes to your sheet
- Copies the used range as a picture
- Saves from clipboard as PNG
- Sends to Telegram
"""

import os
import sys
import time
import argparse
import requests
import win32com.client
import win32clipboard
from datetime import datetime
import schedule
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

*Chi nhánh không có kết quả KH mới ngày {ts}

*Chi nhánh có nhiều KH mới chưa nhâp mã giới thiệu: 
+PNPP012

Thanks."""

FOLDER = os.path.join(os.path.expanduser("~"), "Excel_Screenshots")
os.makedirs(FOLDER, exist_ok=True)

URL_PHOTO = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"
URL_DOC = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"


def open_excel(file_path, sheet_name):
    """Open Excel file and sheet."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    
    try:
        if file_path and os.path.exists(file_path):
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
        elif excel.Workbooks.Count > 0:
            wb = excel.ActiveWorkbook
        else:
            print("ERROR: No Excel file open")
            excel.Quit()
            return None, None
        
        if sheet_name:
            try:
                ws = wb.Worksheets(sheet_name)
            except:
                ws = wb.ActiveSheet
                print(f"Sheet '{sheet_name}' not found, using active sheet")
        else:
            ws = wb.ActiveSheet
        
        return excel, ws
    except Exception as e:
        print(f"ERROR: {e}")
        excel.Quit()
        return None, None


def copy_range_as_image(ws, range_addr=None):
    """Copy range to clipboard as image."""
    try:
        if range_addr:
            rng = ws.Range(range_addr)
        else:
            rng = ws.Range("B1:M49")
        
        print(f"Copying range: {rng.Address}")
        # Copy as picture (Appearance=1=xlScreen, Format=2=xlPicture)
        rng.CopyPicture(Appearance=1, Format=2)
        return True
    except Exception as e:
        print(f"ERROR copying: {e}")
        return False


def save_clipboard_image():
    """Save clipboard image to PNG file."""
    try:
        win32clipboard.OpenClipboard()
        
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            # Get DIB data
            dib = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            
            # Write to temporary BMP
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            bmp_path = os.path.join(FOLDER, f"temp_{timestamp}.bmp")
            with open(bmp_path, 'wb') as f:
                f.write(dib)
            
            win32clipboard.CloseClipboard()
            
            # Convert BMP to PNG
            try:
                from PIL import Image
                img = Image.open(bmp_path)
                png_path = os.path.join(FOLDER, f"excel_{timestamp}.png")
                img.save(png_path, 'PNG')
                os.remove(bmp_path)
                print(f"Saved: {png_path}")
                return png_path
            except:
                # If conversion fails, return BMP
                print(f"Saved: {bmp_path}")
                return bmp_path
        else:
            win32clipboard.CloseClipboard()
            print("No image in clipboard")
            return None
    except Exception as e:
        try:
            win32clipboard.CloseClipboard()
        except:
            pass
        print(f"Clipboard error: {e}")
        return None


def send_to_telegram(file_path, caption, is_image=True):
    """Send file to Telegram."""
    try:
        url = URL_PHOTO if is_image else URL_DOC
        with open(file_path, 'rb') as f:
            files = {'photo': f} if is_image else {'document': f}
            data = {'chat_id': CHAT_ID, 'caption': caption}
            r = requests.post(url, files=files, data=data)
        result = r.json()
        if result.get('ok'):
            print(f"Sent to Telegram: {file_path}")
            return True
        else:
            print(f"Telegram error: {result.get('description')}")
            return False
    except Exception as e:
        print(f"Send error: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Send Excel range to Telegram")
    parser.add_argument('--file', help='Excel file path')
    parser.add_argument('--sheet', help='Sheet name')
    parser.add_argument('--range', help='Range like A1:L27')
    parser.add_argument('--caption', help='Caption', default=None)
    parser.add_argument('--no-file', action='store_true', help='Send only image')
    parser.add_argument('--time', help='Schedule at HH:MM')
    args = parser.parse_args()
    if args.caption is None:
        args.caption = get_caption()
    
    excel = None
    ws = None
    try:
        print("Opening Excel...")
        excel, ws = open_excel(args.file, args.sheet)
        if not ws:
            return
        
        time.sleep(1)
        
        print("Copying range as image...")
        if not copy_range_as_image(ws, args.range):
            return
        
        time.sleep(1)  # Wait for clipboard
        
        print("Saving clipboard image...")
        image_path = save_clipboard_image()
        if not image_path:
            print("ERROR: No image captured")
            return
        
        print("Sending to Telegram...")
        send_to_telegram(image_path, args.caption, is_image=True)
        
        if not args.no_file and args.file and os.path.exists(args.file):
            send_to_telegram(args.file, args.caption, is_image=False)
        
        print("\nDone! Check Telegram.")
        
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass


if __name__ == "__main__":
    try:
        import requests, schedule, win32com.client, win32clipboard, PIL
    except ImportError:
        print("Install: pip install requests schedule pywin32 pillow")
        sys.exit(1)
    main()
