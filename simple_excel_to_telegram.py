#!/usr/bin/env python3
"""
SIMPLE Excel to Telegram Bot
- Opens Excel file
- Analyzes branch data from the sheet
- Goes to your sheet
- Copies the used range as a picture
- Saves from clipboard as PNG
- Sends to Telegram with dynamic caption
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


def analyze_branch_data(ws, range_addr=None):
    """Read Excel data and categorize branches by performance."""
    try:
        if range_addr:
            rng = ws.Range(range_addr)
        else:
            rng = ws.Range("A1:Z200")
        
        used_range = ws.UsedRange
        data = []
        
        for row in range(1, used_range.Rows.Count + 1):
            row_data = []
            for col in range(1, used_range.Columns.Count + 1):
                cell_value = used_range.Cells(row, col).Value
                if hasattr(cell_value, 'strftime'):
                    cell_value = cell_value.strftime('%Y-%m-%d')
                row_data.append(cell_value if cell_value is not None else "")
            data.append(row_data)
        
        # Hardcoded columns per user: Branch=C(2), New Customers=D(3), Referral=E(4), Rating=I(8)
        branch_col_idx = 2
        new_cust_col_idx = 3
        referral_col_idx = 4
        rating_col_idx = 8
        
        print(f"[ANALYSIS] Using: Branch=C({branch_col_idx}), NewCust=D({new_cust_col_idx}), Ref=E({referral_col_idx}), Rating=I({rating_col_idx})")
        
        branches = {}
        data_started = False
        for row_idx in range(len(data)):
            row = data[row_idx]
            if len(row) <= max(branch_col_idx, new_cust_col_idx, referral_col_idx, rating_col_idx):
                continue
            branch = str(row[branch_col_idx]).strip()
            if not branch:
                continue
            if branch.lower() in ['total', 'grand total', 'grand', '', 'nan', 'none', '#n/a']:
                continue
            
            if not data_started:
                data_started = True
                print(f"[ANALYSIS] Data starts at row {row_idx}")
            
            def parse_num(val):
                if val is None:
                    return 0
                try:
                    val_str = str(val).replace(',', '').replace('%', '').replace(' ', '').strip()
                    if not val_str or val_str.lower() in ['nan', 'none', 'null', '#n/a']:
                        return 0
                    return float(val_str)
                except:
                    return 0
            
            new_customers = int(parse_num(row[new_cust_col_idx]))
            referrals = int(parse_num(row[referral_col_idx]))
            rating = parse_num(row[rating_col_idx])
            
            branches[branch] = {'new_customers': new_customers, 'referrals': referrals, 'rating': rating}
            
            if len(branches) <= 3:
                print(f"[ANALYSIS] Added: {branch} (cust={new_customers}, ref={referrals}, rating={rating})")
        
        if not data_started:
            print("[ANALYSIS] ERROR: No data found!")
            return {'low_performers': [], 'no_results': [], 'low_referral': []}
        
        print(f"[ANALYSIS] Found {len(branches)} branches: {list(branches.keys())}")
        
        # 1. Low Performers: 4 branches with LOWEST rating from column I
        sorted_by_rating = sorted(branches.items(), key=lambda x: x[1]['rating'])
        low_performers = [b for b, d in sorted_by_rating[:4]]
        
        # 2. No Results: new_customers == 0
        no_results = [b for b, d in branches.items() if d['new_customers'] == 0]
        
        # 3. Low Referral: referrals == 0 AND new_customers >= 3
        low_referral = [b for b, d in branches.items() if d['referrals'] == 0 and d['new_customers'] >= 3]
        
        print(f"[ANALYSIS] Low performers (4 lowest rating): {low_performers}")
        print(f"[ANALYSIS] No results (cust==0): {no_results}")
        print(f"[ANALYSIS] Low referral (ref==0 AND cust>=3): {low_referral}")
        
        return {
            'low_performers': low_performers,
            'no_results': no_results,
            'low_referral': low_referral
        }
    except Exception as e:
        print(f"[ANALYSIS ERROR] {e}")
        return {'low_performers': [], 'no_results': [], 'low_referral': []}


def get_caption(branch_data=None):
    """Generate caption with dynamic branch data."""
    today = datetime.now().strftime("%d/%m")
    
    if branch_data is None:
        branch_data = {'low_performers': [], 'no_results': [], 'low_referral': []}
    
    low_perf = ", ".join(f"{b}" for b in branch_data['low_performers']) if branch_data['low_performers'] else "None"
    no_res = ", ".join(f"{b}" for b in branch_data['no_results']) if branch_data['no_results'] else "None"
    low_ref = ", ".join(f"{b}" for b in branch_data['low_referral']) if branch_data['low_referral'] else "None"
    
    return f"""DAILY REPORT OF NEW CUSOMERS


នេះជាបញ្ជីឈ្មោះ អតិថិជនថ្មីក្នុងខែ កញ្ញា ដែលមិនទាន់មានពត៌មានរបស់បុគ្គលិកអ្នកណែនាំ
សូមបញ្ចូលព័ត៌មានអ្នកណែនាំ

*សាខាមានលទ្ធផលភ្ញៀវថ្មីតិចជាងគេ: {low_perf}
*សាខាមិនទាន់មានលទ្ធផលភ្ញៀវថ្មី({today}): {no_res}
*សាខាមានលទ្ធផលភ្ញៀវថ្មីច្រើនដែលមានអតិថិជនត្រូវតាមរៀបលាក្រមួយ: {low_ref}

===========================================================================

PKD kính gửi danh sách khách hàng mới trong tháng 9 chưa có thông tin nhân viên .

* Chi nhánh hoàn thành KH mới thấp nhất: {low_perf}
*Chi nhánh không có kết quả KH mới ngày ({today}): {no_res}
*Chi nhánh có nhiều KH mới có tỉ lệ giới thiệu thấp: {low_ref}

Thanks."""


FOLDER = os.path.join(os.path.expanduser("~"), "Excel_Screenshots")
os.makedirs(FOLDER, exist_ok=True)

URL_PHOTO = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"
URL_DOC = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"


def open_excel(file_path, sheet_name):
    """Open Excel file and sheet. Returns (excel, ws, actual_file_path)."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    
    try:
        actual_file_path = None
        if file_path and os.path.exists(file_path):
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            actual_file_path = os.path.abspath(file_path)
        elif excel.Workbooks.Count > 0:
            wb = excel.ActiveWorkbook
            actual_file_path = os.path.abspath(wb.FullName)
        else:
            print("ERROR: No Excel file open")
            excel.Quit()
            return None, None, None
        
        if sheet_name:
            try:
                ws = wb.Worksheets(sheet_name)
            except:
                ws = wb.ActiveSheet
                print(f"Sheet '{sheet_name}' not found, using active sheet")
        else:
            ws = wb.ActiveSheet
        
        return excel, ws, actual_file_path
    except Exception as e:
        print(f"ERROR: {e}")
        excel.Quit()
        return None, None, None


def copy_range_as_image(ws, range_addr=None):
    """Copy range to clipboard as image."""
    try:
        if range_addr:
            rng = ws.Range(range_addr)
        else:
            rng = ws.Range("B1:M49")
        
        print(f"Copying range: {rng.Address}")
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
            dib = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            bmp_path = os.path.join(FOLDER, f"temp_{timestamp}.bmp")
            with open(bmp_path, 'wb') as f:
                f.write(dib)
            
            win32clipboard.CloseClipboard()
            
            try:
                from PIL import Image
                img = Image.open(bmp_path)
                png_path = os.path.join(FOLDER, f"excel_{timestamp}.png")
                img.save(png_path, 'PNG')
                os.remove(bmp_path)
                print(f"Saved: {png_path}")
                return png_path
            except:
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
    parser.add_argument('--no-analysis', action='store_true', help='Skip branch analysis')
    parser.add_argument('--time', help='Schedule at HH:MM')
    args = parser.parse_args()
    
    excel = None
    ws = None
    file_to_send = None
    
    try:
        print("Opening Excel...")
        excel, ws, file_to_send = open_excel(args.file, args.sheet)
        if not ws:
            return
        
        # Analyze data
        if not args.no_analysis and not args.caption:
            print("\n" + "="*60)
            print("ANALYZING EXCEL DATA...")
            print("="*60)
            branch_data = analyze_branch_data(ws, args.range)
            print("\n" + "="*60)
            print("ANALYSIS RESULTS:")
            print("="*60)
            print(f"  Low Performers:  {branch_data['low_performers']}")
            print(f"  No Results:      {branch_data['no_results']}")
            print(f"  Low Referral:    {branch_data['low_referral']}")
            print("="*60 + "\n")
            args.caption = get_caption(branch_data)
        elif args.caption is None:
            args.caption = get_caption()
        
        time.sleep(1)
        
        print("Copying range as image...")
        if not copy_range_as_image(ws, args.range):
            return
        
        time.sleep(1)
        
        print("Saving clipboard image...")
        image_path = save_clipboard_image()
        if not image_path:
            print("ERROR: No image captured")
            return
        
        print("Sending to Telegram...")
        send_to_telegram(image_path, args.caption, is_image=True)
        
        if not args.no_file and file_to_send and os.path.exists(file_to_send):
            send_to_telegram(file_to_send, "Excel File: " + os.path.basename(file_to_send), is_image=False)
        
        print("\nDone! Check Telegram.")
        
    finally:
        if excel:
            # Keep Excel open
            print("[Note] Excel kept open")


if __name__ == "__main__":
    try:
        import requests, schedule, win32com.client, win32clipboard, PIL
    except ImportError:
        print("Install: pip install requests schedule pywin32 pillow")
        sys.exit(1)
    main()
