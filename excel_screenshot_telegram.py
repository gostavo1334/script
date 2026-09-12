#!/usr/bin/env python3
"""
Excel Screenshot to Telegram Bot
- Opens specified Excel file and sheet
- Takes screenshot of the sheet
- Sends both screenshot AND Excel file to Telegram
- Can be scheduled to run at specific times
"""

import os
import sys
import time
import argparse
import requests
import win32com.client as win32
from datetime import datetime
import schedule

# ============================================================
# CONFIGURATION - EDIT THESE VALUES
# ============================================================

# Telegram Bot Configuration
BOT_TOKEN = "8992293068:AAEWvcFouV8xc0v9PDfnBsNwaxGgAViYlUQ"
CHAT_ID = "-5221976122"

# Fixed caption for all screenshots
FIXED_CAPTION ="PKD kính gửi danh sách khách hàng mới trong tháng 9 chưa có thông tin nhân viên  .                                                   នេះជាបញ្ជីឈ្មោះ អតិថិជនថ្មីក្នុងខែ កញ្ញា ដែលមិនទាន់មានពត៌មានរបស់បុគ្គលិកអ្នកណែនាំ។សូមបញ្ចូលព័ត៌មានអ្នកណែនាំ៕សូមអរគុណ៕"

# Default Excel file to open (optional)
# Set this to your Excel file path, or use --file argument
DEFAULT_EXCEL_FILE = None  # Example: r"C:\Users\DELL\Documents\report.xlsx"

# Default sheet name to capture (optional)
DEFAULT_SHEET_NAME = None  # Example: "Sheet1"

# Screenshot save location
SCREENSHOT_FOLDER = os.path.join(os.path.expanduser("~"), "Excel_Screenshots")

# ============================================================
# SCRIPT CODE - DO NOT EDIT BELOW THIS LINE
# ============================================================

TELEGRAM_API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"

# Create screenshot folder if it doesn't exist
os.makedirs(SCREENSHOT_FOLDER, exist_ok=True)


def find_excel_window():
    """Find the active Excel window handle."""
    
    def enum_callback(hwnd, result):
        class_name = win32gui.GetClassName(hwnd)
        window_text = win32gui.GetWindowText(hwnd)
        
        # Check for Excel window (class name contains "XL" or window text contains "Excel")
        if "XL" in class_name or "Excel" in window_text.lower():
            result.append(hwnd)
        return True
    
    windows = []
    win32gui.EnumWindows(enum_callback, windows)
    
    if windows:
        # Return the most recent/active Excel window
        return windows[-1]
    
    return None


def capture_excel_window():
    """Capture the Excel window and return as PIL Image."""
    hwnd = find_excel_window()
    
    if not hwnd:
        print("ERROR: No Excel window found!")
        print("Please open Excel and make sure a workbook is visible.")
        return None
    
    # Get window dimensions
    rect = win32gui.GetWindowRect(hwnd)
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top
    
    print(f"Found Excel window: {left}x{top} to {right}x{bottom} (size: {width}x{height})")
    
    # Capture the window
    screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
    
    return screenshot


def save_screenshot(image, timestamp=None):
    """Save screenshot to folder for debugging."""
    if timestamp is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    filename = os.path.join(SCREENSHOT_FOLDER, f"excel_screenshot_{timestamp}.png")
    image.save(filename)
    print(f"Screenshot saved: {filename}")
    return filename


def send_to_telegram(image_path, caption):
    """Send screenshot to Telegram."""
    try:
        with open(image_path, 'rb') as photo:
            files = {'photo': photo}
            data = {'chat_id': CHAT_ID, 'caption': caption}
            
            response = requests.post(TELEGRAM_API_URL, files=files, data=data)
            result = response.json()
            
            if result.get('ok'):
                print(f"SUCCESS: Screenshot sent to Telegram!")
                print(f"Message ID: {result['result']['message_id']}")
                return True
            else:
                print(f"ERROR: Telegram API error: {result.get('description', 'Unknown error')}")
                return False
    except Exception as e:
        print(f"ERROR: Failed to send to Telegram: {e}")
        return False


def take_and_send_screenshot(caption=None):
    """Main function: capture Excel and send to Telegram."""
    if caption is None:
        caption = FIXED_CAPTION
    
    print("=" * 60)
    print(f"Taking Excel screenshot at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Caption: {caption}")
    print("-" * 60)
    
    # Capture Excel window
    screenshot = capture_excel_window()
    
    if screenshot is None:
        return False
    
    # Save temporarily
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_path = save_screenshot(screenshot, timestamp)
    
    # Send to Telegram
    success = send_to_telegram(temp_path, caption)
    
    # Optionally: delete the local file after sending
    # os.remove(temp_path)
    
    return success


def schedule_screenshot(time_str, caption=None):
    """Schedule a screenshot at a specific time."""
    print(f"Scheduling screenshot at: {time_str}")
    schedule.every().day.at(time_str).do(take_and_send_screenshot, caption=caption)
    
    print(f"Scheduled! Waiting to run at {time_str}...")
    
    # Keep the script running
    while True:
        schedule.run_pending()
        time.sleep(1)


def immediate_screenshot(caption=None):
    """Take and send screenshot immediately."""
    success = take_and_send_screenshot(caption)
    if success:
        print("\nDone! Check your Telegram group.")
    else:
        print("\nFailed to send screenshot.")
    sys.exit(0 if success else 1)


def main():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Excel Screenshot to Telegram Bot",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python excel_screenshot_telegram.py                    # Send immediately
  python excel_screenshot_telegram.py --now              # Send immediately
  python excel_screenshot_telegram.py --time "15:30"    # Schedule at 3:30 PM
  python excel_screenshot_telegram.py --time "09:00" --caption "Daily Report"
        """
    )
    
    parser.add_argument(
        '--now',
        action='store_true',
        help='Take and send screenshot immediately'
    )
    
    parser.add_argument(
        '--time',
        type=str,
        default=None,
        help='Time to schedule screenshot (format: "HH:MM" or "HH:MM:SS")'
    )
    
    parser.add_argument(
        '--caption',
        type=str,
        default=None,
        help='Caption for the screenshot (overrides FIXED_CAPTION in script)'
    )
    
    parser.add_argument(
        '--set-caption',
        type=str,
        default=None,
        help='Permanently set the fixed caption in the script'
    )
    
    args = parser.parse_args()
    
    # Handle permanent caption setting
    if args.set_caption:
        update_script_caption(args.set_caption)
        return
    
    # Validate time format if provided
    if args.time:
        try:
            # Validate time format
            time_obj = datetime.strptime(args.time, "%H:%M:%S")
        except ValueError:
            try:
                time_obj = datetime.strptime(args.time + ":00", "%H:%M:%S")
            except ValueError:
                print("ERROR: Invalid time format. Use HH:MM or HH:MM:SS")
                sys.exit(1)
        
        schedule_screenshot(args.time, args.caption)
    else:
        # Default to immediate
        immediate_screenshot(args.caption)


def update_script_caption(new_caption):
    """Update the FIXED_CAPTION in the script file."""
    script_path = __file__
    
    with open(script_path, 'r') as f:
        content = f.read()
    
    # Replace the FIXED_CAPTION line
    old_line = f'FIXED_CAPTION = "{FIXED_CAPTION}"'
    new_line = f'FIXED_CAPTION = "{new_caption}"'
    
    if old_line in content:
        content = content.replace(old_line, new_line)
        
        with open(script_path, 'w') as f:
            f.write(content)
        
        print(f"SUCCESS: Fixed caption updated to: '{new_caption}'")
    else:
        print("ERROR: Could not find FIXED_CAPTION line to update")
        sys.exit(1)


if __name__ == "__main__":
    # Check for required dependencies
    try:
        import win32gui
        import win32con
        from PIL import ImageGrab
        import requests
        import schedule
    except ImportError as e:
        print(f"ERROR: Missing required package: {e}")
        print("\nPlease install dependencies by running:")
        print("  pip install pillow requests pywin32 schedule")
        sys.exit(1)
    
    main()
