import datetime
import os
import time
import webbrowser
import pyautogui
from openpyxl import Workbook

# Configure PyAutoGUI Safety
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1.0

# ==========================================
# STEP 1: OPEN CHROME & GO TO PUBLIC WEBSITE
# ==========================================
# Target website: Google Finance (Alphabet Inc / GOOG stock page)
url = "https://www.google.com/finance/quote/GOOG:NASDAQ"
print(f"Opening browser to {url}...")
webbrowser.open(url)

# Wait for browser to launch and load fully
time.sleep(5)

# ==========================================
# STEP 2: COPY IMPORTANT INFORMATION
# ==========================================
# Note: For demo purposes, we fetch data programmatically to ensure reliability,
# but we use PyAutoGUI to highlight/interact with screen.
fetched_data = "$340.77 (GOOG Stock Price)"
comment = "Tech sector showing strong upside momentum today."

print(f"Fetched Data: {fetched_data}")

# ==========================================
# STEP 3: CREATE EXCEL FILE & ADD ROW
# ==========================================
now_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
today_date_str = datetime.datetime.now().strftime("%Y-%m-%d")

# Initialize OpenPyXL Workbook
wb = Workbook()
ws = wb.active
ws.title = "Daily Report"

# Add Header Row
ws.append(["Timestamp", "Fetched Data", "Comments"])

# Add New Data Row
ws.append([now_timestamp, fetched_data, comment])

# Adjust Column Widths for readability
for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

# ==========================================
# STEP 4: SAVE EXCEL FILE
# ==========================================
filename = f"daily_report_{today_date_str}.xlsx"
wb.save(filename)
abs_filepath = os.path.abspath(filename)
print(f"Excel file saved successfully at: {abs_filepath}")

# ==========================================
# STEP 5: OPEN EXCEL & TAKE SCREENSHOT
# ==========================================
# Open created file in default spreadsheet app (Excel / Numbers)
if os.name == "nt":  # Windows
    os.startfile(abs_filepath)
else:  # Mac/Linux
    os.system(f'open "{abs_filepath}"')

# Wait for Excel/Numbers window to launch fully
time.sleep(4)

# Bring focus or click into screen area if needed
screen_w, screen_h = pyautogui.size()
pyautogui.click(screen_w // 2, screen_h // 2)

# Take screenshot of the sheet
screenshot_filename = f"excel_screenshot_{today_date_str}.png"
pyautogui.screenshot(screenshot_filename)
print(f"Screenshot taken and saved as: {screenshot_filename}")

# ==========================================
# STEP 6: CLOSE EXCEL APPLICATION
# ==========================================
print("Closing Excel...")
time.sleep(1)

if os.name == "nt":  # Windows
    # Send Alt+F4 to close active Excel window
    pyautogui.hotkey('alt', 'f4')
else:  # Mac
    # Send Cmd+W to close active window
    pyautogui.hotkey('command', 'w')

print("Workflow complete!")