"""
diag_click_test2.py
--------------------
Reuses the already-open Foxit window. Cancels any open dialog, then clicks
precisely on the signature ink thumbnail (not the container edge), verifies
a selection highlight appears, then clicks on the document and checks result.
"""
import time, os
import pywinauto
import win32gui
import pyautogui
from PIL import ImageGrab

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'
DBG = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'debug')
os.makedirs(DBG, exist_ok=True)

app = pywinauto.Application(backend='uia').connect(path=FOXIT, timeout=10)
win = app.top_window()
hwnd = win.handle
win32gui.BringWindowToTop(hwnd)
time.sleep(0.5)
wl, wt, wr, wb = win32gui.GetWindowRect(hwnd)
print(f"Window rect: ({wl},{wt},{wr},{wb})")

def shot(name):
    p = os.path.join(DBG, name)
    ImageGrab.grab(bbox=(max(0,wl), max(0,wt), wr, wb), all_screens=True).save(p)
    print(f"Saved {p}")
    return p

# Cancel any open dialog (e.g. leftover Create Signature dialog)
try:
    dlg = app.window(class_name='#32770', top_level_only=True)
    if dlg.exists(timeout=1):
        print("Dialog found, pressing Escape")
        dlg.type_keys('{ESC}')
        time.sleep(1.0)
except Exception as e:
    print(f"dialog check: {e}")
# Also try Escape on the main window in case it's an in-app modal, not a real dialog
win.type_keys('{ESC}')
time.sleep(1.0)
shot('ct2_0_reset.png')

sig_btn = [b for b in win.descendants(control_type='Button') if b.window_text() == 'Signature List']
if not sig_btn:
    print("Signature List not found now"); raise SystemExit(1)
br = sig_btn[0].rectangle()
print(f"Signature List rect: {br}")

# Click precisely on the ink thumbnail: left portion of the container,
# roughly 25% across horizontally, vertically centered.
thumb_x = br.left + int((br.right - br.left) * 0.20)
thumb_y = (br.top + br.bottom) // 2
print(f"Clicking thumbnail ink at ({thumb_x},{thumb_y})")
pyautogui.moveTo(thumb_x, thumb_y, duration=0.3)
time.sleep(0.3)
pyautogui.click()
time.sleep(1.0)
shot('ct2_1_after_thumb_click.png')

apply_btn = [b for b in win.descendants(control_type='Button') if b.window_text() == 'Apply All Signatures']
if apply_btn:
    try:
        print(f"Apply All Signatures enabled = {apply_btn[0].is_enabled()}")
    except Exception as e:
        print(f"enabled check error: {e}")

doc_x = (wl + wr)//2
doc_y = wt + (wb-wt)//2
print(f"Moving to doc position ({doc_x},{doc_y}) and clicking")
pyautogui.moveTo(doc_x, doc_y, duration=0.4)
time.sleep(0.4)
shot('ct2_2_hover_doc.png')
pyautogui.click()
time.sleep(1.0)
shot('ct2_3_after_doc_click.png')

if apply_btn:
    try:
        print(f"Apply All Signatures enabled after doc click = {apply_btn[0].is_enabled()}")
    except Exception as e:
        print(f"enabled check error: {e}")

# check if a dialog popped up (Create Signature would mean thumb click didn't arm it)
try:
    dlg = app.window(class_name='#32770', top_level_only=True)
    print(f"Modal dialog present: {dlg.exists(timeout=1)}")
except Exception as e:
    print(f"dialog post-check error: {e}")

print("Done.")
