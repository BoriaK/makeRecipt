"""
diag_click_test.py
-------------------
Connects to the ALREADY-OPEN Foxit window (from diag_sig_tree.py) and tests
clicking the Signature List button via UIA click_input(), then screenshots
before/after and after moving to a document position, to see exactly what
happens step by step.
"""
import time, os
import pywinauto
import win32gui
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

shot('ct_0_initial.png')

sig_btn = [b for b in win.descendants(control_type='Button') if b.window_text() == 'Signature List']
if not sig_btn:
    print("Signature List not found now"); raise SystemExit(1)
br = sig_btn[0].rectangle()
print(f"Signature List rect: {br}")

print("Clicking via UIA click_input() ...")
sig_btn[0].click_input()
time.sleep(1.5)
shot('ct_1_after_uia_click.png')

# Check button states around ribbon after click
apply_btn = [b for b in win.descendants(control_type='Button') if b.window_text() == 'Apply All Signatures']
if apply_btn:
    try:
        print(f"Apply All Signatures enabled = {apply_btn[0].is_enabled()}")
    except Exception as e:
        print(f"enabled check error: {e}")

# Move mouse toward a plausible document position and screenshot the ghost/preview
doc_x = (wl + wr)//2
doc_y = wt + (wb-wt)//2
import pyautogui
pyautogui.moveTo(doc_x, doc_y, duration=0.4)
time.sleep(0.5)
shot('ct_2_hover_doc.png')

pyautogui.click()
time.sleep(1.0)
shot('ct_3_after_doc_click.png')

if apply_btn:
    try:
        print(f"Apply All Signatures enabled after doc click = {apply_btn[0].is_enabled()}")
    except Exception as e:
        print(f"enabled check error: {e}")

print("Done.")
