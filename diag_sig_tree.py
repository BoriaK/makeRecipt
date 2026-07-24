"""
diag_sig_tree.py
----------------
One-off diagnostic: opens a PDF in Foxit, navigates to Fill & Sign, and dumps
the FULL UIA control tree under the Signature List ribbon area (control_type,
name, automation_id, rect) so we can find the real clickable element for the
saved signature thumbnail instead of guessing pixel coordinates.
"""
import sys, time, os, subprocess
import pywinauto
import win32gui, win32con

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'

pdf = sys.argv[1] if len(sys.argv) > 1 else \
      r'C:\Users\Administrator\Documents\קבלות_על_אימונים\recipt_28_June_2026.pdf'

print(f"Opening {pdf} ...")
subprocess.Popen([FOXIT, pdf])
pdf_name = os.path.basename(pdf)

app = None
for _ in range(30):
    try:
        a = pywinauto.Application(backend='uia').connect(path=FOXIT, timeout=3)
        for w in a.windows():
            if pdf_name in w.window_text():
                app = a
                break
        if app:
            break
    except Exception:
        pass
    time.sleep(1)
if not app:
    print("Foxit did not open in time"); sys.exit(1)
print("PDF loaded")

win = app.top_window()
hwnd = win.handle
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
time.sleep(0.3)
win32gui.MoveWindow(hwnd, 2500, 200, 800, 600, True)
time.sleep(0.3)
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
win32gui.BringWindowToTop(hwnd)
time.sleep(2.5)

win = app.top_window()
fs = [t for t in win.descendants(control_type='TabItem') if t.window_text() == 'Fill & Sign']
if fs:
    fs[0].click_input()
    print("Fill & Sign tab clicked")
else:
    print("Fill & Sign TAB not found. Available TabItems:")
    for t in win.descendants(control_type='TabItem'):
        print(f"  '{t.window_text()}'")
    btn = [b for b in win.descendants(control_type='Button')
           if 'Fill' in b.window_text() and 'Sign' in b.window_text()]
    if btn:
        btn[0].click_input()
        print(f"Fill & Sign BUTTON clicked: [{btn[0].window_text()}]")
    else:
        print("Fill & Sign BUTTON not found either. Available Buttons:")
        for b in win.descendants(control_type='Button'):
            print(f"  '{b.window_text()}'")
time.sleep(2.0)

win = app.top_window()
sig_btn = None
for attempt in range(12):
    win = app.top_window()
    sig_btn = [b for b in win.descendants(control_type='Button')
               if b.window_text() == 'Signature List']
    if sig_btn:
        break
    time.sleep(1.0)
if not sig_btn:
    print("Signature List button not found. Available Buttons:")
    for b in win.descendants(control_type='Button'):
        try:
            print(f"  '{b.window_text()}'  rect={b.rectangle()}")
        except Exception:
            pass
    sys.exit(1)

br = sig_btn[0].rectangle()
print(f"Signature List rect: {br}")

print("\n--- Direct children of 'Signature List' ---")
try:
    for c in sig_btn[0].children():
        try:
            cr = c.rectangle()
            print(f"  [{c.element_info.control_type}] name='{c.window_text()}' "
                  f"automation_id='{c.element_info.automation_id}' rect={cr}")
        except Exception as e:
            print(f"  (error reading child: {e})")
except Exception as e:
    print(f"children() error: {e}")

print("\n--- ALL descendants in ribbon area (y within Signature List rect +/- 20) ---")
for el in win.descendants():
    try:
        er = el.rectangle()
        if br.top - 20 <= er.top and er.bottom <= br.bottom + 20 and er.right > br.left - 10 and er.left < br.right + 250:
            print(f"  [{el.element_info.control_type}] name='{el.window_text()[:40]}' "
                  f"automation_id='{el.element_info.automation_id}' rect=({er.left},{er.top},{er.right},{er.bottom})")
    except Exception:
        pass

print("\nDone. Leaving Foxit open for manual inspection if needed.")
