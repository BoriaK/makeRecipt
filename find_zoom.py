import pywinauto, time

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'
app = pywinauto.Application(backend='uia').connect(path=FOXIT, timeout=5)
win = app.top_window()

# Click Home tab first
home = [t for t in win.descendants(control_type='TabItem') if t.window_text() == 'Home']
if home:
    home[0].click_input()
    time.sleep(0.5)

r = win.rectangle()
wt = r.top
import pyautogui, win32gui, win32con

hwnd = win.handle
win32gui.BringWindowToTop(hwnd)
time.sleep(0.3)

win = app.top_window()
zoom_el = [el for el in win.descendants(control_type='Edit')
           if el.rectangle().left > 3600 and el.rectangle().top > 980]

if zoom_el:
    ze = zoom_el[-1]
    er = ze.rectangle()
    print('Triple-clicking zoom edit at rect=(' + str(er.left) + ',' + str(er.top) + ',' + str(er.right) + ',' + str(er.bottom) + ')')
    ze.double_click_input()
    time.sleep(0.4)
    pyautogui.typewrite('100', interval=0.05)
    pyautogui.press('enter')
    time.sleep(0.8)
    print('Done — zoom should now be 100%')
else:
    print('Zoom edit not found')














