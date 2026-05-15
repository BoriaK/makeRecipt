import pywinauto, win32gui, time

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'
app = pywinauto.Application(backend='uia').connect(path=FOXIT, timeout=10)
win = app.top_window()
win32gui.BringWindowToTop(win.handle)
time.sleep(0.5)

home = [t for t in win.descendants(control_type='TabItem') if t.window_text()=='Home']
if home: home[0].click_input(); time.sleep(0.5)
fs = [t for t in win.descendants(control_type='TabItem') if t.window_text()=='Fill & Sign']
fs[0].click_input()
time.sleep(1.5)

r = win.rectangle()
wt = r.top
print('UI elements in ribbon (y < wt+140):')
for el in win.descendants():
    try:
        er = el.rectangle()
        if er.top < wt + 140 and er.bottom > wt + 30:
            txt = el.window_text()[:50]
            ct = el.element_info.control_type
            print(f'  [{ct}] "{txt}"  rect=({er.left},{er.top},{er.right},{er.bottom})')
    except:
        pass

