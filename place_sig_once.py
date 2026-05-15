"""Open existing PDF in Foxit, apply saved signature (no save)."""
import sys, time, os, subprocess
import fitz
import numpy as np
import cv2
import pywinauto
import win32api, win32con, win32gui
import pyautogui
from PIL import ImageGrab

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'
# Accept PDF path from command line, fall back to hardcoded default
PDF   = sys.argv[1] if len(sys.argv) > 1 else \
        r'C:\Users\Administrator\Documents\קבלות_על_אימונים\recipt_16_April_2026.pdf'

# The שם וחתימה dotted line sits at PDF y ≈ 301-314 pt (PyMuPDF top-down).
# At 2x render that is rendered_y = 602-628.  We take a slightly wider strip.
SIG_STRIP_Y1 = 580    # rendered y top    (2x)
SIG_STRIP_Y2 = 650    # rendered y bottom (2x)
SIG_X_PT     = 465.0  # calibrated horizontal target (PDF points from page left)
# How many pixels ABOVE the centre of the matched strip to click
ABOVE_PX     = 10


def stop(msg):
    print(f"\n*** STOPPING: {msg} ***")
    sys.exit(1)


def raw_click(x, y, pause=0.3):
    win32api.SetCursorPos((x, y))
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.15)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(pause)


def find_sigline_target():
    """
    Render the signature-line strip from the PDF at 2x and template-match it
    against the current Foxit screenshot.

    Returns (tx, ty) — the screen coordinates to click for signature placement:
      tx = left edge of match + SIG_X_PT scaled to screen pixels
      ty = vertical centre of match − ABOVE_PX
    Stops if confidence < 0.50.
    """
    # ── Render the strip ─────────────────────────────────────────
    doc = fitz.open(PDF)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
    rendered = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n)
    if pix.n == 4:
        rendered = rendered[:, :, :3]
    doc.close()

    strip = rendered[SIG_STRIP_Y1:SIG_STRIP_Y2, :]   # full page width

    # ── Screenshot of right monitor ──────────────────────────────
    scr = np.array(ImageGrab.grab(bbox=(1920, 0, 3840, 1080),
                                   all_screens=True).convert('RGB'))

    # ── Template match across scales ─────────────────────────────
    best = None
    for scale in np.arange(0.80, 1.40, 0.02):
        w = int(strip.shape[1] * scale)
        h = int(strip.shape[0] * scale)
        if w >= scr.shape[1] or h >= scr.shape[0]:
            continue
        tmpl = cv2.resize(strip, (w, h))
        res  = cv2.matchTemplate(
            cv2.cvtColor(scr[150:], cv2.COLOR_RGB2GRAY),
            cv2.cvtColor(tmpl,      cv2.COLOR_RGB2GRAY),
            cv2.TM_CCOEFF_NORMED)
        _, val, _, loc = cv2.minMaxLoc(res)
        if best is None or val > best[0]:
            # loc is relative to scr[150:]; convert to absolute screen coords
            best = (val, scale, loc[0] + 1920, loc[1] + 150, w, h)

    val, scale, mx, my, mw, mh = best
    print(f"Sigline match: confidence={val:.3f}  scale={scale:.2f}")
    print(f"Match top-left on screen: ({mx},{my})  template size: {mw}x{mh}")

    if val < 0.50:
        stop(f"Signature line not found on screen (confidence={val:.3f} < 0.50). "
             "Is the dotted line visible?")

    # mx is the screen x of the PAGE LEFT EDGE at this y level (full-width template)
    tx = mx + int(SIG_X_PT * 2 * scale)        # SIG_X_PT in rendered px at 2x
    ty = my + mh // 2 - ABOVE_PX               # centre of strip, shifted up

    print(f"Target: ({tx},{ty})")
    return tx, ty


# ── 0. Open the PDF ──────────────────────────────────────────────
print(f"Opening {os.path.basename(PDF)} ...")
subprocess.Popen([FOXIT, PDF])

pdf_name = os.path.basename(PDF)
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
    stop("Foxit did not open the PDF within 30 s")
print("PDF loaded")

# ── 1. Maximize + verify full screen ────────────────────────────
win  = app.top_window()
hwnd = win.handle
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
win32gui.BringWindowToTop(hwnd)
time.sleep(3.0)   # give Foxit time to fully render the page

win = app.top_window()
r   = win.rectangle()
wl, wt, wr, wb = r.left, r.top, r.right, r.bottom
print(f"Window: ({wl},{wt}) -> ({wr},{wb})  size={wr-wl}x{wb-wt}")
if (wr - wl) < 1600 or (wb - wt) < 800:
    stop("Window is not full screen — maximize failed.")
print("Full screen confirmed")

# ── 2. Click Home tab ────────────────────────────────────────────
home = [t for t in win.descendants(control_type='TabItem') if t.window_text() == 'Home']
if not home:
    stop("'Home' tab not found via UIA")
home[0].click_input()
time.sleep(0.8)
print("Home tab clicked")

# ── 2b. Set zoom to 100% via status bar edit box ─────────────────
win = app.top_window()
zoom_el = [el for el in win.descendants(control_type='Edit')
           if el.rectangle().left > 3600 and el.rectangle().top > 980]
if zoom_el:
    zoom_el[-1].double_click_input()
    time.sleep(0.4)
    pyautogui.typewrite('100', interval=0.05)
    pyautogui.press('enter')
    time.sleep(0.8)
    print("Zoom set to 100%")
else:
    print("  (zoom edit not found, continuing)")

# ── 3. Press Page Up several times to ensure we are at the top of the document ─
win32gui.BringWindowToTop(hwnd)
time.sleep(0.3)
for _ in range(8):
    pyautogui.press('pageup')
    time.sleep(0.15)
time.sleep(0.5)
print("Scrolled to document top")

# ── 4. Locate the dotted signature line via template matching ────
tx, ty = find_sigline_target()

# ── 5. Click Fill & Sign (Button in fresh window, Tab if already used) ──
win = app.top_window()
fs = [t for t in win.descendants(control_type='TabItem') if t.window_text() == 'Fill & Sign']
if fs:
    fs[0].click_input()
    print("Fill & Sign tab clicked")
else:
    btn = [b for b in win.descendants(control_type='Button')
           if 'Fill' in b.window_text() and 'Sign' in b.window_text()]
    if not btn:
        stop("'Fill & Sign' button/tab not found via UIA")
    btn[0].click_input()
    print(f"Fill & Sign button clicked: [{btn[0].window_text()}]")
time.sleep(2.0)

# ── 6. Find Signature List button (retry up to 8 s) ─────────────
sig_btn = None
for attempt in range(8):
    win = app.top_window()
    sig_btn = [b for b in win.descendants(control_type='Button')
               if b.window_text() == 'Signature List']
    if sig_btn:
        break
    print(f"  Signature List not ready... ({attempt+1}/8)")
    time.sleep(1.0)
if not sig_btn:
    stop("'Signature List' button not found via UIA after 8 s")

br    = sig_btn[0].rectangle()
sig_x = br.left + 80
sig_y = (br.top + br.bottom) // 2
print(f"Signature List rect: ({br.left},{br.top},{br.right},{br.bottom})  clicking at ({sig_x},{sig_y})")

# ── 7. Pick up signature + place it ─────────────────────────────
raw_click(sig_x, sig_y, pause=1.5)   # activate signature cursor from thumbnail
raw_click(tx,    ty,    pause=1.5)   # stamp on the document

# ── 8. Move mouse away (no click, no Escape) ─────────────────────
win32api.SetCursorPos((wl + 150, wt + 300))
time.sleep(0.5)

# ── 9. Apply All Signatures ──────────────────────────────────────
win = app.top_window()
apply_btn = [b for b in win.descendants(control_type='Button')
             if b.window_text() == 'Apply All Signatures']
if not apply_btn:
    stop("'Apply All Signatures' button not found via UIA")
apply_btn[0].click_input()
time.sleep(2.0)
print("Apply All Signatures clicked")

# ── 10. Save (Ctrl+S) + close (Ctrl+W) ───────────────────────────
win32gui.BringWindowToTop(hwnd)
time.sleep(0.3)
pyautogui.hotkey('ctrl', 's')
time.sleep(2.0)
pyautogui.press('enter')   # dismiss overwrite dialog if shown
time.sleep(2.0)
print("Saved")

pyautogui.hotkey('alt', 'f4')
time.sleep(1.5)
pyautogui.press('enter')   # dismiss unsaved-changes dialog if shown
time.sleep(1.0)
print("Foxit closed — all done!")
