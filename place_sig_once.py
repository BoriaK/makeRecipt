"""
place_sig_once.py
-----------------
Can be used two ways:
  1. As a module:   from place_sig_once import apply_signature
                    apply_signature(pdf_path, log_fn=print)
  2. As a script:   python place_sig_once.py [pdf_path]
"""
import sys, time, os, subprocess
import fitz
import numpy as np
import cv2
import win32api, win32con, win32gui
import pyautogui
from PIL import ImageGrab

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'

SIG_STRIP_Y1 = 580
SIG_STRIP_Y2 = 650
SIG_X_PT     = 465.0
ABOVE_PX     = 10


def _raw_click(x, y, pause=0.3):
    win32api.SetCursorPos((x, y))
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.15)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(pause)


def _find_sigline_target(pdf_path, log, win_rect):
    """win_rect = (left, top, right, bottom) of the Foxit window."""
    doc = fitz.open(pdf_path)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
    rendered = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n)
    if pix.n == 4:
        rendered = rendered[:, :, :3]
    doc.close()

    strip = rendered[SIG_STRIP_Y1:SIG_STRIP_Y2, :]

    # Screenshot of whatever monitor the Foxit window is on
    wl, wt, wr, wb = win_rect
    # Clamp to screen bounds (window may have negative top with chrome)
    scr_left  = max(0, wl)
    scr_top   = max(0, wt)
    scr_right = wr
    scr_bot   = wb
    scr = np.array(ImageGrab.grab(bbox=(scr_left, scr_top, scr_right, scr_bot),
                                   all_screens=True).convert('RGB'))

    ribbon_h = 150   # skip ribbon rows in the match
    best = None
    for scale in np.arange(0.80, 1.40, 0.02):
        w = int(strip.shape[1] * scale)
        h = int(strip.shape[0] * scale)
        if w >= scr.shape[1] or h >= scr.shape[0] - ribbon_h:
            continue
        tmpl = cv2.resize(strip, (w, h))
        res  = cv2.matchTemplate(
            cv2.cvtColor(scr[ribbon_h:], cv2.COLOR_RGB2GRAY),
            cv2.cvtColor(tmpl,           cv2.COLOR_RGB2GRAY),
            cv2.TM_CCOEFF_NORMED)
        _, val, _, loc = cv2.minMaxLoc(res)
        if best is None or val > best[0]:
            # Convert back to absolute screen coords
            best = (val, scale,
                    loc[0] + scr_left,
                    loc[1] + scr_top + ribbon_h,
                    w, h)

    val, scale, mx, my, mw, mh = best
    log(f"Sigline match: confidence={val:.3f}  scale={scale:.2f}")
    if val < 0.50:
        raise RuntimeError(
            f"Signature line not found on screen (confidence={val:.3f} < 0.50).")
    tx = mx + int(SIG_X_PT * 2 * scale)
    ty = my + mh // 2 - ABOVE_PX
    log(f"Target: ({tx},{ty})")
    return tx, ty


def apply_signature(pdf_path: str, log_fn=print) -> None:
    """
    Open *pdf_path* in Foxit, place the saved Fill & Sign signature on the
    שם וחתימה line, apply all signatures, save and close Foxit.

    Raises RuntimeError on any failure.
    log_fn is called with progress strings (safe to pass a GUI logger).
    """
    import pywinauto

    # ── 0. Open the PDF ───────────────────────────────────────────
    log_fn(f"Opening {os.path.basename(pdf_path)} ...")
    subprocess.Popen([FOXIT, pdf_path])

    pdf_name = os.path.basename(pdf_path)
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
        raise RuntimeError("Foxit did not open the PDF within 30 s")
    log_fn("PDF loaded")

    # ── 1. Force window to fill the right monitor ────────────────
    win  = app.top_window()
    hwnd = win.handle

    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.5)

    # Move into right monitor then maximize — let Windows handle the exact size.
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.3)
    win32gui.MoveWindow(hwnd, 2500, 200, 800, 600, True)
    time.sleep(0.3)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.BringWindowToTop(hwnd)
    time.sleep(2.5)   # wait for maximize + page render

    wl, wt, wr, wb = win32gui.GetWindowRect(hwnd)
    log_fn(f"Window: ({wl},{wt}) -> ({wr},{wb})  size={wr-wl}x{wb-wt}")
    if (wr - wl) < 1400 or (wb - wt) < 700:
        raise RuntimeError(
            f"Window too small after positioning ({wr-wl}x{wb-wt}). "
            "Check monitor layout.")
    log_fn("Full screen confirmed")

    # ── 2. Click Home tab ─────────────────────────────────────────
    win = app.top_window()
    home = [t for t in win.descendants(control_type='TabItem') if t.window_text() == 'Home']
    if not home:
        raise RuntimeError("'Home' tab not found via UIA")
    er = home[0].rectangle()
    _raw_click((er.left + er.right) // 2, (er.top + er.bottom) // 2)
    time.sleep(0.8)
    log_fn("Home tab clicked")

    # ── 2b. Set zoom to 100% ──────────────────────────────────────
    win = app.top_window()
    # Zoom edit box is in the status bar — far right, near window bottom
    zoom_el = [el for el in win.descendants(control_type='Edit')
               if el.rectangle().left > wr - 250 and el.rectangle().top > wb - 60]
    if zoom_el:
        zoom_el[-1].double_click_input()
        time.sleep(0.4)
        pyautogui.typewrite('100', interval=0.05)
        pyautogui.press('enter')
        time.sleep(0.8)
        log_fn("Zoom set to 100%")
    else:
        log_fn("  (zoom edit not found, continuing)")

    # ── 3. Scroll to top of document ─────────────────────────────
    # Click somewhere in the document body to return focus from the zoom box,
    # then press Page Up several times to guarantee we are at page 1 top.
    doc_cx = (wl + wr) // 2
    doc_cy = wt + (wb - wt) // 2
    _raw_click(doc_cx, doc_cy, pause=0.5)
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.3)
    for _ in range(10):
        pyautogui.press('pageup')
        time.sleep(0.12)
    time.sleep(0.5)
    log_fn("Scrolled to document top")

    # ── 4. Template-match signature line ─────────────────────────
    tx, ty = _find_sigline_target(pdf_path, log_fn, (wl, wt, wr, wb))

    # ── 5. Click Fill & Sign ──────────────────────────────────────
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.3)
    win = app.top_window()
    fs = [t for t in win.descendants(control_type='TabItem') if t.window_text() == 'Fill & Sign']
    if fs:
        er = fs[0].rectangle()
        _raw_click((er.left + er.right) // 2, (er.top + er.bottom) // 2)
        log_fn("Fill & Sign tab clicked")
    else:
        btn = [b for b in win.descendants(control_type='Button')
               if 'Fill' in b.window_text() and 'Sign' in b.window_text()]
        if not btn:
            raise RuntimeError("'Fill & Sign' button/tab not found via UIA")
        er = btn[0].rectangle()
        _raw_click((er.left + er.right) // 2, (er.top + er.bottom) // 2)
        log_fn(f"Fill & Sign button clicked: [{btn[0].window_text()}]")
    time.sleep(2.0)

    # ── 6. Find Signature List button ─────────────────────────────
    sig_btn = None
    for attempt in range(8):
        win = app.top_window()
        sig_btn = [b for b in win.descendants(control_type='Button')
                   if b.window_text() == 'Signature List']
        if sig_btn:
            break
        log_fn(f"  Signature List not ready... ({attempt+1}/8)")
        time.sleep(1.0)
    if not sig_btn:
        raise RuntimeError("'Signature List' button not found via UIA after 8 s")

    br    = sig_btn[0].rectangle()
    sig_x = br.left + 80
    sig_y = (br.top + br.bottom) // 2
    log_fn(f"Signature List rect: ({br.left},{br.top},{br.right},{br.bottom})  clicking thumbnail at ({sig_x},{sig_y})")

    # ── 7. Pick up signature + place it ──────────────────────────
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.3)
    _raw_click(sig_x, sig_y, pause=2.0)   # click thumbnail — sig attaches to cursor
    log_fn(f"Signature thumbnail clicked — now placing at ({tx},{ty})")

    # Save debug screenshot to confirm what's on screen before placement
    _dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'debug', 'before_place_new.png')
    os.makedirs(os.path.dirname(_dbg), exist_ok=True)
    ImageGrab.grab(bbox=(max(0,wl), max(0,wt), wr, wb), all_screens=True).save(_dbg)
    log_fn(f"Debug screenshot saved: {_dbg}")

    _raw_click(tx, ty, pause=1.5)          # stamp on document
    log_fn("Signature placed")

    # ── 8. Move mouse away ────────────────────────────────────────
    win32api.SetCursorPos((wl + 150, wt + 300))
    time.sleep(0.5)

    # ── 9. Apply All Signatures ───────────────────────────────────
    win = app.top_window()
    apply_btn = [b for b in win.descendants(control_type='Button')
                 if b.window_text() == 'Apply All Signatures']
    if not apply_btn:
        raise RuntimeError("'Apply All Signatures' button not found via UIA")
    er = apply_btn[0].rectangle()
    _raw_click((er.left + er.right) // 2, (er.top + er.bottom) // 2)
    time.sleep(2.0)
    log_fn("Apply All Signatures clicked")

    # ── 10. Save ──────────────────────────────────────────────────
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 's')
    time.sleep(2.0)
    # Only dismiss a dialog if one actually appeared — never send Enter to the document
    try:
        dlg = app.window(class_name='#32770', top_level_only=True)
        if dlg.exists(timeout=1):
            dlg.type_keys('{ENTER}')
            time.sleep(1.0)
    except Exception:
        pass
    log_fn("Saved")

    # ── 11. Close Foxit ───────────────────────────────────────────
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1.5)
    # Only dismiss unsaved-changes dialog if it appeared
    try:
        dlg = app.window(class_name='#32770', top_level_only=True)
        if dlg.exists(timeout=1):
            dlg.type_keys('{ENTER}')
            time.sleep(0.5)
    except Exception:
        pass
    log_fn("Foxit closed — all done!")


# ── Standalone script entry point ────────────────────────────────────────────
if __name__ == '__main__':
    pdf = sys.argv[1] if len(sys.argv) > 1 else \
          r'C:\Users\Administrator\Documents\קבלות_על_אימונים\recipt_28_May_2026.pdf'
    apply_signature(pdf, log_fn=print)
