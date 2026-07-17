"""
place_sig_once.py
-----------------
Can be used two ways:
  1. As a module:   from place_sig_once import apply_signature
                    apply_signature(pdf_path, log_fn=print)
  2. As a script:   python place_sig_once.py [pdf_path]
"""
import sys, time, os, subprocess

# Force this process to be per-monitor DPI-aware BEFORE any window/screenshot
# interaction. Without this, win32gui/win32api/PIL ImageGrab coordinates can
# be virtualized (scaled) inconsistently between separate process launches,
# while pywinauto's UIA backend always reports true physical pixels — a
# mismatch that silently corrupts click math. Safe to call even when this
# module is imported into a process that already set it (e.g. via gui.py);
# a second call just raises, which is ignored.
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE_V2
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import numpy as np
import win32api, win32con, win32gui
import pyautogui
from PIL import ImageGrab

FOXIT = r'C:\Program Files (x86)\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe'

# The Word template is always US Letter (8.5in x 11in = 215.9mm wide).
# Used to calibrate pixels-per-mm from the live screenshot (page white span)
# so the vertical placement offset is a real physical distance, not a
# fixed pixel guess that breaks if the window/zoom size ever changes.
PAGE_WIDTH_MM = 215.9

# Desired gap between the BOTTOM of the placed signature stamp and the
# underline itself.
SIG_ABOVE_MM = 1.5

# Empirically observed: Foxit's Fill & Sign anchors the placement click
# roughly at the stamp's vertical centre rather than its bottom edge, so the
# stamp's ink extends a few px BELOW wherever we click. Verified via pixel
# analysis of a placed stamp (debug/stamp_zoom.png): with a 10px click
# offset, the stamp's ink still reached to within ~4px of the line, i.e. it
# extended ~6px below the click point.
_STAMP_OVERHANG_PX = 6


def _raw_click(x, y, pause=0.3):
    win32api.SetCursorPos((x, y))
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.15)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(pause)


def _find_sigline_visual(win_rect, log, min_run=60, max_run=400, ribbon_h=150):
    """Locate the blank signature underline ('___') by scanning a LIVE
    screenshot of the Foxit window for a long, isolated horizontal run of
    dark pixels — the same visual/variance-detector philosophy used to find
    the signature thumbnail, applied here to the document itself instead of
    matching against a separately fitz-rendered copy of the PDF (which was
    prone to scale/DPI mismatches against the live screen).

    win_rect = (left, top, right, bottom) of the Foxit window.
    """
    wl, wt, wr, wb = win_rect
    scr_left = max(0, wl)
    scr_top  = max(0, wt)
    img = ImageGrab.grab(bbox=(scr_left, scr_top, wr, wb), all_screens=True).convert('L')
    arr = np.array(img)
    H, W = arr.shape

    dark = arr < 150
    candidates = []
    for y in range(ribbon_h, H):
        row = dark[y]
        if not row.any():
            continue
        idx = np.where(row)[0]
        splits = np.where(np.diff(idx) > 1)[0]
        for g in np.split(idx, splits + 1):
            run_len = int(g[-1] - g[0] + 1)
            if min_run < run_len < max_run:
                candidates.append((run_len, int(y), int(g[0]), int(g[-1])))

    if not candidates:
        raise RuntimeError(
            "Signature underline not found on screen (no isolated dark run of "
            f"{min_run}-{max_run}px detected below y={ribbon_h}).")

    # The signature line sits near the end of the document, so prefer the
    # bottom-most matching run; break ties by longest run length.
    candidates.sort(key=lambda c: (c[1], c[0]))
    run_len, y, x0, x1 = candidates[-1]
    log(f"Underline detected: run={run_len}px  row_y={y}  x=[{x0},{x1}]  "
        f"(from {len(candidates)} candidate(s))")

    # Calibrate pixels-per-mm from the page's own white horizontal span in
    # this same row (page background is white; margins/canvas around it are
    # not), so the vertical offset below is a real ~1mm regardless of window
    # size, zoom level, or monitor DPI.
    white_idx = np.where(arr[y] >= 250)[0]
    if len(white_idx) >= 2:
        page_px_width = int(white_idx[-1] - white_idx[0] + 1)
        px_per_mm = page_px_width / PAGE_WIDTH_MM
    else:
        px_per_mm = 8.8  # fallback: matches the calibrated 1940-wide window
    above_px = _STAMP_OVERHANG_PX + round(SIG_ABOVE_MM * px_per_mm)
    log(f"Calibration: px_per_mm={px_per_mm:.2f}  above_px={above_px}")

    cx = (x0 + x1) // 2
    tx = scr_left + cx
    ty = scr_top + y - above_px
    log(f"Target: ({tx},{ty})")
    return tx, ty


def apply_signature(pdf_path: str, log_fn=print) -> None:
    """
    Public entry point. Wraps _apply_signature_impl so that ANY exception
    (including ones from pywinauto/win32 that aren't explicitly raised as
    RuntimeError) gets its full traceback written to debug/run_log.txt
    before propagating — this lets us diagnose failures after the fact
    without needing a live console.
    """
    try:
        return _apply_signature_impl(pdf_path, log_fn)
    except Exception:
        import traceback
        _tb = traceback.format_exc()
        try:
            _p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'debug', 'run_log.txt')
            os.makedirs(os.path.dirname(_p), exist_ok=True)
            with open(_p, 'a', encoding='utf-8') as _f:
                _f.write(_tb + "\n")
        except Exception:
            pass
        raise


def _apply_signature_impl(pdf_path: str, log_fn=print) -> None:
    """
    Open *pdf_path* in Foxit, place the saved Fill & Sign signature on the
    שם וחתימה line, apply all signatures, save and close Foxit.

    Raises RuntimeError on any failure.
    log_fn is called with progress strings (safe to pass a GUI logger).
    """
    import pywinauto
    from datetime import datetime as _dt

    _debug_dir  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'debug')
    os.makedirs(_debug_dir, exist_ok=True)
    _log_path   = os.path.join(_debug_dir, 'run_log.txt')
    _user_log   = log_fn

    def log_fn(msg):  # noqa: F811 - intentional shadow, wraps user log_fn with file logging
        try:
            with open(_log_path, 'a', encoding='utf-8') as _f:
                _f.write(f"[{_dt.now():%H:%M:%S}] {msg}\n")
        except Exception:
            pass
        _user_log(msg)

    log_fn("="*20 + " apply_signature start " + "="*20)

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

    # ── 4. Click Fill & Sign ──────────────────────────────────────
    # NOTE: this must happen BEFORE sigline template matching. The Fill & Sign
    # ribbon adds an extra thumbnail-strip row that pushes the whole document
    # viewport down, so any target coordinates computed beforehand go stale.
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

    # ── 5. Find Signature List button ─────────────────────────────
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
    log_fn(f"Signature List rect: ({br.left},{br.top},{br.right},{br.bottom})  "
           f"size={br.right - br.left}x{br.bottom - br.top}")

    # ── 7. Click the signature thumbnail ──────────────────────────
    # Confirmed via live diagnostic testing: the ink thumbnail sits in the
    # LEFT portion of the 'Signature List' container (roughly the first
    # fifth of its width), well clear of the wider 'Create Signature' (+)
    # button further right. Clicking the container's centre or relying on
    # a narrower sub-region anchored off the (unreliable) 'Create Signature'
    # button edge both missed the thumbnail in earlier attempts.
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.3)

    _click_x = br.left + int((br.right - br.left) * 0.20)
    _click_y = (br.top + br.bottom) // 2
    log_fn(f"Clicking thumbnail at ({_click_x},{_click_y})")
    # Use pyautogui (real interpolated mouse movement) instead of an instant
    # SetCursorPos jump. Foxit's Fill & Sign attaches a floating signature
    # preview to the cursor after this click; some Foxit builds only update
    # that internal "attached" state correctly when they actually receive
    # WM_MOUSEMOVE events en route, not just a teleport + click.
    pyautogui.moveTo(_click_x, _click_y, duration=0.25)
    time.sleep(0.2)
    pyautogui.click()
    time.sleep(2.0)
    log_fn("Signature thumbnail clicked")

    # ── Re-detect signature line NOW, with the Fill & Sign ribbon (and its
    # extra thumbnail-strip row) already in its final, settled state. Doing
    # this earlier (before the ribbon changed size) produced stale coordinates
    # that missed the dotted line entirely.
    win = app.top_window()
    wl2, wt2, wr2, wb2 = win32gui.GetWindowRect(hwnd)
    tx, ty = _find_sigline_visual((wl2, wt2, wr2, wb2), log_fn)

    # Save debug screenshot to confirm what's on screen before placement
    _dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'debug', 'before_place_new.png')
    os.makedirs(os.path.dirname(_dbg), exist_ok=True)
    ImageGrab.grab(bbox=(max(0,wl2), max(0,wt2), wr2, wb2), all_screens=True).save(_dbg)
    log_fn(f"Debug screenshot saved: {_dbg}")

    # Move smoothly onto the document (generates real WM_MOUSEMOVE events so
    # Foxit's floating signature preview tracks the cursor and is actually
    # "hovering" at this spot when the click commits), then click to stamp.
    pyautogui.moveTo(tx, ty, duration=0.3)
    time.sleep(0.3)
    pyautogui.click()
    time.sleep(1.5)
    log_fn("Signature placed")

    # Verify visually: did anything actually change at the target spot?
    _after_place = os.path.join(_debug_dir, 'after_place_click.png')
    ImageGrab.grab(bbox=(max(0,wl2), max(0,wt2), wr2, wb2), all_screens=True).save(_after_place)
    log_fn(f"Debug screenshot saved: {_after_place}")

    # Check whether the click-to-place actually attached anything by looking
    # at 'Apply All Signatures' enabled state. If it's still disabled, retry
    # with a real drag-and-drop gesture (mouse down on thumbnail, move while
    # held, mouse up on the target) — some Foxit builds only accept placement
    # via drag rather than click-then-click.
    win = app.top_window()
    _apply_check = [b for b in win.descendants(control_type='Button')
                    if b.window_text() == 'Apply All Signatures']
    _still_disabled = True
    if _apply_check:
        try:
            _still_disabled = not _apply_check[0].is_enabled()
        except Exception:
            _still_disabled = True
    if _still_disabled:
        log_fn("Click-to-place did not attach — retrying via drag-and-drop")
        pyautogui.moveTo(_click_x, _click_y, duration=0.2)
        time.sleep(0.2)
        pyautogui.mouseDown()
        time.sleep(0.2)
        pyautogui.moveTo(tx, ty, duration=0.5)
        time.sleep(0.2)
        pyautogui.mouseUp()
        time.sleep(1.5)
        log_fn("Drag-and-drop placement attempted")
        _after_drag = os.path.join(_debug_dir, 'after_drag_place.png')
        ImageGrab.grab(bbox=(max(0,wl2), max(0,wt2), wr2, wb2), all_screens=True).save(_after_drag)
        log_fn(f"Debug screenshot saved: {_after_drag}")

    # ── 8. Move mouse away ────────────────────────────────────────
    win32api.SetCursorPos((wl2 + 150, wt2 + 300))
    time.sleep(0.5)

    # ── 9. Apply All Signatures ───────────────────────────────────
    win = app.top_window()
    apply_btn = [b for b in win.descendants(control_type='Button')
                 if b.window_text() == 'Apply All Signatures']
    if not apply_btn:
        raise RuntimeError("'Apply All Signatures' button not found via UIA")
    try:
        _enabled = apply_btn[0].is_enabled()
    except Exception:
        _enabled = None
    log_fn(f"'Apply All Signatures' enabled={_enabled}")
    if _enabled is False:
        log_fn("  ⚠ WARNING: button appears DISABLED — the signature was likely "
               "NOT placed (click at (%d,%d) missed the document)" % (tx, ty))
    er = apply_btn[0].rectangle()
    _raw_click((er.left + er.right) // 2, (er.top + er.bottom) // 2)
    time.sleep(2.0)
    log_fn("Apply All Signatures clicked")

    _after_apply = os.path.join(_debug_dir, 'after_apply_click.png')
    ImageGrab.grab(bbox=(max(0,wl2), max(0,wt2), wr2, wb2), all_screens=True).save(_after_apply)
    log_fn(f"Debug screenshot saved: {_after_apply}")

    # ── 10. Save ──────────────────────────────────────────────────
    win32gui.BringWindowToTop(hwnd)
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 's')
    time.sleep(2.0)
    # Only dismiss a dialog if one actually appeared — never send Enter to the document.
    # Log the dialog's text first so an unexpected dialog (e.g. Save As, permission
    # error) is visible in run_log.txt instead of being silently dismissed.
    try:
        dlg = app.window(class_name='#32770', top_level_only=True)
        if dlg.exists(timeout=1):
            try:
                _dlg_texts = [w.window_text() for w in dlg.descendants()
                              if w.window_text().strip()]
                log_fn(f"Dialog appeared after Ctrl+S: {_dlg_texts}")
            except Exception:
                pass
            dlg.type_keys('{ENTER}')
            time.sleep(1.0)
        else:
            log_fn("No dialog appeared after Ctrl+S (saved silently)")
    except Exception as _e:
        log_fn(f"Save-dialog check error: {_e}")
    log_fn("Saved")

    _after_save = os.path.join(_debug_dir, 'after_save.png')
    ImageGrab.grab(bbox=(max(0,wl2), max(0,wt2), wr2, wb2), all_screens=True).save(_after_save)
    log_fn(f"Debug screenshot saved: {_after_save}")

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
