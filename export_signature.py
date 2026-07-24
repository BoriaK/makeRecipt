"""
export_signature.py
--------------------
One-time helper for the NEW (Foxit-free) receipt flow.

Foxit stores your saved "Fill & Sign" ink signature as a tiny single-page PDF
under your Windows profile. This script renders that page to a transparent
PNG, crops it tightly to the actual ink, and saves it as assets/signature.png
so the new flow can stamp it directly onto generated receipts without ever
touching Foxit again.

Usage:
    python export_signature.py

It will try every known candidate source PDF, render each to
debug/signature_candidate_<n>.png so you can look at them, and write the
best-looking one (largest non-transparent ink area) to assets/signature.png.
"""
import os
import fitz
import numpy as np
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(HERE, "assets")
DEBUG_DIR = os.path.join(HERE, "debug")

CANDIDATE_PDFS = [
    r"C:\Users\Administrator\AppData\Roaming\Foxit Software\Foxit PDF Editor\InstaSign\UserSign\FXPSqnhvwnvhcjzaudj.pdf",
    r"C:\Users\Administrator\AppData\Roaming\Foxit Software\Foxit PDF Editor\FoxitSign\UserSign\FXPSjopsbpfmzmvhuzx.pdf",
]

ZOOM = 4.0  # render at 4x for crisp ink
ALPHA_THRESHOLD = 10  # pixels with alpha above this are considered "ink"
PADDING_PX = 6  # keep a small margin around the cropped ink


def _render_and_crop(pdf_path: str):
    """Render page 0 with a transparent background and crop to the ink's
    bounding box. Returns a PIL RGBA Image, or None if no ink was found."""
    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM), alpha=True)
        img = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
    finally:
        doc.close()

    arr = np.array(img)
    alpha = arr[:, :, 3]
    ink_rows, ink_cols = np.where(alpha > ALPHA_THRESHOLD)
    if ink_rows.size == 0:
        return None, 0

    top = max(0, ink_rows.min() - PADDING_PX)
    bottom = min(arr.shape[0], ink_rows.max() + PADDING_PX + 1)
    left = max(0, ink_cols.min() - PADDING_PX)
    right = min(arr.shape[1], ink_cols.max() + PADDING_PX + 1)

    cropped = img.crop((left, top, right, bottom))
    ink_area = ink_rows.size
    return cropped, ink_area


def main():
    os.makedirs(ASSETS_DIR, exist_ok=True)
    os.makedirs(DEBUG_DIR, exist_ok=True)

    best_img = None
    best_area = -1
    best_source = None

    for i, pdf_path in enumerate(CANDIDATE_PDFS, start=1):
        if not os.path.exists(pdf_path):
            print(f"  (skip, not found) {pdf_path}")
            continue
        img, ink_area = _render_and_crop(pdf_path)
        if img is None:
            print(f"  candidate {i}: no ink found in {pdf_path}")
            continue
        debug_path = os.path.join(DEBUG_DIR, f"signature_candidate_{i}.png")
        img.save(debug_path)
        print(f"  candidate {i}: ink_area={ink_area}px  size={img.size}  -> {debug_path}")
        if ink_area > best_area:
            best_area = ink_area
            best_img = img
            best_source = pdf_path

    if best_img is None:
        raise RuntimeError(
            "No signature ink found in any candidate PDF. Make sure a "
            "signature has been saved in Foxit's Fill & Sign feature."
        )

    out_path = os.path.join(ASSETS_DIR, "signature.png")
    best_img.save(out_path)
    print(f"\n✓  Best signature: {best_source}")
    print(f"✓  Saved: {out_path}  (size={best_img.size})")


if __name__ == "__main__":
    main()
