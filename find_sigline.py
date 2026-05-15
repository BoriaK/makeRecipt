import fitz
doc = fitz.open(r'C:\Users\Administrator\Documents\קבלות_על_אימונים\recipt_16_April_2026.pdf')
page = doc[0]
paths = page.get_drawings()
print(f'Total drawings: {len(paths)}')
for p in paths:
    r = p['rect']
    if 280 < r.y0 < 340:
        w = r.x1 - r.x0
        print(f"  rect=({r.x0:.1f},{r.y0:.1f},{r.x1:.1f},{r.y1:.1f})  width={w:.1f}pt  center_x={((r.x0+r.x1)/2):.1f}")

