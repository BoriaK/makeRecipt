from PIL import ImageGrab
import numpy as np

# Window maximized at (1910,-10) -> (3850,1030), ribbon ~130px tall
img = ImageGrab.grab(bbox=(1920, 140, 3840, 1030), all_screens=True)
img.save('fullscreen_doc.png')
scr = np.array(img.convert('RGB'))
print(f"Screenshot size: {scr.shape}")

# Sample horizontal lines to find actual white page edges
for y_offset in [50, 200, 400, 600, 750]:
    row = scr[y_offset]
    w = np.where((row[:,0]>250) & (row[:,1]>250) & (row[:,2]>250))[0]
    if len(w):
        print(f"  screen_y={y_offset+140}: white x={w[0]+1920}..{w[-1]+1920}  width={w[-1]-w[0]}px")
    else:
        print(f"  screen_y={y_offset+140}: no white pixels")

# Find the grey value of the margins
print("\nSample pixel colors (left margin, middle, right margin) at y_offset=400:")
for x in [10, 50, 100, 500, 900, 1700, 1850, 1900]:
    px = scr[400, x]
    print(f"  x={x+1920}: RGB({px[0]},{px[1]},{px[2]})")

