"""Make a PNG contact sheet of selected ticks from a harness log.
usage: python3 tools/harness/strip.py <name> <tick> [<tick> ...]  -> out/<name>_strip.png"""
import os, sys
from PIL import Image, ImageDraw
HERE = os.path.dirname(os.path.abspath(__file__))
W, H = 192, 100
name, ticks = sys.argv[1], [int(t) for t in sys.argv[2:]]
lines = open(os.path.join(HERE, "out", name + ".log")).read().split("\n")
frames = {}
i = 0
while i < len(lines):
    if lines[i].startswith("F "):
        t = int(lines[i].split()[1])
        if t in ticks:
            img = Image.new("RGB", (W, H)); px = img.load()
            for y in range(H):
                row = lines[i + 1 + y]
                for x in range(W):
                    v = int(row[x * 6:x * 6 + 6], 16)
                    px[x, y] = (v & 255, (v >> 8) & 255, (v >> 16) & 255)
            frames[t] = img
        i += H + 1
    else:
        i += 1
S, cols = 3, min(4, len(ticks))
rows = (len(ticks) + cols - 1) // cols
sheet = Image.new("RGB", (cols * W * S, rows * (H * S + 14)), (0, 0, 0))
d = ImageDraw.Draw(sheet)
for k, t in enumerate(ticks):
    if t in frames:
        x, y = (k % cols) * W * S, (k // cols) * (H * S + 14)
        sheet.paste(frames[t].resize((W * S, H * S), Image.NEAREST), (x, y + 14))
        d.text((x + 4, y + 1), f"tick {t}", fill=(255, 255, 255))
out = os.path.join(HERE, "out", name + "_strip.png")
sheet.save(out)
print(out)
