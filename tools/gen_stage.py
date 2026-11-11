"""Generate the stage background (data/stage_sanmun.txt + preview PNG).

Stage: dawn courtyard on a mountain top (sanmun = mountain gate).
Colors are quantized into bands on purpose: fewer distinct runs per row
means fewer Range paint calls in Excel.

usage: python tools/gen_stage.py
"""
import math
import os
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
W, H = 120, 64
GROUND_Y = 56          # must match Renderer.GROUND_Y

PAL = {
    # sky bands (top -> horizon)
    "0": "1E1B3A", "1": "2A2550", "2": "3D2F63", "3": "5A3A6E",
    "4": "804670", "5": "A8566A", "6": "CF7262", "7": "E8955A",
    "u": "F7D9A0", "U": "FBE9C4",              # sun
    "f": "7A5F86", "F": "8C6E93",              # far mountains / their lit edge
    "m": "4F4672", "M": "5E5480",              # mid mountains
    "n": "2E3052", "N": "393B5E",              # near ridge
    "z": "A98AA8",                              # mist
    "t": "1C1C33",                              # pagoda / silhouettes
    "y": "C9503A",                              # pagoda lit window
    "g": "1F3A34", "G": "2C5045", "T": "3B2A26",  # pine foliage / light / trunk
    "r": "8A7A6A", "R": "6A5C50", "q": "A89780",  # stone railing
    "s": "7D6E5E", "S": "6B5D4F", "e": "9C8B74", "d": "4E433A",  # floor
}

def ridge(x, base, amps):
    v = base
    for a, f, ph in amps:
        v += a * math.sin(x * f + ph)
    return v

def build():
    g = [["0"] * W for _ in range(H)]
    # sky
    bands = "01234567"
    for y in range(GROUND_Y):
        t = min(1.0, y / 44.0)
        g[y] = [bands[min(7, int(t * 7.999))]] * W
    # sun
    sx, sy, sr = 86, 23, 7.5
    for y in range(H):
        for x in range(W):
            d = math.hypot(x + 0.5 - sx, y + 0.5 - sy)
            if d <= sr:
                g[y][x] = "U" if d <= sr - 2.2 else "u"
    # mountain layers
    layers = [
        ("f", "F", lambda x: ridge(x, 33, [(5, 0.07, 0.3), (2.5, 0.19, 1.7), (1.0, 0.5, 0.2)])),
        ("m", "M", lambda x: ridge(x, 40, [(5, 0.05, 2.0), (2.0, 0.23, 0.4), (0.8, 0.61, 1.1)])),
        ("n", "N", lambda x: ridge(x, 46, [(3, 0.045, 4.0), (1.5, 0.17, 2.2), (0.7, 0.7, 0.5)])),
    ]
    tops = {}
    for col, lit, fn in layers:
        top = [int(fn(x)) for x in range(W)]
        tops[col] = top
        for x in range(W):
            for y in range(max(0, top[x]), GROUND_Y):
                g[y][x] = col
            # lit edge on the sun side of each ridge
            if top[x] >= 0 and x > 0 and top[x] < top[x - 1]:
                g[top[x]][x] = lit
            if 0 <= top[x] < GROUND_Y:
                g[top[x]][x] = lit if (x * 7) % 3 == 0 else g[top[x]][x]
        if col == "f":
            # mist strips over the far range
            for y in (38, 39, 42):
                for x in range(W):
                    if (x // 6 + y) % 3 != 0 and g[y][x] in ("f", "F", "m"):
                        g[y][x] = "z"
    # pagoda on the mid ridge (left side)
    px = 22
    base = tops["m"][px]
    tiers = [(9, 0), (7, -4), (5, -8), (3, -11)]
    for w, dy in tiers:
        yb = base + dy
        for x in range(px - w // 2, px + w // 2 + 1):
            for y in range(yb - 3, yb + 1):
                if 0 <= y < H:
                    g[y][x] = "t"
        # curved eaves
        for x in (px - w // 2 - 1, px + w // 2 + 1):
            if 0 <= yb - 3 < H:
                g[yb - 3][x] = "t"
        if 0 <= yb - 1 < H:
            g[yb - 1][px] = "y"
    for y in range(base - 15, base - 11):
        g[y][px] = "t"
    # pine tree on the right edge (foreground silhouette)
    tx = 108
    for y in range(20, GROUND_Y):
        for x in (tx, tx + 1):
            g[y][x] = "T"
    for cy, half in ((22, 9), (28, 11), (34, 8), (40, 10)):
        for dy in range(-2, 3):
            span = half - abs(dy) * 2
            for x in range(tx - span, tx + span // 2 + 2):
                if 0 <= x < W and 0 <= cy + dy < H:
                    g[cy + dy][x] = "G" if dy == -2 and x % 2 == 0 else "g"
    # a branch reaching left
    for x in range(96, 108):
        y = 30 + (108 - x) // 5
        g[y][x] = "T"
    # stone railing behind the fighters
    rail_top = 50
    for x in range(W):
        g[rail_top][x] = "q"
        g[rail_top + 1][x] = "r"
        if x % 12 in (0, 1):
            for y in range(rail_top - 1, GROUND_Y):
                g[y][x] = "r" if x % 12 == 0 else "R"
            g[rail_top - 1][x] = "q"
        elif x % 4 == 2:
            for y in range(rail_top + 2, GROUND_Y):
                g[y][x] = "R"
    # stone floor with tile seams
    for y in range(GROUND_Y, H):
        for x in range(W):
            g[y][x] = "s"
        if y == GROUND_Y:
            g[y] = ["e"] * W
    for y in (59, 62):
        for x in range(W):
            g[y][x] = "S"
    for y in range(GROUND_Y + 1, H):
        row_band = 0 if y < 59 else (1 if y < 62 else 2)
        offs = (0, 8, 4)[row_band]
        for x in range(W):
            if (x + offs) % 16 == 0 and g[y][x] == "s":
                g[y][x] = "d"
    return g

def main():
    g = build()
    out = ["# stage background - generated by tools/gen_stage.py, do not edit by hand",
           "# PAL <char> <RRGGBB> ; STAGE <name> <w> <h> ; rows"]
    for k, v in PAL.items():
        out.append(f"PAL {k} {v}")
    out.append(f"STAGE sanmun {W} {H}")
    runs = 0
    for row in g:
        out.append("".join(row))
        runs += 1 + sum(1 for i in range(1, W) if row[i] != row[i - 1])
    with open(os.path.join(ROOT, "data", "stage_sanmun.txt"), "w", newline="\r\n") as f:
        f.write("\n".join(out) + "\n")
    img = Image.new("RGB", (W, H))
    px = img.load()
    for y in range(H):
        for x in range(W):
            h = PAL[g[y][x]]
            px[x, y] = (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    img.resize((W * 8, H * 8), Image.NEAREST).save(os.path.join(ROOT, "docs", "preview", "stage_sanmun.png"))
    print("runs (first-frame paint calls):", runs)

if __name__ == "__main__":
    main()
