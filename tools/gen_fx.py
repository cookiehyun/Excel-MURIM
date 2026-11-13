"""Generate effect sprites (projectiles) -> data/fx.txt + preview.
Anchor = sprite center. usage: python tools/gen_fx.py"""
import math, os
from PIL import Image
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAL = {"k": "0E2A3A", "c": "4FD8E8", "C": "9AF2F2", "w": "FFFFFF", "b": "2A7FB8",
       "g": "6AE070", "G": "C8FFB0", "x": "14201A", "v": "4A8A50",
       "F": "E04050", "y": "F0C850", "Q": "FFE6A0", "z": "2A1018",
       "r": "C01E28", "R": "FF6A3C", "o": "300610"}

def wave(w, h, phase, big):
    g = [["."] * w for _ in range(h)]
    cx, cy = w * 0.62, h / 2
    R = h * 0.48
    for y in range(h):
        for x in range(w):
            dx, dy = (x + 0.5 - cx) / R, (y + 0.5 - cy) / R
            d = math.hypot(dx * 1.25, dy)
            if d > 1.0:
                continue
            # crescent: cut out an offset inner disc on the back side
            inner = math.hypot((x + 0.5 - (cx - R * 0.95)) / (R * 0.8), (y + 0.5 - cy) / (R * 0.85))
            if inner < 1.0:
                continue
            ang = math.atan2(dy, dx) + phase
            swirl = (math.sin(ang * 3 + d * 6) + 1) / 2
            if d > 0.86:
                g[y][x] = "b"
            elif swirl > 0.72:
                g[y][x] = "w"
            elif swirl > 0.35:
                g[y][x] = "C"
            else:
                g[y][x] = "c"
    # trailing streaks
    for k in range(3 if big else 2):
        yy = int(cy + (k - (1 if big else 0.5)) * (h / 4))
        ln = int(w * 0.35) - k * 2 + (2 if phase else 0)
        for x in range(0, max(0, ln)):
            if 0 <= yy < h and g[yy][x] == "." and (x + k + int(phase * 3)) % 3 != 0:
                g[yy][x] = "c" if x > ln // 2 else "b"
    # outline
    out = [row[:] for row in g]
    for y in range(h):
        for x in range(w):
            if g[y][x] == ".":
                for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < w and 0 <= yy < h and g[yy][xx] in "wCc":
                        out[y][x] = "k"; break
    return out

def spin_fan(n, phase):
    """open fan spinning: half-disc that flips orientation between frames"""
    g = [["."] * n for _ in range(n)]
    c = (n - 1) / 2
    for y in range(n):
        for x in range(n):
            dx, dy = x - c, y - c
            d = math.hypot(dx, dy)
            if d > c + 0.3 or d < 1.0:
                continue
            ang = math.degrees(math.atan2(dy, dx)) + (0 if phase == 0 else 90)
            if (ang % 360) < 180:                      # half disc
                if d > c - 0.8:
                    g[y][x] = "Q"
                elif int((ang % 180) / 30) % 2 == 0:
                    g[y][x] = "F"
                else:
                    g[y][x] = "y"
    g[int(c)][int(c)] = "y"
    out = [row[:] for row in g]
    for y in range(n):
        for x in range(n):
            if g[y][x] == ".":
                for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < n and 0 <= yy < n and g[yy][xx] != ".":
                        out[y][x] = "z"; break
    return out


def needles(w, h, phase):
    """three poison needles flying in a tight fan, green glint + faint trail"""
    g = [["."] * w for _ in range(h)]
    for k, y in enumerate((1, 3, 5)):
        if y >= h:
            continue
        x0 = 3 + (k % 2) + phase
        for x in range(x0, min(w, x0 + 7)):
            g[y][x] = "G" if x >= x0 + 5 else "g"
        for x in range(max(0, x0 - 3), x0):
            if (x + phase) % 2 == 0:
                g[y][x] = "v"
    out = [row[:] for row in g]
    for y in range(h):
        for x in range(w):
            if g[y][x] == ".":
                for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < w and 0 <= yy < h and g[yy][xx] in "gG":
                        out[y][x] = "x"; break
    return out


def blade(w, h, phase):
    """sword wave: thin silver crescent standing upright, with a light trail"""
    g = [["."] * w for _ in range(h)]
    cx, cy, R = w * 0.35, h / 2, h * 0.5
    for y in range(h):
        for x in range(w):
            d = math.hypot((x + 0.5 - cx) / (R * 0.55), (y + 0.5 - cy) / R)
            inner = math.hypot((x + 0.5 - (cx - 2.2)) / (R * 0.5), (y + 0.5 - cy) / (R * 0.92))
            if d <= 1 and inner > 1:
                g[y][x] = "w" if (x + y + phase) % 3 else "C"
    for y in range(2, h - 2, 3):                   # trail streaks
        for x in range(0, int(cx) - 1):
            if g[y][x] == "." and (x + phase) % 2 == 0:
                g[y][x] = "c"
    out = [row[:] for row in g]
    for y in range(h):
        for x in range(w):
            if g[y][x] == ".":
                for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < w and 0 <= yy < h and g[yy][xx] in "wC":
                        out[y][x] = "k"; break
    return out


def demon_orb(n, phase):
    """Cheonma's demon orb: crimson fireball with a gold core and flickering tail"""
    w, h = n + 6, n
    g = [["."] * w for _ in range(h)]
    cx, cy, R = w - n / 2 - 0.5, (h - 1) / 2, n / 2
    for y in range(h):
        for x in range(w):
            d = math.hypot(x - cx, y - cy)
            if d <= R * 0.45:
                g[y][x] = "Q" if (x + y + phase) % 3 else "y"
            elif d <= R * 0.75:
                g[y][x] = "R"
            elif d <= R:
                g[y][x] = "r"
            elif x < cx and abs(y - cy) <= R * (0.8 - (cx - x) / (w * 1.1)) and (x + y * 2 + phase) % 3:
                g[y][x] = "r" if (cx - x) < 4 else "o"     # flame tail
    out = [row[:] for row in g]
    for y in range(h):
        for x in range(w):
            if g[y][x] == ".":
                for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < w and 0 <= yy < h and g[yy][xx] in "QyRr":
                        out[y][x] = "o"; break
    return out


def main():
    out = ["# effect sprites - generated by tools/gen_fx.py, do not edit by hand"]
    for k, v in PAL.items():
        out.append(f"PAL {k} {v}")
    sprites = {"fan_0": spin_fan(11, 0), "fan_1": spin_fan(11, 1),
               "needle_0": needles(12, 6, 0), "needle_1": needles(12, 6, 1),
               "blade_0": blade(9, 16, 0), "blade_1": blade(9, 16, 1),
               "orb_0": demon_orb(11, 0), "orb_1": demon_orb(11, 1),
               "wind_0": wave(16, 12, 0.0, False), "wind_1": wave(16, 12, 1.0, False),
               "bigwind_0": wave(26, 20, 0.0, True), "bigwind_1": wave(26, 20, 1.0, True)}
    prev = Image.new("RGBA", (230, 26), (40, 40, 60, 255))
    px0 = 0
    for name, g in sprites.items():
        h, w = len(g), len(g[0])
        out.append(f"SPRITE {name} {w} {h} {w // 2} {h // 2}")
        out += ["".join(r) for r in g]
        for y in range(h):
            for x in range(w):
                c = g[y][x]
                if c != ".":
                    hx = PAL[c]
                    prev.putpixel((px0 + x, y), (int(hx[:2], 16), int(hx[2:4], 16), int(hx[4:], 16), 255))
        px0 += w + 2
    out.append("ANIM wind 2 wind_0 wind_1")
    out.append("ANIM proj_muyeong 2 blade_0 blade_1")
    out.append("ANIM proj_heukrim 2 needle_0 needle_1")
    out.append("ANIM proj_hongryeon 1 fan_0 fan_1")
    out.append("ANIM proj_cheonma 2 orb_0 orb_1")
    out.append("ANIM bigwind 2 bigwind_0 bigwind_1")
    with open(os.path.join(os.environ.get("OUT_DIR", os.path.join(ROOT, "data")), "fx.txt"), "w", newline="\r\n") as f:
        f.write("\n".join(out) + "\n")
    prev.resize((prev.width * 8, prev.height * 8), Image.NEAREST).save(os.path.join(os.environ.get("OUT_DIR", os.path.join(ROOT, "docs", "preview")), "fx_sheet.png"))

if __name__ == "__main__":
    main()
