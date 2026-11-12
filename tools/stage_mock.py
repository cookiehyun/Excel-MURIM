"""Preview a stage with both fighters on it -> docs/preview/stage_<name>_mock.png
usage: python tools/stage_mock.py <stage name> [danwooSprite cheolsanSprite]"""
import os, sys
from PIL import Image
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load(path):
    pal, spr, stage = {}, {}, None
    L = open(path, newline="").read().replace("\r", "").split("\n")
    i = 0
    while i < len(L):
        p = L[i].split(" ")
        if p[0] == "PAL":
            pal[p[1]] = tuple(int(p[2][k:k + 2], 16) for k in (0, 2, 4))
        elif p[0] == "SPRITE":
            w, h, ax, ay = map(int, p[2:6]); spr[p[1]] = (w, h, ax, ay, L[i + 1:i + 1 + h]); i += h
        elif p[0] == "STAGE":
            w, h, g = int(p[2]), int(p[3]), int(p[4]); stage = (w, h, g, L[i + 1:i + 1 + h]); i += h
        i += 1
    return pal, spr, stage

def main():
    name = sys.argv[1]
    s1 = sys.argv[2] if len(sys.argv) > 2 else "lp_1"
    s2 = sys.argv[3] if len(sys.argv) > 3 else "idle_0"
    sp, _, (W, H, G, rows) = load(os.path.join(ROOT, "data", f"stage_{name}.txt"))
    img = Image.new("RGB", (W, H)); px = img.load()
    for y in range(H):
        for x in range(W):
            px[x, y] = sp[rows[y][x]]
    for f, nm, X, flip in (("danwoo", s1, int(W * 0.42), False), ("cheolsan", s2, int(W * 0.56), True)):
        pal, spr, _ = load(os.path.join(ROOT, "data", f + ".txt"))
        w, h, ax, ay, r = spr[nm]
        for sy in range(h):
            for sx in range(w):
                c = r[sy][sx]
                if c != ".":
                    x = X + ax - sx if flip else X - ax + sx
                    y = G - ay + sy
                    if 0 <= x < W and 0 <= y < H:
                        px[x, y] = pal[c]
    out = os.path.join(ROOT, "docs", "preview", f"stage_{name}_mock.png")
    img.resize((W * 5, H * 5), Image.NEAREST).save(out)
    print(out)

if __name__ == "__main__":
    main()
