"""Preview move animations from data/<char>.txt as a strip (frames in order, anim timing ignored).
usage: python3 tools/anim_preview.py <char> <anim-or-frame>... -o out.png"""
import sys, os
from PIL import Image, ImageDraw
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load(path):
    pal, spr, anim, cur = {}, {}, {}, None
    for ln in open(path):
        ln = ln.rstrip("\r\n")
        if ln.startswith("PAL "):
            _, c, h = ln.split(" ", 2); pal[c] = h; cur = None
        elif ln.startswith("SPRITE "):
            t = ln.split(); cur = t[1]; spr[cur] = (int(t[4]), int(t[5]), [])
        elif ln.startswith("ANIM "):
            t = ln.split(); anim[t[1]] = t[3:]; cur = None
        elif ln.startswith("#"):
            cur = None
        elif cur:
            spr[cur][2].append(ln)
    return pal, spr, anim

def main():
    args = sys.argv[1:]
    out = args[args.index("-o") + 1]; args = args[:args.index("-o")]
    ch, names = args[0], args[1:]
    pal, spr, anim = load(os.path.join(ROOT, "data", ch + ".txt"))
    frames = []
    for n in names:
        frames += anim.get(n, [n])
    S, cw, base = 6, 64, 56
    im = Image.new("RGB", (cw * S * len(frames), (base + 14) * S), (150, 160, 175))
    d = ImageDraw.Draw(im)
    for i, f in enumerate(frames):
        ax, ay, rows = spr[f]
        ox = i * cw * S + (cw // 2 - ax) * S
        oy = (base - ay) * S
        for y, r in enumerate(rows):
            for x, c in enumerate(r):
                if c != ".":
                    d.rectangle([ox + x * S, oy + y * S, ox + x * S + S - 1, oy + y * S + S - 1], fill="#" + pal[c])
        d.line([(i * cw * S, base * S), ((i + 1) * cw * S - 8, base * S)], fill=(90, 90, 110), width=2)
        d.text((i * cw * S + 6, base * S + 8), f, fill="black")
    im.save(out)

main()
