"""Generate fighter sprites from a simple 2D rig.

Each pose = joint angles. The rig is rasterized to a pixel grid, outlined,
and written to data/<name>.txt as palette-indexed text that VBA loads at
runtime. Also writes enlarged PNG previews to docs/preview/.

Angles (degrees) for limbs are absolute, measured from straight DOWN,
positive = toward facing direction (+x). Torso lean: from straight UP,
positive = forward. Sprites face right; the game mirrors them.

usage: python tools/gen_sprites.py
"""
import math
import os
import numpy as np
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANVAS = 72          # working canvas (px); sprite is cropped afterwards
AX, AY = 36, 62      # anchor in canvas = ground point under the hip

# ---------------------------------------------------------------- geometry

def dirv(angle_from_down):
    a = math.radians(angle_from_down)
    return math.sin(a), math.cos(a)

def seg_mask(p, q, r):
    ys, xs = np.mgrid[0:CANVAS, 0:CANVAS]
    px, py = xs + 0.5, ys + 0.5
    vx, vy = q[0] - p[0], q[1] - p[1]
    l2 = vx * vx + vy * vy
    if l2 == 0:
        t = np.zeros_like(px)
    else:
        t = np.clip(((px - p[0]) * vx + (py - p[1]) * vy) / l2, 0, 1)
    dx = px - (p[0] + t * vx)
    dy = py - (p[1] + t * vy)
    return dx * dx + dy * dy <= r * r

def circle_mask(c, r):
    return seg_mask(c, c, r)

def poly_mask(pts):
    ys, xs = np.mgrid[0:CANVAS, 0:CANVAS]
    px, py = xs + 0.5, ys + 0.5
    inside = np.zeros((CANVAS, CANVAS), bool)
    n = len(pts)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        cond = (y1 > py) != (y2 > py)
        xint = (x2 - x1) * (py - y1) / ((y2 - y1) if y2 != y1 else 1e-9) + x1
        inside ^= cond & (px < xint)
    return inside

def dilate(m):
    d = m.copy()
    d[1:, :] |= m[:-1, :]
    d[:-1, :] |= m[1:, :]
    d[:, 1:] |= m[:, :-1]
    d[:, :-1] |= m[:, 1:]
    return d

# ---------------------------------------------------------------- character specs

DANWOO = dict(
    name="danwoo",
    thigh=7.5, shin=7.5, torso=10.0, upper=5.5, fore=5.0, head_r=3.6,
    w_thigh=2.0, w_shin=1.7, w_upper=1.6, w_fore=1.4, w_hand=1.6,
    chest=4.0, waist=3.0, skirt=4.5,
    palette={
        "k": "1B1626",  # outline
        "s": "F2C6A0", "S": "C98D6B",          # skin
        "h": "22202A",                          # hair
        "b": "2F6FB5", "B": "1F4A80",          # robe blue / dark
        "w": "E8E8EF",                          # white trim
        "d": "2A2230",                          # belt
        "p": "3B3F4F", "P": "2A2D39",          # pants
        "n": "4A3324",                          # boots
        "r": "C23B3B",                          # headband red
    },
    look=dict(arm_front="b", arm_back="B", arm_front2="b", arm_back2="B",
              leg_front="p", leg_back="P", boot="n",
              torso="b", torso_dark="B", skirt="b", skirt_dark="B",
              belt="d", trim="w", cuff="w"),
    head="danwoo",
)

CHEOLSAN = dict(
    name="cheolsan",
    thigh=8.5, shin=8.5, torso=12.5, upper=6.6, fore=6.0, head_r=3.9,
    w_thigh=3.0, w_shin=2.4, w_upper=2.5, w_fore=2.2, w_hand=2.3,
    chest=6.4, waist=5.0, skirt=7.0,
    palette={
        "k": "1E1410",
        "s": "D9A074", "S": "A8704F",
        "h": "2A221C",
        "o": "C4652A", "O": "8D4420",          # robe
        "r": "7A1F1F",                          # sash
        "e": "5A3A1E", "E": "9A7040",          # prayer beads
        "g": "9A8F7C", "G": "6E6555",          # leg wraps
        "n": "3B2A1E",
        "w": "F2CDA8",                          # scalp highlight
    },
    look=dict(arm_front="s", arm_back="O", arm_front2="s", arm_back2="S",
              leg_front="g", leg_back="G", boot="n",
              torso="o", torso_dark="O", skirt="o", skirt_dark="O",
              belt="r", trim=None, cuff=None),
    head="cheolsan",
)

# ---------------------------------------------------------------- poses
# legs: (thigh, shin) ; arms: (upper, fore) ; F = front/lead (drawn on top), B = back
GUARD_F = (50, 135)
GUARD_B = (20, 150)
STANCE_F = (25, 5)
STANCE_B = (-25, -10)
CROUCH_F = (75, -15)
CROUCH_B = (35, -70)
TUCK_F = (100, -10)
TUCK_B = (60, -45)

def P(lean=8, lf=STANCE_F, lb=STANCE_B, af=GUARD_F, ab=GUARD_B,
      ground="feet", hip_y=None, rot=0, open_hand=False):
    return dict(lean=lean, lf=lf, lb=lb, af=af, ab=ab, ground=ground,
                hip_y=hip_y, rot=rot, open_hand=open_hand)

BASE_POSES = {
    "idle_0": P(),
    "idle_1": P(lf=(29, 9), lb=(-29, -14), af=(48, 130), ab=(18, 145)),
    "walk_0": P(lf=(25, 5), lb=(-20, -25)),
    "walk_1": P(lf=(8, 0), lb=(-5, -40)),
    "walk_2": P(lf=(-18, -25), lb=(25, 5)),
    "walk_3": P(lf=(-5, -40), lb=(8, 0)),
    "crouch": P(lean=20, lf=CROUCH_F, lb=CROUCH_B),
    "jump": P(lean=10, lf=TUCK_F, lb=TUCK_B, ground="air"),
    "lp_0": P(lean=10, af=(45, 110)),
    "lp_1": P(lean=16, af=(92, 90)),
    "hp_0": P(lean=-14, lf=(30, 5), lb=(-30, -15), af=(-50, 150), ab=(40, 140)),
    "hp_1": P(lean=24, lf=(40, 0), lb=(-40, -40), af=(90, 88), ab=(10, 120)),
    "lk_0": P(lean=-5, lf=(90, 10), lb=(-8, 0)),
    "lk_1": P(lean=-10, lf=(82, 86), lb=(-8, 0)),
    "hk_0": P(lean=-15, lf=(120, 30), lb=(-5, 0), af=(60, 150)),
    "hk_1": P(lean=-28, lf=(122, 128), lb=(-5, 0), af=(40, 160), ab=(-30, 20)),
    "clp": P(lean=20, lf=CROUCH_F, lb=CROUCH_B, af=(88, 86)),
    "chp_0": P(lean=24, lf=CROUCH_F, lb=CROUCH_B, af=(-30, 30)),
    "chp_1": P(lean=2, lf=(45, -5), lb=(10, -50), af=(160, 176), ab=(30, 120)),
    "clk": P(lean=14, lf=(68, 74), lb=(55, -75), ground="hip_low"),
    "chk": P(lean=45, lf=(70, 78), lb=(88, -88), af=(25, 15), ab=(35, 25), ground="hip_low"),
    "jp": P(lean=12, lf=TUCK_F, lb=TUCK_B, af=(60, 55), ground="air"),
    "jk": P(lean=-12, lf=(68, 64), lb=(92, -30), ground="air"),
    "guard": P(lean=-6, lf=(20, 5), lb=(-30, -15), af=(30, 165), ab=(45, 150)),
    "cguard": P(lean=12, lf=CROUCH_F, lb=CROUCH_B, af=(40, 165), ab=(55, 150)),
    "hit_0": P(lean=-28, lf=(15, 0), lb=(-30, -22), af=(-20, 30), ab=(-45, -10)),
    "hit_1": P(lean=-16, lf=(20, 3), lb=(-28, -18), af=(10, 60), ab=(-25, 20)),
    "chit": P(lean=-12, lf=CROUCH_F, lb=CROUCH_B, af=(10, 60), ab=(-20, 30)),
    "down_0": P(lean=0, lf=(40, 20), lb=(10, -20), af=(160, 150), ab=(130, 170), ground="air", rot=-55),
    "down_1": P(lean=0, lf=(5, 0), lb=(-5, -5), af=(170, 175), ab=(150, 160), ground="bbox", rot=-90),
    "throw_0": P(lean=15, af=(80, 85), ab=(75, 80), open_hand=True),
    "throw_1": P(lean=-28, lf=(-10, 0), lb=(-40, -20), af=(160, 175), ab=(150, 170), open_hand=True),
    "sp_0": P(lean=-6, af=(-15, 80), ab=(-25, 75), open_hand=True),
    "sp_1": P(lean=20, lf=(40, 0), lb=(-40, -40), af=(90, 90), ab=(84, 86), open_hand=True),
}

CHEOLSAN_OVERRIDES = {
    "idle_0": P(lean=12, lf=(32, 8), lb=(-32, -12), af=(60, 105), ab=(30, 120), open_hand=True),
    "idle_1": P(lean=13, lf=(35, 11), lb=(-35, -15), af=(58, 100), ab=(28, 115), open_hand=True),
    "walk_0": P(lean=12, lf=(30, 8), lb=(-22, -25), af=(60, 105), ab=(30, 120), open_hand=True),
    "walk_1": P(lean=12, lf=(10, 0), lb=(-5, -40), af=(60, 105), ab=(30, 120), open_hand=True),
    "walk_2": P(lean=12, lf=(-20, -25), lb=(30, 8), af=(60, 105), ab=(30, 120), open_hand=True),
    "walk_3": P(lean=12, lf=(-5, -40), lb=(10, 0), af=(60, 105), ab=(30, 120), open_hand=True),
    "sp_0": P(lean=10, af=(110, 120), ab=(100, 125), open_hand=True),
    "sp_1": P(lean=30, lf=(40, 0), lb=(-40, -40), af=(95, 75), ab=(90, 80), open_hand=True),
}

ANIMS = [  # name, ticks per frame, frames
    ("idle", 6, ["idle_0", "idle_1"]),
    ("walk", 3, ["walk_0", "walk_1", "walk_2", "walk_3"]),
    ("crouch", 1, ["crouch"]),
    ("jump", 1, ["jump"]),
]

# ---------------------------------------------------------------- rasterizer

class Canvas:
    def __init__(self):
        self.c = np.full((CANVAS, CANVAS), "", dtype="<U1")

    def filled(self):
        return self.c != ""

    def put(self, mask, color, inner_outline=False):
        if inner_outline:
            ring = dilate(mask) & ~mask & self.filled()
            self.c[ring] = "k"
        self.c[mask] = color

    def outline(self):
        f = self.filled()
        ring = dilate(f) & ~f
        self.c[ring] = "k"


def build(spec, pose):
    L = spec
    lk = spec["look"]
    rot = math.radians(pose["rot"])

    def R(p):  # rotate body-space point around hip, then translate to canvas
        x, y = p
        rx = x * math.cos(rot) - y * math.sin(rot)
        ry = x * math.sin(rot) + y * math.cos(rot)
        return rx, ry

    # --- body space: hip at (0,0), y down
    def limb(origin, a1, l1, a2, l2):
        d1 = dirv(a1)
        j = (origin[0] + d1[0] * l1, origin[1] + d1[1] * l1)
        d2 = dirv(a2)
        e = (j[0] + d2[0] * l2, j[1] + d2[1] * l2)
        return j, e

    lean = math.radians(pose["lean"])
    up = (math.sin(lean), -math.cos(lean))
    neck = (up[0] * L["torso"], up[1] * L["torso"])
    shoulder = (up[0] * (L["torso"] - 1.5), up[1] * (L["torso"] - 1.5))
    head_c = (neck[0] + up[0] * (L["head_r"] + 0.6), neck[1] + up[1] * (L["head_r"] + 0.6))

    hipF = (0.8, 0.0)
    hipB = (-0.8, 0.0)
    kF, fF = limb(hipF, pose["lf"][0], L["thigh"], pose["lf"][1], L["shin"])
    kB, fB = limb(hipB, pose["lb"][0], L["thigh"], pose["lb"][1], L["shin"])
    eF, hF = limb(shoulder, pose["af"][0], L["upper"], pose["af"][1], L["fore"])
    eB, hB = limb(shoulder, pose["ab"][0], L["upper"], pose["ab"][1], L["fore"])

    pts = dict(neck=neck, shoulder=shoulder, head=head_c, kF=kF, fF=fF, kB=kB, fB=fB,
               eF=eF, hF=hF, eB=eB, hB=hB, hip=(0, 0))
    pts = {k: R(v) for k, v in pts.items()}

    # --- place hip on canvas
    if pose["ground"] in ("feet", "hip_low"):
        foot_r = L["w_shin"]
        lowest = max(pts["fF"][1], pts["fB"][1], pts["kF"][1], pts["kB"][1]) + foot_r
        hip_y = -lowest
    elif pose["ground"] == "air":
        hip_y = -(L["thigh"] + L["shin"]) + 0.5
    else:
        hip_y = -8  # bbox: re-seated after rasterizing
    ox, oy = AX, AY + hip_y
    T = lambda p: (p[0] + ox, p[1] + oy)
    pt = {k: T(v) for k, v in pts.items()}

    cv = Canvas()
    shin_foot = lambda knee, foot: (foot[0] + (foot[0] - knee[0]) * 0 + 2.2 * math.cos(math.atan2(foot[1] - knee[1], foot[0] - knee[0]) - math.pi / 2) * -1,
                                    foot[1] + 2.2 * math.sin(math.atan2(foot[1] - knee[1], foot[0] - knee[0]) - math.pi / 2) * -1)

    def draw_leg(knee, foot, col, inner):
        hipp = pt["hip"]
        m = seg_mask(hipp, knee, L["w_thigh"]) | seg_mask(knee, foot, L["w_shin"])
        cv.put(m, col, inner)
        toe = shin_foot(knee, foot)
        cv.put(seg_mask(foot, toe, L["w_shin"] * 0.75) | circle_mask(foot, L["w_shin"] * 0.9), lk["boot"], inner)

    def draw_arm(elbow, hand, col_u, col_f, inner):
        sh = pt["shoulder"]
        cv.put(seg_mask(sh, elbow, L["w_upper"]) | seg_mask(elbow, hand, L["w_fore"]), col_u, inner)
        if col_f != col_u:
            cv.put(seg_mask(elbow, hand, L["w_fore"]), col_f, False)
        if lk["cuff"]:
            d = (hand[0] - elbow[0], hand[1] - elbow[1])
            n = math.hypot(*d) or 1
            cuff = (hand[0] - d[0] / n * 1.6, hand[1] - d[1] / n * 1.6)
            cv.put(circle_mask(cuff, L["w_fore"] + 0.2), lk["cuff"], False)
        hr = L["w_hand"] + (0.4 if pose["open_hand"] else 0)
        cv.put(circle_mask(hand, hr), "s" if inner else "S", inner)

    # back arm, back leg
    draw_arm(pt["eB"], pt["hB"], lk["arm_back"], lk["arm_back2"], False)
    draw_leg(pt["kB"], pt["fB"], lk["leg_back"], False)

    # torso quad
    nx, ny = pt["neck"]
    hx, hy = pt["hip"]
    ax_, ay_ = nx - hx, ny - hy
    n = math.hypot(ax_, ay_)
    px_, py_ = -ay_ / n, ax_ / n  # perpendicular
    quad = [(nx + px_ * L["chest"], ny + py_ * L["chest"]),
            (nx - px_ * L["chest"], ny - py_ * L["chest"]),
            (hx - px_ * L["waist"], hy - py_ * L["waist"]),
            (hx + px_ * L["waist"], hy + py_ * L["waist"])]
    tm = poly_mask(quad) | circle_mask(pt["shoulder"], L["chest"] * 0.85)
    cv.put(tm, lk["torso"])
    # shade the back half of the torso
    ys, xs = np.mgrid[0:CANVAS, 0:CANVAS]
    side = (xs + 0.5 - hx) * (-px_) + (ys + 0.5 - hy) * (-py_)   # >0 toward front? use axis-perp
    # front direction in world = +x rotated; pick perpendicular that points forward
    fwd = 1 if (-px_) * math.cos(rot) + 0 > 0 else -1
    back_half = tm & (side * fwd < -L["chest"] * 0.25)
    cv.c[back_half] = lk["torso_dark"]

    # skirt / robe flaps
    sk = L["skirt"]
    flap = [(hx - L["waist"] - 0.5, hy - 1), (hx + L["waist"] + 0.5, hy - 1),
            (hx + L["waist"] + 1.5, hy + sk), (hx - L["waist"] - 1.0, hy + sk)]
    flap = [R((p[0] - hx, p[1] - hy)) for p in flap]
    flap = [(p[0] + hx, p[1] + hy) for p in flap]

    draw_leg(pt["kF"], pt["fF"], lk["leg_front"], True)
    fm = poly_mask(flap)
    cv.put(fm, lk["skirt"], True)
    cv.c[fm & (xs + 0.5 < hx - 0.5)] = lk["skirt_dark"]
    # belt
    belt_c = (hx + ax_ / n * 1.2, hy + ay_ / n * 1.2)
    bm = seg_mask((belt_c[0] + px_ * L["waist"], belt_c[1] + py_ * L["waist"]),
                  (belt_c[0] - px_ * L["waist"], belt_c[1] - py_ * L["waist"]), 0.9) & (tm | fm)
    cv.put(bm, lk["belt"])

    # head
    hc = pt["head"]
    r = L["head_r"]
    hm = circle_mask(hc, r)
    cv.put(hm | seg_mask(pt["neck"], hc, 1.5), "s")
    cx, cy = hc
    rel_x = xs + 0.5 - cx
    rel_y = ys + 0.5 - cy
    # head-local axes (follow lean + rotation)
    ang = lean + rot
    ux, uy = math.sin(ang), -math.cos(ang)          # head up
    fx, fy = math.cos(ang), math.sin(ang)            # head forward
    hu = rel_x * ux + rel_y * uy
    hf = rel_x * fx + rel_y * fy
    if spec["head"] == "danwoo":
        hair = hm & ((hu > 0.6) | (hf < -1.0))
        cv.c[hair] = "h"
        band = hm & (hu > 0.2) & (hu <= 1.4) & (hf > -1.0)
        cv.c[band] = "r"
        knot = circle_mask((cx + ux * (r + 0.6) - fx * 1.2, cy + uy * (r + 0.6) - fy * 1.2), 1.5)
        cv.put(knot, "h")
        tail = seg_mask((cx - fx * r, cy - fy * r + 0.0 + uy * -0.6), (cx - fx * (r + 3.0) - ux * 1.5, cy - fy * (r + 3.0) - uy * 1.5), 0.6)
        cv.put(tail & ~hm, "r")
        trim = seg_mask(pt["neck"], (pt["neck"][0] + fx * 1.5 - ux * 2.5, pt["neck"][1] + fy * 1.5 - uy * 2.5), 0.7) & tm
        cv.c[trim] = "w"
    else:
        brow = hm & (hu > 0.3) & (hu <= 1.0) & (hf > 0.8)
        cv.c[brow] = "h"
        beard = hm & (hu < -2.0) & (hf > 0.0)
        cv.c[beard] = "h"
        shine = hm & (hu > 2.0) & (hf > -1.5) & (hf < 0.2)
        cv.c[shine] = "w" if "w" in spec["palette"] else "s"
        # diagonal sash across chest
        sash = seg_mask(pt["shoulder"], (hx + px_ * 0 + fx * L["waist"], hy + fy * L["waist"]), 1.0) & tm
        cv.c[sash] = "r"
        # prayer beads around the neck/chest
        bc = (pt["neck"][0] - ux * 1.5, pt["neck"][1] - uy * 1.5)
        for i, t in enumerate(np.linspace(-1.3, 1.3, 7)):
            bx = bc[0] + fx * math.sin(t) * 3.6 - ux * (math.cos(t) * 2.6)
            by = bc[1] + fy * math.sin(t) * 3.6 - uy * (math.cos(t) * 2.6)
            cv.c[circle_mask((bx, by), 0.9)] = "E" if i % 2 else "e"
    # eye
    eye = (cx + fx * (r * 0.55) + ux * (-0.1), cy + fy * (r * 0.55) + uy * (-0.1))
    ex, ey = int(eye[0]), int(eye[1])
    if 0 <= ex < CANVAS and 0 <= ey < CANVAS and cv.c[ey, ex] != "":
        cv.c[ey, ex] = "k"

    # front arm last (on top), with inner outline
    draw_arm(pt["eF"], pt["hF"], lk["arm_front"], lk["arm_front2"], True)

    cv.outline()

    grid = cv.c
    if pose["ground"] == "bbox":
        rows = np.where((grid != "").any(axis=1))[0]
        shift = AY - 1 - rows.max()
        grid = np.roll(grid, shift, axis=0)
    return grid, pt


def crop(grid):
    f = grid != ""
    rows = np.where(f.any(axis=1))[0]
    cols = np.where(f.any(axis=0))[0]
    y0, y1 = rows.min(), max(rows.max(), AY - 1)
    x0, x1 = cols.min(), cols.max()
    return grid[y0:y1 + 1, x0:x1 + 1], AX - x0, AY - y0   # anchor relative to crop


def to_png(grid, palette, scale=8):
    h, w = grid.shape
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    px = img.load()
    for y in range(h):
        for x in range(w):
            ch = grid[y, x]
            if ch:
                hx = palette[ch]
                px[x, y] = (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16), 255)
    return img.resize((w * scale, h * scale), Image.NEAREST)


def generate(spec, overrides):
    poses = dict(BASE_POSES)
    poses.update(overrides)
    out = []
    out.append(f"# {spec['name']} sprites - generated by tools/gen_sprites.py, do not edit by hand")
    out.append("# PAL <char> <RRGGBB> ; SPRITE <name> <w> <h> <anchorX> <anchorY> ; rows ('.' = transparent)")
    out.append("# ANIM <name> <ticksPerFrame> <frames...>")
    for ch, hx in spec["palette"].items():
        out.append(f"PAL {ch} {hx}")
    previews = []
    points = []
    for name, pose in poses.items():
        grid, pt = build(spec, pose)
        if pose["ground"] != "bbox":
            rel = lambda k: (round(pt[k][0] - AX, 1), round(pt[k][1] - AY, 1))
            points.append(f"{name}: hF={rel('hF')} hB={rel('hB')} fF={rel('fF')} fB={rel('fB')} head={rel('head')}")
        g, ax, ay = crop(grid)
        h, w = g.shape
        out.append(f"SPRITE {name} {w} {h} {ax} {ay}")
        for row in g:
            out.append("".join(c if c else "." for c in row))
        previews.append((name, g))
    for name, tpf, frames in ANIMS:
        out.append(f"ANIM {name} {tpf} " + " ".join(frames))
    path = os.path.join(ROOT, "data", f"{spec['name']}.txt")
    with open(path, "w", newline="\r\n") as f:
        f.write("\n".join(out) + "\n")

    with open(os.path.join(ROOT, "docs", f"{spec['name']}_points.txt"), "w") as f:
        f.write("# limb end positions relative to anchor (x right, y down) - reference for hitboxes\n")
        f.write("\n".join(points) + "\n")

    # contact sheet preview
    scale = 6
    cell_w = max(g.shape[1] for _, g in previews) + 4
    cell_h = max(g.shape[0] for _, g in previews) + 4
    cols = 8
    rows = math.ceil(len(previews) / cols)
    sheet = Image.new("RGBA", (cols * cell_w * scale, rows * cell_h * scale), (70, 80, 100, 255))
    for i, (name, g) in enumerate(previews):
        im = to_png(g, spec["palette"], scale)
        cx = (i % cols) * cell_w * scale + (cell_w * scale - im.width) // 2
        cy = (i // cols) * cell_h * scale + (cell_h * scale - im.height) - 2 * scale
        sheet.alpha_composite(im, (cx, cy))
    sheet.save(os.path.join(ROOT, "docs", "preview", f"{spec['name']}_sheet.png"))
    return [n for n, _ in previews]


if __name__ == "__main__":
    for spec, ov in ((DANWOO, {}), (CHEOLSAN, CHEOLSAN_OVERRIDES)):
        names = generate(spec, ov)
        print(spec["name"], len(names), "sprites")
