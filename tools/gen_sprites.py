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

MUYEONG = dict(
    name="muyeong",
    thigh=7.8, shin=7.8, torso=10.6, upper=5.6, fore=5.2, head_r=3.4,
    w_thigh=1.8, w_shin=1.5, w_upper=1.5, w_fore=1.3, w_hand=1.5,
    chest=3.7, waist=2.7, skirt=6.5,
    palette={
        "k": "15141C",
        "s": "EBC09C", "S": "C08A68",
        "h": "1A1A22",
        "g": "4C5668", "G": "343B4A",          # slate robe
        "w": "E6E4DA",                          # white sash / collar / hair tie
        "p": "34343E", "P": "26262E",
        "n": "3A2C24",
        "W": "DCE4EC", "Q": "FFFFFF",          # blade / glint
        "y": "C9A040", "m": "5A2A20",          # guard / hilt
        "r": "C23434",                          # tassel
    },
    look=dict(arm_front="g", arm_back="G", arm_front2="g", arm_back2="G",
              leg_front="p", leg_back="P", boot="n",
              torso="g", torso_dark="G", skirt="g", skirt_dark="G",
              belt="w", trim="w", cuff="w"),
    head="muyeong",
    weapon=dict(kind="sword", length=14),
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
      ground="feet", hip_y=None, rot=0, open_hand=False, wa=None):
    """wa = weapon angle (absolute, from down, + forward); None = continue the front forearm"""
    return dict(lean=lean, lf=lf, lb=lb, af=af, ab=ab, ground=ground,
                hip_y=hip_y, rot=rot, open_hand=open_hand, wa=wa)

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
    "dash": P(lean=28, lf=(60, -5), lb=(-45, -60), af=(70, 120), ab=(-35, 50)),
    "bdash": P(lean=-18, lf=(45, -15), lb=(15, -40), af=(40, 150), ab=(20, 140), ground="air"),
    "dp_0": P(lean=22, lf=CROUCH_F, lb=CROUCH_B, af=(10, 50), ab=(-20, 40), open_hand=True),
    "dp_1": P(lean=-6, lf=(85, -15), lb=(-10, -25), af=(168, 178), ab=(140, 110), ground="air", open_hand=True),
    "rush_0": P(lean=-12, lf=(30, 5), lb=(-35, -20), af=(20, 100), ab=(-25, 70)),
    "rush_1": P(lean=42, lf=(55, 10), lb=(-55, -40), af=(115, 55), ab=(95, 35)),
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

SEOLHWA = dict(
    name="seolhwa",
    thigh=7.6, shin=8.0, torso=9.6, upper=5.0, fore=4.6, head_r=3.3,
    w_thigh=1.8, w_shin=1.45, w_upper=1.35, w_fore=1.2, w_hand=1.35,
    chest=3.3, waist=2.4, skirt=7.0,
    palette={
        "k": "1E1A2A",
        "s": "F4D2B8", "S": "D0A48A",
        "h": "2A1E22",                          # dark brown hair
        "b": "EEF2F6", "B": "BAC8D8",          # white robe / shade
        "w": "6AAEE0",                          # sky blue trim
        "d": "4A8AC8",                          # sash
        "p": "C8D0DC", "P": "A0AABA",          # light trousers
        "n": "2E4A78",                          # dark blue shoes
        "r": "E86A8A",                          # pink ribbons
    },
    look=dict(arm_front="b", arm_back="B", arm_front2="b", arm_back2="B",
              leg_front="p", leg_back="P", boot="n",
              torso="b", torso_dark="B", skirt="b", skirt_dark="B",
              belt="d", trim="w", cuff="w"),
    head="seolhwa",
)

KICK_GUARD_F = (55, 140)
SEOLHWA_OVERRIDES = {
    "idle_0": P(lean=4, lf=(18, 4), lb=(-26, -10), af=KICK_GUARD_F, ab=(25, 160), open_hand=True),
    "idle_1": P(lean=5, lf=(21, 7), lb=(-29, -13), af=(53, 136), ab=(23, 156), open_hand=True),
    "walk_0": P(lean=4, lf=(25, 5), lb=(-20, -25), af=KICK_GUARD_F, ab=(25, 160), open_hand=True),
    "walk_1": P(lean=4, lf=(8, 0), lb=(-5, -40), af=KICK_GUARD_F, ab=(25, 160), open_hand=True),
    "walk_2": P(lean=4, lf=(-18, -25), lb=(25, 5), af=KICK_GUARD_F, ab=(25, 160), open_hand=True),
    "walk_3": P(lean=4, lf=(-5, -40), lb=(8, 0), af=KICK_GUARD_F, ab=(25, 160), open_hand=True),
    "lk_1": P(lean=-12, lf=(84, 88), lb=(-6, 0), af=(60, 150), ab=(-20, 40)),
    "hk_0": P(lean=-20, lf=(130, 20), lb=(-4, 0), af=(40, 160), ab=(-30, 30)),
    "hk_1": P(lean=-34, lf=(132, 136), lb=(-4, 0), af=(30, 165), ab=(-40, 10)),       # head-high kick
    "jk": P(lean=-18, lf=(80, 82), lb=(110, 20), af=(40, 160), ground="air"),         # flying kick
    "dp_0": P(lean=24, lf=CROUCH_F, lb=CROUCH_B, af=(30, 70), ab=(-10, 60)),
    "dp_1": P(lean=-28, lf=(158, 165), lb=(40, -20), af=(-20, 40), ab=(60, 140), ground="air"),   # rising spin kick
    "rush_0": P(lean=-6, lf=(70, 10), lb=(-15, -5), af=(50, 150), ab=(-20, 40)),
    "rush_1": P(lean=-24, lf=(94, 92), lb=(-12, -2), af=(40, 165), ab=(-30, 20)),     # driving side kick
}

SWORD_GUARD = (45, 115)
MUYEONG_OVERRIDES = {
    "idle_0": P(lean=6, af=SWORD_GUARD, wa=125),
    "idle_1": P(lf=(29, 9), lb=(-29, -14), af=(43, 112), ab=(18, 145), wa=122),
    "walk_0": P(lf=(25, 5), lb=(-20, -25), af=SWORD_GUARD, wa=125),
    "walk_1": P(lf=(8, 0), lb=(-5, -40), af=SWORD_GUARD, wa=125),
    "walk_2": P(lf=(-18, -25), lb=(25, 5), af=SWORD_GUARD, wa=125),
    "walk_3": P(lf=(-5, -40), lb=(8, 0), af=SWORD_GUARD, wa=125),
    "crouch": P(lean=20, lf=CROUCH_F, lb=CROUCH_B, af=(50, 100), wa=110),
    "jump": P(lean=10, lf=TUCK_F, lb=TUCK_B, af=(40, 120), ground="air", wa=135),
    "lp_0": P(lean=8, af=(40, 80), wa=95),                                   # thrust
    "lp_1": P(lean=18, lf=(32, 2), lb=(-32, -18), af=(92, 90), wa=90),
    "hp_0": P(lean=-12, af=(-50, 170), ab=(30, 140), wa=205),               # big horizontal slash
    "hp_1": P(lean=24, lf=(40, 0), lb=(-40, -40), af=(100, 100), ab=(10, 120), wa=100),
    "clp": P(lean=20, lf=CROUCH_F, lb=CROUCH_B, af=(88, 86), wa=88),
    "chp_0": P(lean=24, lf=CROUCH_F, lb=CROUCH_B, af=(-20, 30), wa=-30),
    "chp_1": P(lean=2, lf=(45, -5), lb=(10, -50), af=(160, 176), wa=178),   # rising slash
    "chk": P(lean=40, lf=(70, 78), lb=(88, -88), af=(70, 75), ab=(35, 25), ground="hip_low", wa=82),  # low slash
    "jp": P(lean=12, lf=TUCK_F, lb=TUCK_B, af=(60, 55), ground="air", wa=45),
    "guard": P(lean=-6, lf=(20, 5), lb=(-30, -15), af=(40, 150), ab=(45, 150), wa=172),
    "cguard": P(lean=12, lf=CROUCH_F, lb=CROUCH_B, af=(45, 150), ab=(55, 150), wa=172),
    "sp_0": P(lean=-8, af=(-40, 160), wa=215),                               # sword wave wind-up
    "sp_1": P(lean=22, lf=(40, 0), lb=(-40, -40), af=(95, 95), wa=95),
    "rush_0": P(lean=-10, lf=(30, 5), lb=(-35, -20), af=(10, 40), wa=-25),   # dash slash
    "rush_1": P(lean=34, lf=(60, 5), lb=(-55, -45), af=(95, 88), wa=92),
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
    if spec["head"] == "seolhwa":
        hair = hm & ((hu > 0.7) | (hf < -0.9) | ((hu > -0.2) & (hf < -0.2)))
        cv.c[hair] = "h"
        bun = circle_mask((cx - fx * 1.2 + ux * (r + 0.7), cy - fy * 1.2 + uy * (r + 0.7)), 1.7)
        cv.put(bun, "h")
        for k, (lx, ly) in enumerate(((4.2, 1.8), (3.6, 3.6))):          # two pink ribbons flowing back
            a0 = (cx - fx * 2.4 + ux * (r + 0.6), cy - fy * 2.4 + uy * (r + 0.6))
            a1 = (a0[0] - fx * lx - ux * ly, a0[1] - fy * lx - uy * ly)
            cv.put(seg_mask(a0, a1, 0.55), "r")
        trim = seg_mask(pt["neck"], (pt["neck"][0] + fx * 1.5 - ux * 2.5, pt["neck"][1] + fy * 1.5 - uy * 2.5), 0.7) & tm
        cv.c[trim] = "w"
    elif spec["head"] == "muyeong":
        hair = hm & ((hu > 0.9) | (hf < -0.8))
        cv.c[hair] = "h"
        tie = circle_mask((cx - fx * (r - 0.2) + ux * 1.6, cy - fy * (r - 0.2) + uy * 1.6), 0.8)
        cv.put(tie & ~hm, "w")
        tail_a = (cx - fx * (r + 0.4) + ux * 1.4, cy - fy * (r + 0.4) + uy * 1.4)
        tail_b = (tail_a[0] - fx * 3.5 - ux * 5.5, tail_a[1] - fy * 3.5 - uy * 5.5)
        cv.put(seg_mask(tail_a, tail_b, 0.9) & ~hm, "h")
        trim = seg_mask(pt["neck"], (pt["neck"][0] + fx * 1.5 - ux * 2.5, pt["neck"][1] + fy * 1.5 - uy * 2.5), 0.7) & tm
        cv.c[trim] = "w"
    elif spec["head"] == "danwoo":
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

    # weapon held in the front hand
    wp = spec.get("weapon")
    if wp:
        wa = pose["wa"] if pose["wa"] is not None else pose["af"][1]
        d = dirv(wa)
        d = (d[0] * math.cos(rot) - d[1] * math.sin(rot), d[0] * math.sin(rot) + d[1] * math.cos(rot))
        hx_, hy_ = pt["hF"]
        nx_, ny_ = -d[1], d[0]
        base = (hx_ + d[0] * 1.2, hy_ + d[1] * 1.2)
        tip = (hx_ + d[0] * wp["length"], hy_ + d[1] * wp["length"])
        cv.put(seg_mask(base, tip, 0.62), "W")
        cv.put(seg_mask((tip[0] - d[0] * 2.5, tip[1] - d[1] * 2.5), tip, 0.62), "Q")      # tip glint
        cv.put(seg_mask((base[0] + nx_ * 1.8, base[1] + ny_ * 1.8), (base[0] - nx_ * 1.8, base[1] - ny_ * 1.8), 0.6), "y")  # guard
        pommel = (hx_ - d[0] * 2.2, hy_ - d[1] * 2.2)
        cv.put(seg_mask((hx_, hy_), pommel, 0.7), "m")                                     # hilt
        cv.put(seg_mask(pommel, (pommel[0] - d[0] * 0.6, pommel[1] + 2.2), 0.5), "r")      # tassel
        pt["tip"] = tip

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


SCALE_KEYS = ("thigh", "shin", "torso", "upper", "fore", "head_r", "w_thigh", "w_shin", "w_upper",
              "w_fore", "w_hand", "chest", "waist", "skirt")
SPRITE_SCALE = float(os.environ.get("SPRITE_SCALE", "1"))   # experiment: uniform rig scale
OUT_DIR = os.environ.get("OUT_DIR")                          # experiment: write data/previews elsewhere


def generate(spec, overrides):
    if SPRITE_SCALE != 1:
        spec = dict(spec)
        for k in SCALE_KEYS:
            spec[k] = spec[k] * SPRITE_SCALE
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
            extra = f" tip={rel('tip')}" if "tip" in pt else ""
            points.append(f"{name}: hF={rel('hF')} hB={rel('hB')} fF={rel('fF')} fB={rel('fB')} head={rel('head')}{extra}")
        g, ax, ay = crop(grid)
        h, w = g.shape
        out.append(f"SPRITE {name} {w} {h} {ax} {ay}")
        for row in g:
            out.append("".join(c if c else "." for c in row))
        previews.append((name, g))
    for name, tpf, frames in ANIMS:
        out.append(f"ANIM {name} {tpf} " + " ".join(frames))
    path = os.path.join(OUT_DIR or os.path.join(ROOT, "data"), f"{spec['name']}.txt")
    with open(path, "w", newline="\r\n") as f:
        f.write("\n".join(out) + "\n")

    with open(os.path.join(OUT_DIR or os.path.join(ROOT, "docs"), f"{spec['name']}_points.txt"), "w") as f:
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
    sheet.save(os.path.join(OUT_DIR or os.path.join(ROOT, "docs", "preview"), f"{spec['name']}_sheet.png"))
    return [n for n, _ in previews]


if __name__ == "__main__":
    only = os.environ.get("ONLY")
    for spec, ov in ((DANWOO, {}), (CHEOLSAN, CHEOLSAN_OVERRIDES), (MUYEONG, MUYEONG_OVERRIDES),
                     (SEOLHWA, SEOLHWA_OVERRIDES)):
        if only and spec["name"] != only:
            continue
        names = generate(spec, ov)
        print(spec["name"], len(names), "sprites")
