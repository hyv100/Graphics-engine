#!/usr/bin/env python3
"""
SereneHealth Graphics Engine (v2)
Bulk-generates square healthcare marketing graphics from a CSV or XLSX content calendar.

v2 enhancements
---------------
- 2x supersampled rendering: every curve, pill, chip and icon is anti-aliased.
- Geometry-aware text fitting: the navy panel curve is a function, so headline,
  subheadline and body copy can never spill past the curve, whatever you feed in.
  A balanced-wrap solver picks the line breaks that fit best before shrinking type.
- Redrawn footer icons (globe / phone / mail) plus new card icons:
  map, clipboard, chat, star, clock, doc, pin, shield (with friendly aliases).
- Auto-fit everywhere text can overflow: day pill, right slogan, handwritten
  callout, floating badge, outcome strip, outcome script and footer contacts.
- PNG output (just name the file *.png, or pass --format png), --format override,
  --only <substring> to render a subset, --validate to QA a calendar without
  rendering, per-row error isolation with an end-of-run summary.

Usage
-----
  python serene_engine.py --calendar content_calendar.csv --output output
  python serene_engine.py --calendar content_calendar.xlsx --output output --format png
  python serene_engine.py --calendar content_calendar.csv --validate
  python serene_engine.py --calendar content_calendar.csv --only workflow --output output

Optional:
  --font-dir fonts            folder containing custom fonts
  --size 1254                 output width/height
"""
import argparse, csv, math, os, re, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageEnhance

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

SS = 2  # internal supersampling factor (rendered at size*SS, downsampled with LANCZOS)

# ---------------------------------------------------------------- palette ----
NAVY = "#090A34"
NAVY2 = "#11144A"
YELLOW = "#FFB82E"
CREAM = "#FFF4D8"
WHITE = "#FFFFFF"
CARD = "#FBFBFE"
LAVENDER = "#F3F2FF"
MID = "#6D6F8E"
LINE = "#D9DBEA"

# ------------------------------------------------------------- typography ----
FONT_CANDIDATES = {
    "bold":     ["Poppins-ExtraBold.ttf", "Poppins-Bold.ttf", "Montserrat-ExtraBold.ttf",
                 "Montserrat-Bold.ttf", "Lato-Heavy.ttf", "Lato-Bold.ttf", "DejaVuSans-Bold.ttf"],
    "semibold": ["Poppins-SemiBold.ttf", "Montserrat-SemiBold.ttf", "Lato-Semibold.ttf",
                 "DejaVuSans-Bold.ttf"],
    "regular":  ["Poppins-Regular.ttf", "Montserrat-Regular.ttf", "Lato-Regular.ttf", "DejaVuSans.ttf"],
    "script":   ["BrushScript.ttf", "DancingScript-Bold.ttf", "DancingScript-SemiBold.ttf",
                 "Caveat-Bold.ttf", "Pacifico-Regular.ttf", "URWChanceryL-MediumItalic.otf",
                 "DejaVuSerif-Italic.ttf"],
}
_FONT_PATH_CACHE = {}
_FONT_OBJ_CACHE = {}

def find_font(kind, font_dir=None):
    key = (kind, str(font_dir))
    if key in _FONT_PATH_CACHE:
        return _FONT_PATH_CACHE[key]
    roots = []
    if font_dir:
        roots.append(Path(font_dir))
    roots += [Path("/usr/share/fonts/truetype/lato"), Path("/usr/share/fonts/truetype/dejavu"),
              Path("/usr/share/fonts/opentype/urw-base35")]
    found = None
    for root in roots:
        if not root.exists():
            continue
        for name in FONT_CANDIDATES[kind]:
            p = root / name
            if p.exists():
                found = str(p); break
        if found:
            break
        pats = {"bold": ["*Bold*.ttf", "*Heavy*.ttf"], "semibold": ["*Semibold*.ttf", "*Demi*.ttf"],
                "regular": ["*Regular*.ttf"], "script": ["*Chancery*.otf", "*Italic*.ttf"]}
        for pat in pats[kind]:
            hits = list(root.glob(pat))
            if hits:
                found = str(hits[0]); break
        if found:
            break
    _FONT_PATH_CACHE[key] = found
    return found

def F(px, kind="regular", font_dir=None):
    """Load (and cache) a font at a pixel size."""
    px = max(6, int(round(px)))
    key = (px, kind, str(font_dir))
    if key not in _FONT_OBJ_CACHE:
        p = find_font(kind, font_dir)
        _FONT_OBJ_CACHE[key] = ImageFont.truetype(p, px) if p else ImageFont.load_default()
    return _FONT_OBJ_CACHE[key]

# ------------------------------------------------------- navy panel curve ----
# Outer edge of the navy panel in 1254-design-space coordinates. The panel
# interior is LEFT of this edge; x shrinks as y grows.
PANEL_PTS = [(0, 0), (795, 0), (770, 105), (735, 230), (695, 350),
             (645, 430), (560, 480), (390, 500), (0, 500)]

def panel_x_over_band(y0, y1):
    """Minimum panel-edge x over the horizontal design-space band [y0, y1]."""
    y0 = max(0.0, min(500.0, float(y0)))
    y1 = max(0.0, min(500.0, float(y1)))
    if y1 < y0:
        y0, y1 = y1, y0
    best = 1e9
    for (xa, ya), (xb, yb) in zip(PANEL_PTS, PANEL_PTS[1:]):
        if ya == yb:
            continue
        lo, hi = (ya, yb) if ya < yb else (yb, ya)
        c0, c1 = max(lo, y0), min(hi, y1)
        if c0 > c1:
            continue
        xat = lambda y: xa + (xb - xa) * (y - ya) / (yb - ya)
        best = min(best, xat(c0), xat(c1))
    for x, y in PANEL_PTS:
        if y0 <= y <= y1:
            best = min(best, x)
    return best

# ------------------------------------------------------------- icon system ----
ICON_ALIASES = {
    "cal": "calendar", "schedule": "calendar", "date": "calendar", "booking": "calendar",
    "reminder": "bell", "reminders": "bell", "alarm": "bell", "notification": "bell",
    "team": "people", "person": "people", "group": "people", "users": "people", "patient": "people",
    "tick": "check", "confirm": "check", "confirmed": "check", "done": "check", "yes": "check",
    "love": "heart", "favorite": "heart", "care": "heart",
    "chart": "growth", "stats": "growth", "graph": "growth", "increase": "growth",
    "workflow": "map", "route": "map", "journey": "map", "roadmap": "map",
    "task": "clipboard", "tasks": "clipboard", "todo": "clipboard", "checklist": "clipboard",
    "message": "chat", "sms": "chat", "text": "chat", "support": "chat",
    "rating": "star", "quality": "star",
    "time": "clock", "hours": "clock",
    "document": "doc", "file": "doc", "notes": "doc", "chartnote": "doc",
    "location": "pin",
    "security": "shield", "compliance": "shield", "hipaa": "shield",
}
KNOWN_ICONS = {"calendar", "bell", "people", "heart", "growth", "check", "map", "clipboard",
               "chat", "star", "clock", "doc", "pin", "shield", "globe", "phone", "mail"}

def resolve_icon(name):
    n = str(name or "").strip().lower()
    if n in KNOWN_ICONS:
        return n, None
    if n in ICON_ALIASES:
        return ICON_ALIASES[n], f"icon '{name}' -> '{ICON_ALIASES[n]}'"
    return "check", f"unknown icon '{name}' (used 'check')"

def draw_icon(d, cx, cy, kind, s=1.0, color=NAVY):
    """Draw an icon centered at (cx, cy). `s` scales design units (1.0 = card size at 1254px)."""
    kind, _ = resolve_icon(kind)
    W = lambda v: max(2, int(round(v * s)))
    if kind == "calendar":
        d.rounded_rectangle((cx-25*s, cy-20*s, cx+25*s, cy+28*s), radius=5*s, outline=color, width=W(4.6))
        d.line((cx-25*s, cy-5*s, cx+25*s, cy-5*s), fill=color, width=W(4))
        for dx in (-13, 0, 13):
            for dy in (4, 16):
                d.ellipse((cx+dx*s-3*s, cy+dy*s-3*s, cx+dx*s+3*s, cy+dy*s+3*s), fill=color)
        d.line((cx-15*s, cy-27*s, cx-15*s, cy-17*s), fill=color, width=W(4.6))
        d.line((cx+15*s, cy-27*s, cx+15*s, cy-17*s), fill=color, width=W(4.6))
    elif kind == "bell":
        d.arc((cx-22*s, cy-25*s, cx+22*s, cy+20*s), 180, 360, fill=color, width=W(4.6))
        d.line((cx-22*s, cy-2*s, cx-27*s, cy+20*s), fill=color, width=W(4.6))
        d.line((cx+22*s, cy-2*s, cx+27*s, cy+20*s), fill=color, width=W(4.6))
        d.line((cx-27*s, cy+20*s, cx+27*s, cy+20*s), fill=color, width=W(4.6))
        d.arc((cx-7*s, cy+17*s, cx+7*s, cy+29*s), 0, 180, fill=color, width=W(4))
        d.ellipse((cx-4*s, cy-31*s, cx+4*s, cy-23*s), fill=color)
    elif kind == "people":
        d.ellipse((cx-10*s, cy-26*s, cx+10*s, cy-6*s), fill=color)
        d.ellipse((cx-31*s, cy-20*s, cx-15*s, cy-4*s), fill=color)
        d.ellipse((cx+15*s, cy-20*s, cx+31*s, cy-4*s), fill=color)
        d.pieslice((cx-28*s, cy-2*s, cx+28*s, cy+38*s), 180, 360, fill=color)
        d.pieslice((cx-43*s, cy-1*s, cx-6*s, cy+31*s), 180, 360, fill=color)
        d.pieslice((cx+6*s, cy-1*s, cx+43*s, cy+31*s), 180, 360, fill=color)
    elif kind == "heart":
        pts = []
        for t in range(0, 361, 5):
            a = math.radians(t)
            x = 16*s*math.sin(a)**3
            y = -(13*s*math.cos(a) - 5*s*math.cos(2*a) - 2*s*math.cos(3*a) - math.cos(4*a))
            pts.append((cx+x, cy+y))
        d.line(pts + [pts[0]], fill=color, width=W(4), joint="curve")
    elif kind == "growth":
        d.rectangle((cx-25*s, cy+5*s, cx-8*s, cy+25*s), fill=color)
        d.rectangle((cx-3*s, cy-8*s, cx+14*s, cy+25*s), fill=color)
        d.rectangle((cx+19*s, cy-25*s, cx+36*s, cy+25*s), fill=color)
        d.line((cx-27*s, cy-2*s, cx+28*s, cy-32*s), fill=YELLOW, width=W(4.6))
        d.polygon([(cx+20*s, cy-30*s), (cx+31*s, cy-32*s), (cx+29*s, cy-21*s)], fill=YELLOW)
    elif kind == "check":
        d.ellipse((cx-24*s, cy-24*s, cx+24*s, cy+24*s), outline=color, width=W(4))
        d.line((cx-12*s, cy, cx-2*s, cy+11*s), fill=color, width=W(4.6))
        d.line((cx-2*s, cy+11*s, cx+15*s, cy-10*s), fill=color, width=W(4.6))
    elif kind == "map":
        d.rounded_rectangle((cx-26*s, cy-20*s, cx+26*s, cy+20*s), radius=4*s, outline=color, width=W(4))
        d.line((cx-9*s, cy-20*s, cx-9*s, cy+20*s), fill=color, width=W(2.6))
        d.line((cx+9*s, cy-20*s, cx+9*s, cy+20*s), fill=color, width=W(2.6))
        d.line((cx-18*s, cy+10*s, cx-2*s, cy-2*s), fill=color, width=W(3))
        d.ellipse((cx-21*s, cy+7*s, cx-13*s, cy+15*s), fill=YELLOW)
        d.ellipse((cx+12*s, cy-12*s, cx+20*s, cy-4*s), fill=YELLOW)
    elif kind == "clipboard":
        d.rounded_rectangle((cx-20*s, cy-22*s, cx+20*s, cy+24*s), radius=4*s, outline=color, width=W(4))
        d.rounded_rectangle((cx-9*s, cy-27*s, cx+9*s, cy-16*s), radius=3*s, fill=color)
        d.line((cx-11*s, cy-4*s, cx+11*s, cy-4*s), fill=color, width=W(3))
        d.line((cx-11*s, cy+5*s, cx+11*s, cy+5*s), fill=color, width=W(3))
        d.line((cx-11*s, cy+14*s, cx+3*s, cy+14*s), fill=color, width=W(3))
    elif kind == "chat":
        d.rounded_rectangle((cx-24*s, cy-20*s, cx+24*s, cy+12*s), radius=7*s, outline=color, width=W(4))
        d.polygon([(cx-10*s, cy+12*s), (cx-2*s, cy+12*s), (cx-12*s, cy+24*s)], fill=color)
        for dx in (-11, 0, 11):
            d.ellipse((cx+dx*s-2.6*s, cy-5*s-2.6*s, cx+dx*s+2.6*s, cy-5*s+2.6*s), fill=color)
    elif kind == "star":
        pts = []
        for i in range(10):
            a = math.radians(-90 + i*36)
            r = (24*s) if i % 2 == 0 else (10.5*s)
            pts.append((cx + r*math.cos(a), cy + r*math.sin(a)))
        d.polygon(pts, fill=color)
    elif kind == "clock":
        d.ellipse((cx-23*s, cy-23*s, cx+23*s, cy+23*s), outline=color, width=W(4))
        d.line((cx, cy, cx, cy-13*s), fill=color, width=W(3.4))
        d.line((cx, cy, cx+9*s, cy+5*s), fill=color, width=W(3.4))
    elif kind == "doc":
        d.polygon([(cx-17*s, cy-24*s), (cx+8*s, cy-24*s), (cx+17*s, cy-15*s),
                   (cx+17*s, cy+24*s), (cx-17*s, cy+24*s)], outline=color, width=W(4))
        d.line((cx+8*s, cy-24*s, cx+8*s, cy-15*s), fill=color, width=W(2.6))
        d.line((cx+8*s, cy-15*s, cx+17*s, cy-15*s), fill=color, width=W(2.6))
        for dy in (-4, 5, 14):
            d.line((cx-9*s, cy+dy*s, cx+9*s, cy+dy*s), fill=color, width=W(2.6))
    elif kind == "pin":
        d.pieslice((cx-16*s, cy-26*s, cx+16*s, cy+6*s), 140, 40, fill=color)
        d.polygon([(cx-11*s, cy+2*s), (cx+11*s, cy+2*s), (cx, cy+26*s)], fill=color)
        d.ellipse((cx-6*s, cy-16*s, cx+6*s, cy-4*s), fill=CREAM)
    elif kind == "shield":
        d.polygon([(cx, cy-24*s), (cx+20*s, cy-16*s), (cx+18*s, cy+6*s),
                   (cx, cy+25*s), (cx-18*s, cy+6*s), (cx-20*s, cy-16*s)],
                  outline=color, width=W(4))
        d.line((cx-8*s, cy-1*s, cx-2*s, cy+7*s), fill=color, width=W(3.6))
        d.line((cx-2*s, cy+7*s, cx+10*s, cy-9*s), fill=color, width=W(3.6))
    elif kind == "globe":
        d.ellipse((cx-16*s, cy-16*s, cx+16*s, cy+16*s), outline=color, width=W(3))
        d.ellipse((cx-7*s, cy-16*s, cx+7*s, cy+16*s), outline=color, width=W(2.4))
        d.line((cx-15*s, cy, cx+15*s, cy), fill=color, width=W(2.4))
    elif kind == "phone":
        d.rounded_rectangle((cx-11*s, cy-19*s, cx+11*s, cy+19*s), radius=5*s, outline=color, width=W(3.2))
        d.line((cx-4*s, cy+11*s, cx+4*s, cy+11*s), fill=color, width=W(2.4))
    elif kind == "mail":
        d.rounded_rectangle((cx-21*s, cy-14*s, cx+21*s, cy+14*s), radius=3*s, outline=color, width=W(3.2))
        d.line((cx-20*s, cy-12.5*s, cx, cy+3*s), fill=color, width=W(2.8))
        d.line((cx, cy+3*s, cx+20*s, cy-12.5*s), fill=color, width=W(2.8))

# ------------------------------------------------------------ render context ----
class Ctx:
    """Per-render state: canvas, scale factor, font dir, warnings, fit report."""
    def __init__(self, size, font_dir):
        self.size = size
        self.S = size * SS / 1254.0          # pixels per design unit
        self.font_dir = font_dir
        self.canvas = Image.new("RGB", (size * SS, size * SS), WHITE)
        self.d = ImageDraw.Draw(self.canvas)
        self.warnings = []
        self.report = {}

    def sc(self, n):
        return int(n * self.S)

    def warn(self, msg):
        self.warnings.append(msg)

# --------------------------------------------------------------- measuring ----
def ink(d, text, f):
    """(width, height, bbox) of text drawn with default 'la' anchor."""
    bb = d.textbbox((0, 0), text, font=f)
    return bb[2] - bb[0], bb[3] - bb[1], bb

def nav_limit_px(c, x_des, y0_px, y1_px, margin_des=12.0):
    """Max text width in px for text starting at design-x x_des so it stays
    inside the navy panel across the pixel band [y0_px, y1_px]."""
    edge = panel_x_over_band(y0_px / c.S, y1_px / c.S)
    return max(80.0, (edge - x_des - margin_des) * c.S)

def best_wrap(d, text, f, limits, max_lines):
    """DP wrap: choose line breaks minimizing the max width/limit ratio so the
    fewest lines that fit win. Returns (lines, ratio) or (None, ratio>1)."""
    words = str(text).split()
    n = len(words)
    if not words:
        return [], 0.999
    widths = {}

    def w(i, j):
        if (i, j) not in widths:
            widths[(i, j)] = d.textbbox((0, 0), " ".join(words[i:j]), font=f)[2]
        return widths[(i, j)]

    lim = lambda k: limits[min(k, len(limits) - 1)]
    INF = float("inf")
    dp = [[INF] * (n + 1) for _ in range(max_lines + 1)]
    bk = [[-1] * (n + 1) for _ in range(max_lines + 1)]
    dp[0][0] = 0.0
    for k in range(1, max_lines + 1):
        for i in range(k, n + 1):
            for j in range(k - 1, i):
                if dp[k-1][j] == INF or j == i:
                    continue
                ratio = w(j, i) / lim(k-1)
                val = max(dp[k-1][j], ratio)
                if val < dp[k][i]:
                    dp[k][i] = val
                    bk[k][i] = j
    for k in range(1, max_lines + 1):
        if dp[k][n] <= 1.0:
            lines, i = [], n
            for kk in range(k, 0, -1):
                j = bk[kk][i]
                lines.append(" ".join(words[j:i]))
                i = j
            lines.reverse()
            return lines, dp[k][n]
    return None, dp[max_lines][n]

# ------------------------------------------------------------- fit helpers ----
def fit_px(c, text, start, mins, kind, max_w_px):
    """Largest font size (descending by 1px design units) whose single-line ink
    fits max_w_px. Returns (font, size, fits)."""
    for size in range(start, mins - 1, -1):
        f = F(c.sc(size), kind, c.font_dir)
        w, _, _ = ink(c.d, text, f)
        if w <= max_w_px:
            return f, size, True
    f = F(c.sc(mins), kind, c.font_dir)
    return f, mins, ink(c.d, text, f)[0] <= max_w_px

def fit_block_center(c, lines, start, mins, kind, max_w_px):
    """One uniform size for every line in a centered block (slogan/handwritten/...)."""
    lines = [l for l in lines if l != ""] or [""]
    for size in range(start, mins - 1, -1):
        f = F(c.sc(size), kind, c.font_dir)
        if all(ink(c.d, l, f)[0] <= max_w_px for l in lines):
            return f, size, True
    f = F(c.sc(mins), kind, c.font_dir)
    ok = all(ink(c.d, l, f)[0] <= max_w_px for l in lines)
    return f, mins, ok

def truncate_to_width(d, text, font, max_w_px):
    """Ellipsize text so its ink never exceeds max_w_px."""
    if d.textbbox((0, 0), text, font=font)[2] <= max_w_px:
        return text
    out = text
    while out:
        cand = (out.rsplit(" ", 1)[0] if " " in out else out[:-1])
        if cand and d.textbbox((0, 0), cand.rstrip() + "…", font=font)[2] <= max_w_px:
            return cand.rstrip() + "…"
        if not cand:
            break
        out = cand
    return "…"

def fit_nav_line(c, text, x_des, y_des, start, mins, kind, margin_des=12.0):
    """Fit a single line of navy-panel text against the curve. Never overflows:
    if it can't fit at min size the text is ellipsized. Returns (font, size, text, fits)."""
    for size in range(start, mins - 1, -1):
        f = F(c.sc(size), kind, c.font_dir)
        w, h, bb = ink(c.d, text, f)
        y0, y1 = c.sc(y_des) + bb[1], c.sc(y_des) + bb[3]
        if w <= nav_limit_px(c, x_des, y0, y1, margin_des):
            return f, size, text, True
    f = F(c.sc(mins), kind, c.font_dir)
    w, h, bb = ink(c.d, text, f)
    limit = nav_limit_px(c, x_des, c.sc(y_des) + bb[1], c.sc(y_des) + bb[3], margin_des)
    if w <= limit:
        return f, mins, text, True
    return f, mins, truncate_to_width(c.d, text, f, limit), False

def fit_nav_paragraph(c, text, x_des, y_des, max_lines, start, mins, kind, gap_des=7.0, margin_des=12.0):
    """Fit a wrapped paragraph inside the navy panel using balanced wrapping.
    Returns (font, size, lines, fits)."""
    for size in range(start, mins - 1, -1):
        f = F(c.sc(size), kind, c.font_dir)
        # provisional per-line limits from a uniform (conservative) line height
        _, h_prov, _ = ink(c.d, text, f)
        line0_top = c.sc(y_des)
        pitch = h_prov + c.sc(gap_des)
        limits = [nav_limit_px(c, x_des, line0_top + i * pitch,
                               line0_top + i * pitch + h_prov, margin_des)
                  for i in range(max_lines)]
        lines, ratio = best_wrap(c.d, text, f, limits, max_lines)
        if lines is None:
            continue
        # verify with true per-line ink bands
        ok, y = True, c.sc(y_des)
        for ln in lines:
            w, h, bb = ink(c.d, ln, f)
            if w > nav_limit_px(c, x_des, y + bb[1], y + bb[3], margin_des):
                ok = False
                break
            y += h + c.sc(gap_des)
        if ok:
            return f, size, lines, True
    f = F(c.sc(mins), kind, c.font_dir)
    # last resort at min size: wrap against the true per-line curve limits,
    # ellipsizing the final line — text can never spill past the curve.
    _, h_prov, _ = ink(c.d, text, f)
    pitch = h_prov + c.sc(gap_des)
    limits = [nav_limit_px(c, x_des, c.sc(y_des) + i * pitch,
                           c.sc(y_des) + i * pitch + h_prov, margin_des)
              for i in range(max_lines)]
    lines_, cur = [], ""
    for wd in str(text).split():
        t = (cur + " " + wd).strip()
        if c.d.textbbox((0, 0), t, font=f)[2] <= limits[min(len(lines_), max_lines - 1)]:
            cur = t
        else:
            if cur:
                lines_.append(cur)
                if len(lines_) == max_lines:
                    break
            cur = wd
    if cur and len(lines_) < max_lines:
        lines_.append(cur)
    if lines_:
        li = min(len(lines_) - 1, max_lines - 1)
        dropped = bool(cur) and len(lines_) == max_lines
        last = lines_[li]
        if dropped and not last.endswith("…"):
            last = last + "…"
        lines_[li] = truncate_to_width(c.d, last, f, limits[li])
    return f, mins, lines_, False

def fit_block_ma(d, xy, lines, font, fill, spacing, anchor="ma"):
    y = xy[1]
    for ln in lines:
        d.text((xy[0], y), ln, font=font, fill=fill, anchor=anchor)
        _, h, _ = ink(d, ln, font)
        y += h + spacing

# ----------------------------------------------------------------- loading ----
def load_rows(path):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))
    if path.suffix.lower() == ".xlsx":
        if load_workbook is None:
            raise RuntimeError("Install openpyxl to read XLSX files.")
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        vals = list(ws.values)
        if not vals:
            return []
        headers = [str(x or "").strip() for x in vals[0]]
        return [dict(zip(headers, row)) for row in vals[1:] if any(v is not None for v in row)]
    raise ValueError("Calendar must be .csv or .xlsx")

def val(row, key, default=""):
    v = row.get(key, default)
    if v is None:
        return default
    return str(v).replace("\\n", "\n")

def val_req(row, key, default=""):
    """Like val(), but an empty/blank cell falls back to the default."""
    v = val(row, key, "").strip()
    return v if v else default

def fit_crop(img, box):
    x1, y1, x2, y2 = box
    tw, th = x2 - x1, y2 - y1
    src = img.copy().convert("RGB")
    sw, sh = src.size
    scale = max(tw / sw, th / sh)
    nw, nh = int(sw * scale), int(sh * scale)
    src = src.resize((nw, nh), Image.Resampling.LANCZOS)
    left, top = (nw - tw) // 2, (nh - th) // 2
    return src.crop((left, top, left + tw, top + th))

# ---------------------------------------------------------------- generate ----
def generate(row, assets_dir, out_path, font_dir=None, size=1254, save=True):
    """Render one graphic. Returns (out_path, report_dict)."""
    c = Ctx(size, font_dir)
    d = c.d

    # --- HERO / TOP SECTION ---
    hero_path = val(row, "image").strip()
    if hero_path:
        hp = Path(hero_path)
        if not hp.is_absolute():
            hp = Path(assets_dir) / hero_path
    else:
        hp = Path(assets_dir) / "hero_default.jpg"
    try:
        hero = Image.open(hp).convert("RGB") if hp.exists() else None
        if hero is None:
            c.warn(f"hero image not found: {hp} (used placeholder)")
    except Exception as e:
        hero = None
        c.warn(f"hero image unreadable: {hp} ({e})")
    if hero is None:
        hero = Image.new("RGB", (600, 500), (220, 225, 230))
    hero = fit_crop(hero, (c.sc(760), 0, size * SS, c.sc(505)))
    hero = ImageEnhance.Contrast(hero).enhance(1.02)
    c.canvas.paste(hero, (c.sc(760), 0))

    # white right strip
    d.rectangle((c.sc(1032), 0, size * SS, c.sc(505)), fill="#FFFDF8")
    # navy curved panel (anti-aliased thanks to supersampling)
    pts = [(c.sc(x), c.sc(y)) for x, y in PANEL_PTS]
    d.polygon(pts, fill=NAVY)

    # logo
    logo_path = Path(assets_dir) / "serenehealth_logo.png"
    if logo_path.exists():
        lg = Image.open(logo_path).convert("RGBA")
        ratio = min(c.sc(350) / lg.width, c.sc(85) / lg.height)
        lg = lg.resize((int(lg.width * ratio), int(lg.height * ratio)), Image.Resampling.LANCZOS)
        c.canvas.paste(lg, (c.sc(42), c.sc(18)), lg)
    else:
        c.warn("logo not found (assets/serenehealth_logo.png) — drew a text logotype")
        d.text((c.sc(45), c.sc(38)), "SereneHealth", font=F(c.sc(34), "bold", font_dir), fill=WHITE)

    # sub-brand line
    d.text((c.sc(132), c.sc(77)), val(row, "brand_subtitle", "QUALITY MEDICAL VIRTUAL ASSISTANT SERVICES"),
           font=F(c.sc(15), "semibold", font_dir), fill="#A7A8D5", anchor="lm")

    # day pill (auto-fitting day name)
    day = val_req(row, "day", "WEDNESDAY").upper()
    fday, day_size, ok = fit_px(c, day, 17, 11, "bold", c.sc(136))
    if not ok:
        c.warn(f"day name too long: '{day}'")
    d.rounded_rectangle((c.sc(536), c.sc(28), c.sc(754), c.sc(68)), c.sc(22), fill="#E9E9FF")
    draw_icon(d, c.sc(570), c.sc(48), "calendar", c.S * .38, NAVY)
    d.text((c.sc(600), c.sc(49)), day, font=fday, fill=NAVY, anchor="lm")

    # right slogan (per-block fit; skipped entirely when left blank)
    slogan = [l for l in val(row, "right_slogan", "SUPPORTING\nHEALTHIER\nPRACTICES\nTOGETHER").split("\n")]
    if any(l.strip() for l in slogan):
        fslog, slog_size, ok = fit_block_center(c, slogan, 14, 9, "semibold", c.sc(178))
        if not ok:
            c.warn("right_slogan lines exceed the white strip even at minimum size")
        fit_block_ma(d, (c.sc(1100), c.sc(32)), slogan, fslog, NAVY, c.sc(4))
        d.rounded_rectangle((c.sc(1106), c.sc(126), c.sc(1190), c.sc(132)), c.sc(3), fill=YELLOW)

    # headline (curve-aware, unified size for both lines — ellipsized if truly impossible)
    hw, hy = val_req(row, "headline_white", "A Fuller Schedule."), val_req(row, "headline_yellow", "Fewer Missed Visits.")
    f1, s1, hw_t, ok1 = fit_nav_line(c, hw, 48, 145, 54, 30, "bold")
    f2, s2, hy_t, ok2 = fit_nav_line(c, hy, 48, 213, 54, 30, "bold")
    shared = min(s1, s2)                       # keep both headline lines on one type scale
    if shared != s1 or shared != s2:
        f1 = F(c.sc(shared), "bold", font_dir)
        f2 = f1
        w1, h1, bb1 = ink(c.d, hw, f1)
        w2, h2, bb2 = ink(c.d, hy, f2)
        lim1 = nav_limit_px(c, 48, c.sc(145) + bb1[1], c.sc(145) + bb1[3])
        lim2 = nav_limit_px(c, 48, c.sc(213) + bb2[1], c.sc(213) + bb2[3])
        hw_t = hw if w1 <= lim1 else truncate_to_width(c.d, hw, f1, lim1)
        hy_t = hy if w2 <= lim2 else truncate_to_width(c.d, hy, f2, lim2)
        ok1, ok2 = w1 <= lim1, w2 <= lim2
    if not (ok1 and ok2):
        c.warn(f"headline shrunk/truncated to fit the panel curve (size {shared})")
    d.text((c.sc(48), c.sc(145)), hw_t, font=f1, fill=WHITE)
    d.text((c.sc(48), c.sc(213)), hy_t, font=f2, fill=YELLOW)

    # subheadline
    sub = val_req(row, "subheadline", "Appointment coordination made easier.")
    fsub, ssub, sub_t, oks = fit_nav_line(c, sub, 50, 297, 34, 22, "bold")
    if not oks:
        c.warn("subheadline shrunk/truncated to fit the panel curve")
    d.text((c.sc(50), c.sc(297)), sub_t, font=fsub, fill=WHITE)
    d.rounded_rectangle((c.sc(50), c.sc(357), c.sc(168), c.sc(364)), c.sc(3), fill=YELLOW)

    # body (balanced wrap against the curve)
    body = val_req(row, "body", "SereneHealth helps support your front-end workflow so your day runs smoother.")
    fbody, sbody, blines, okb = fit_nav_paragraph(c, body, 50, 387, 2, 25, 18, "regular")
    if not okb:
        c.warn("body copy truncated to fit the panel (consider shorter copy)")
    y = c.sc(387)
    for ln in blines:
        d.text((c.sc(50), y), ln, font=fbody, fill=WHITE)
        y += ink(d, ln, fbody)[1] + c.sc(7)
    d.rounded_rectangle((c.sc(50), c.sc(357), c.sc(168), c.sc(364)), c.sc(3), fill=YELLOW)

    # handwritten-style message (skipped entirely when left blank)
    hand = [l for l in val(row, "handwritten", "Your\nPractice\nOur Support.").split("\n")]
    if any(l.strip() for l in hand):
        fhand, hsize, okh = fit_block_center(c, hand, 23, 13, "script", c.sc(170))
        if not okh:
            c.warn("handwritten block does not fit the side column at minimum size")
        fit_block_ma(d, (c.sc(1165), c.sc(180)), hand, fhand, NAVY, c.sc(3))
        d.rounded_rectangle((c.sc(1110), c.sc(333), c.sc(1200), c.sc(338)), c.sc(3), fill=YELLOW)

    # floating badge (fit both width and total height; skipped when blank)
    badge_lines = [l for l in val(row, "badge", "More Time\nfor What\nMatters Most.").split("\n")]
    if any(l.strip() for l in badge_lines):
        fbadge = F(c.sc(17), "bold", font_dir)
    for bsize in range(17, 10, -1):
        cand = F(c.sc(bsize), "bold", font_dir)
        total = sum(ink(d, l, cand)[1] + c.sc(1) for l in badge_lines)
        if all(ink(d, l, cand)[0] <= c.sc(142) for l in badge_lines) and total <= c.sc(76):
            fbadge = cand
            break
        if all(ink(d, l, fbadge)[0] > c.sc(142) for l in badge_lines):
            c.warn("badge text too long for the badge card")
        d.rounded_rectangle((c.sc(995), c.sc(380), c.sc(1230), c.sc(487)), c.sc(19), fill=CREAM)
        draw_icon(d, c.sc(1032), c.sc(433), "heart", c.S * .95, YELLOW)
        fit_block_ma(d, (c.sc(1075), c.sc(402)), badge_lines, fbadge, NAVY, c.sc(1), anchor="la")

    # --- SUPPORT CARD ---
    d.rounded_rectangle((c.sc(29), c.sc(515), c.sc(1224), c.sc(893)), c.sc(25), fill=CARD,
                        outline="#E5E6F0", width=c.sc(2))
    title = val_req(row, "support_title", "SereneHealth Supports:")
    ftitle, tsize, okt = fit_px(c, title, 29, 20, "bold", c.sc(480))
    if not okt:
        c.warn("support_title too long even at minimum size")
    d.rounded_rectangle((c.sc(48), c.sc(518), c.sc(574), c.sc(582)), c.sc(20), fill=YELLOW)
    d.text((c.sc(82), c.sc(550)), title, font=ftitle, fill="#05052D", anchor="lm")

    default_items = (["Scheduling", "Reminders", "Confirmations"],
                     ["Book appointments and manage your calendar.",
                      "Send patient reminders to reduce no-shows.",
                      "Follow up and confirm upcoming appointments."],
                     ["calendar", "bell", "people"])
    centers = [c.sc(215), c.sc(626), c.sc(1008)]
    for i in range(1, 4):
        it_title = val_req(row, f"item{i}_title", default_items[0][i-1])
        it_desc = val_req(row, f"item{i}_desc", default_items[1][i-1])
        icon_raw = val(row, f"item{i}_icon", default_items[2][i-1]).strip()
        if icon_raw:
            it_icon, note = resolve_icon(icon_raw)
            if note:
                c.warn(f"item{i}: {note}")
        else:
            it_icon = default_items[2][i-1]
        cx = centers[i-1]
        if i < 3:
            d.line((c.sc(429 + (i-1)*380), c.sc(602), c.sc(429 + (i-1)*380), c.sc(872)), fill=LINE, width=c.sc(1))
        d.ellipse((cx - c.sc(72), c.sc(592), cx + c.sc(72), c.sc(736)), fill=CREAM)
        draw_icon(d, cx, c.sc(664), it_icon, c.S * 1.0, NAVY)
        ft, _, _ = fit_px(c, it_title, 31, 20, "bold", c.sc(270))
        d.text((cx, c.sc(754)), it_title, font=ft, fill=NAVY, anchor="ma")
        # description: shrink until it wraps to <= 3 lines in the column
        fdesc, dlines = F(c.sc(24), "regular", font_dir), None
        for dsize in range(24, 17, -1):
            cand = F(c.sc(dsize), "regular", font_dir)
            ls, cur = [], ""
            for wd in it_desc.split():
                t = (cur + " " + wd).strip()
                if d.textbbox((0, 0), t, font=cand)[2] <= c.sc(260):
                    cur = t
                else:
                    if cur:
                        ls.append(cur)
                    cur = wd
            if cur:
                ls.append(cur)
            if len(ls) <= 3:
                fdesc, dlines = cand, ls
                break
        if dlines is None:
            dlines = ls[:3]
            c.warn(f"item{i}_desc truncated (too long for the card)")
        yy = c.sc(788)
        for ln in dlines:
            d.text((cx, yy), ln, font=fdesc, fill=NAVY, anchor="ma")
            yy += ink(d, ln, fdesc)[1] + c.sc(4)

    # --- OUTCOME STRIP ---
    d.rounded_rectangle((c.sc(25), c.sc(908), c.sc(1228), c.sc(1044)), c.sc(22), fill=LAVENDER)
    draw_icon(d, c.sc(173), c.sc(976), "growth", c.S * .75, NAVY)
    d.line((c.sc(292), c.sc(926), c.sc(292), c.sc(1026)), fill="#BDBFE1", width=c.sc(2))
    outcome_lines = [l for l in val(row, "outcome", "Improve workflow and\nkeep your calendar moving.").split("\n")]
    if any(l.strip() for l in outcome_lines):
        fout, osz, oko = fit_block_center(c, outcome_lines, 32, 20, "bold", c.sc(650))
        if not oko:
            c.warn("outcome strip text too long at minimum size")
        fit_block_ma(d, (c.sc(357), c.sc(944)), outcome_lines, fout, NAVY, c.sc(0), anchor="la")
    hand2_lines = [l for l in val(row, "outcome_script", "Better Care.\nBrighter Days.").split("\n")]
    if any(l.strip() for l in hand2_lines):
        fh2, _, okh2 = fit_block_center(c, hand2_lines, 21, 13, "script", c.sc(150))
        if not okh2:
            c.warn("outcome_script too long at minimum size")
        fit_block_ma(d, (c.sc(1140), c.sc(948)), hand2_lines, fh2, NAVY, c.sc(2))
        d.rounded_rectangle((c.sc(1088), c.sc(1010), c.sc(1193), c.sc(1016)), c.sc(3), fill=YELLOW)

    # --- FOOTER ---
    d.rectangle((0, c.sc(1051), size * SS, size * SS), fill=NAVY)
    d.rectangle((0, c.sc(1049), size * SS, c.sc(1053)), fill=YELLOW)
    d.rounded_rectangle((c.sc(48), c.sc(1090), c.sc(688), c.sc(1184)), c.sc(45), fill=YELLOW)
    cta = val_req(row, "cta", "Let’s support your practice")
    fcta, _, okc = fit_px(c, cta, 29, 16, "bold", c.sc(500))
    if not okc:
        c.warn("CTA text too long for the button")
    d.text((c.sc(100), c.sc(1137)), cta, font=fcta, fill="#05052D", anchor="lm")
    d.text((c.sc(648), c.sc(1137)), "›", font=F(c.sc(44), "regular", font_dir), fill="#05052D", anchor="mm")
    d.line((c.sc(730), c.sc(1080), c.sc(730), c.sc(1220)), fill=YELLOW, width=c.sc(2))

    website = val_req(row, "website", "SereneHealthVA.com")
    phone = val_req(row, "phone", "(630) 394-2838")
    email = val_req(row, "email", "Contact@serenehealthva.com")
    for label, start in ((website, 23), (phone, 23), (email, 21)):
        f, _, okf = fit_px(c, label, start, 14, "bold", c.sc(376))
        if not okf:
            c.warn(f"footer contact too long: '{label}'")
    d.text((c.sc(854), c.sc(1109)), website, font=fit_px(c, website, 23, 14, "bold", c.sc(376))[0],
           fill=WHITE, anchor="lm")
    d.text((c.sc(854), c.sc(1155)), phone, font=fit_px(c, phone, 23, 14, "bold", c.sc(376))[0],
           fill=WHITE, anchor="lm")
    d.text((c.sc(854), c.sc(1201)), email, font=fit_px(c, email, 21, 14, "bold", c.sc(376))[0],
           fill=WHITE, anchor="lm")
    draw_icon(d, c.sc(803), c.sc(1110), "globe", c.S, WHITE)
    draw_icon(d, c.sc(803), c.sc(1156), "phone", c.S, YELLOW)
    draw_icon(d, c.sc(803), c.sc(1202), "mail", c.S, YELLOW)

    # report
    c.report = {
        "headline_white": {"size": s1, "fits": ok1},
        "headline_yellow": {"size": s2, "fits": ok2},
        "subheadline": {"size": ssub, "fits": oks},
        "body": {"size": sbody, "lines": len(blines), "fits": okb},
        "day": {"size": day_size, "fits": ok},
        "support_title": {"size": tsize, "fits": okt},
        "warnings": c.warnings,
    }

    out = c.canvas.resize((size, size), Image.Resampling.LANCZOS)
    if save:
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        if out_path.suffix.lower() in (".jpg", ".jpeg"):
            out.save(out_path, quality=96, subsampling=0, optimize=True)
        elif out_path.suffix.lower() == ".png":
            out.save(out_path, optimize=True)
        else:
            out.save(out_path, quality=96, subsampling=0, optimize=True)
    return out_path, c.report

# --------------------------------------------------------------------- main ----
def slug_for(row, i):
    """Returns (stem, suffix) from the filename column; suffix kept only if it's an image type."""
    slug = val(row, "filename").strip()
    if not slug:
        slug = re.sub(r"[^A-Za-z0-9_-]+", "-", val(row, "date", str(i))).strip("-") or f"graphic-{i:03d}"
    p = Path(slug)
    suffix = p.suffix.lower() if p.suffix.lower() in (".jpg", ".jpeg", ".png") else ""
    return p.stem, suffix

def main():
    ap = argparse.ArgumentParser(description="SereneHealth Graphics Engine v2")
    ap.add_argument("--calendar", required=True)
    ap.add_argument("--output", default="output")
    ap.add_argument("--assets", default="assets")
    ap.add_argument("--font-dir", default="fonts")
    ap.add_argument("--size", type=int, default=1254)
    ap.add_argument("--format", choices=["auto", "jpg", "png"], default="auto",
                    help="auto = follow each row's filename extension (default jpg)")
    ap.add_argument("--only", default="", help="render only rows whose filename/date contains this substring")
    ap.add_argument("--validate", action="store_true", help="QA the calendar (fit report) without saving files")
    args = ap.parse_args()

    rows = load_rows(args.calendar)
    if not rows:
        raise SystemExit("No rows found in calendar.")
    if args.only:
        rows = [r for r in rows if args.only.lower() in (val(r, "filename") + val(r, "date")).lower()]
        if not rows:
            raise SystemExit(f"No rows match --only {args.only!r}")

    ext = ".jpg" if args.format == "jpg" else ".png" if args.format == "png" else None
    os.makedirs(args.output, exist_ok=True)

    # QA guard: duplicate output names would silently overwrite each other
    seen_names, dup_rows = {}, []
    for i, row in enumerate(rows, 1):
        stem, sfx = slug_for(row, i)
        name = stem + (ext or sfx or ".jpg")
        if name in seen_names:
            dup_rows.append((i, name, seen_names[name]))
        else:
            seen_names[name] = i
    if dup_rows:
        print(f"WARNING: {len(dup_rows)} duplicate output filename(s) — rows would overwrite "
              f"each other. Auto-suffixing the later rows.")
        for i, name, first_i in dup_rows:
            print(f"   row {i} would overwrite row {first_i} -> {name}")

    # QA guard: rows sharing one hero photo will look alike in the feed
    img_usage = {}
    for i, row in enumerate(rows, 1):
        img_usage.setdefault(val_req(row, "image", "hero_default.jpg").strip() or "hero_default.jpg", []).append(i)
    for img, i_list in img_usage.items():
        if len(i_list) > 1:
            print(f"NOTE: rows {', '.join(map(str, i_list))} share the same hero image ({img}) — "
                  f"these posts will look very similar in the feed. Vary the 'image' column for variety.")

    ok_count, fail_count, warn_count = 0, 0, 0
    used_names = {}
    for i, row in enumerate(rows, 1):
        stem, sfx = slug_for(row, i)
        name = stem + (ext or sfx or ".jpg")
        if name in used_names:                       # never silently overwrite
            n = 2
            while f"{stem}-{n}{Path(name).suffix}" in used_names:
                n += 1
            name = f"{stem}-{n}{Path(name).suffix}"
        used_names[name] = i
        out = Path(args.output) / name
        try:
            if args.validate:
                _, report = generate(row, args.assets, out, args.font_dir, args.size, save=False)
                warns = report["warnings"]
                status = "OK" if not warns else "WARN"
                warn_count += len(warns)
                fits = {k: v for k, v in report.items() if isinstance(v, dict)}
                shrunk = [k for k, v in fits.items() if v.get("fits") is False]
                print(f"[{status}] {stem}" + (f"  (shrunk: {', '.join(shrunk)})" if shrunk else ""))
                for w in warns:
                    print(f"         - {w}")
                ok_count += 1
            else:
                _, report = generate(row, args.assets, out, args.font_dir, args.size, save=True)
                print("Created", out)
                for w in report["warnings"]:
                    print(f"   warn: {w}")
                    warn_count += 1
                ok_count += 1
        except Exception as e:
            fail_count += 1
            print(f"FAILED {stem}: {e}", file=sys.stderr)

    mode = "validated" if args.validate else "created"
    print(f"\n{ok_count} row(s) {mode}, {fail_count} failed, {warn_count} warning(s).")
    if fail_count:
        sys.exit(1)

if __name__ == "__main__":
    main()
