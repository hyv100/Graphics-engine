
#!/usr/bin/env python3
"""
SereneHealth Graphics Engine
Bulk-generates square healthcare marketing graphics from a CSV or XLSX content calendar.

Usage:
  python serene_engine.py --calendar content_calendar.csv --output output
  python serene_engine.py --calendar content_calendar.xlsx --output output --assets assets

Optional:
  --reference reference.jpg   design reference (for documentation only)
  --font-dir fonts            folder containing custom fonts
  --size 1254                 output width/height
"""
import argparse, csv, os, re, textwrap, math
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

W = H = 1254
NAVY = "#090A34"
NAVY2 = "#11144A"
YELLOW = "#FFB82E"
CREAM = "#FFF4D8"
WHITE = "#FFFFFF"
CARD = "#FBFBFE"
LAVENDER = "#F3F2FF"
MID = "#6D6F8E"
LINE = "#D9DBEA"

FONT_CANDIDATES = {
    "bold": ["Poppins-ExtraBold.ttf","Poppins-Bold.ttf","Montserrat-ExtraBold.ttf","Montserrat-Bold.ttf","Lato-Heavy.ttf","Lato-Bold.ttf","DejaVuSans-Bold.ttf"],
    "semibold": ["Poppins-SemiBold.ttf","Montserrat-SemiBold.ttf","Lato-Semibold.ttf","DejaVuSans-Bold.ttf"],
    "regular": ["Poppins-Regular.ttf","Montserrat-Regular.ttf","Lato-Regular.ttf","DejaVuSans.ttf"],
    "script": ["BrushScript.ttf","URWChanceryL-MediumItalic.otf","DejaVuSerif-Italic.ttf"]
}

def find_font(kind, font_dir=None):
    roots=[]
    if font_dir:
        roots.append(Path(font_dir))
    roots += [Path("/usr/share/fonts/truetype/lato"), Path("/usr/share/fonts/truetype/dejavu"), Path("/usr/share/fonts/opentype/urw-base35")]
    for root in roots:
        if not root.exists(): continue
        for name in FONT_CANDIDATES[kind]:
            p=root/name
            if p.exists(): return str(p)
        # fallback scan
        pats={"bold":["*Bold*.ttf","*Heavy*.ttf"],"semibold":["*Semibold*.ttf","*Demi*.ttf"],"regular":["*Regular*.ttf"],"script":["*Chancery*.otf","*Italic*.ttf"]}
        for pat in pats[kind]:
            hits=list(root.glob(pat))
            if hits: return str(hits[0])
    return None

def F(size, kind="regular", font_dir=None):
    p=find_font(kind,font_dir)
    return ImageFont.truetype(p,size) if p else ImageFont.load_default()

def fit_text(draw, text, max_width, start_size, kind, font_dir, min_size=20):
    size=start_size
    while size>=min_size:
        f=F(size,kind,font_dir)
        if draw.textbbox((0,0), text, font=f)[2] <= max_width:
            return f
        size-=2
    return F(min_size,kind,font_dir)

def wrap(draw, text, font, max_width):
    words=str(text).split()
    lines=[]; cur=""
    for w in words:
        test=(cur+" "+w).strip()
        if draw.textbbox((0,0),test,font=font)[2] <= max_width:
            cur=test
        else:
            if cur: lines.append(cur)
            cur=w
    if cur: lines.append(cur)
    return lines

def text_block(draw, xy, text, font, fill, max_width, spacing=8, anchor="la", max_lines=None):
    x,y=xy
    lines=wrap(draw,text,font,max_width)
    if max_lines: lines=lines[:max_lines]
    bb=draw.textbbox((0,0),"Ag",font=font)
    lh=bb[3]-bb[1]
    for i,line in enumerate(lines):
        draw.text((x,y+i*(lh+spacing)),line,font=font,fill=fill,anchor=anchor)
    return y+len(lines)*(lh+spacing)-spacing

def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box,radius=radius,fill=fill,outline=outline,width=width)

def fit_crop(img, box):
    x1,y1,x2,y2=box
    tw,th=x2-x1,y2-y1
    src=img.copy().convert("RGB")
    sw,sh=src.size
    scale=max(tw/sw,th/sh)
    nw,nh=int(sw*scale),int(sh*scale)
    src=src.resize((nw,nh),Image.Resampling.LANCZOS)
    left=(nw-tw)//2; top=(nh-th)//2
    return src.crop((left,top,left+tw,top+th))

def draw_icon(draw, cx, cy, kind, scale=1.0, color=NAVY):
    # Simple line icons matching the reference style.
    s=scale
    if kind=="calendar":
        draw.rounded_rectangle((cx-25*s,cy-20*s,cx+25*s,cy+28*s),radius=5*s,outline=color,width=max(2,int(5*s)))
        draw.line((cx-25*s,cy-5*s,cx+25*s,cy-5*s),fill=color,width=max(2,int(4*s)))
        for dx in (-13,0,13):
            draw.ellipse((cx+dx*s-3*s,cy+4*s-3*s,cx+dx*s+3*s,cy+4*s+3*s),fill=color)
            draw.ellipse((cx+dx*s-3*s,cy+16*s-3*s,cx+dx*s+3*s,cy+16*s+3*s),fill=color)
        draw.line((cx-15*s,cy-27*s,cx-15*s,cy-17*s),fill=color,width=max(2,int(5*s)))
        draw.line((cx+15*s,cy-27*s,cx+15*s,cy-17*s),fill=color,width=max(2,int(5*s)))
    elif kind=="bell":
        draw.arc((cx-22*s,cy-25*s,cx+22*s,cy+20*s),180,360,fill=color,width=max(2,int(5*s)))
        draw.line((cx-22*s,cy-2*s,cx-27*s,cy+20*s),fill=color,width=max(2,int(5*s)))
        draw.line((cx+22*s,cy-2*s,cx+27*s,cy+20*s),fill=color,width=max(2,int(5*s)))
        draw.line((cx-27*s,cy+20*s,cx+27*s,cy+20*s),fill=color,width=max(2,int(5*s)))
        draw.arc((cx-7*s,cy+17*s,cx+7*s,cy+29*s),0,180,fill=color,width=max(2,int(4*s)))
        draw.ellipse((cx-4*s,cy-31*s,cx+4*s,cy-23*s),fill=color)
    elif kind=="people":
        draw.ellipse((cx-10*s,cy-26*s,cx+10*s,cy-6*s),fill=color)
        draw.ellipse((cx-31*s,cy-20*s,cx-15*s,cy-4*s),fill=color)
        draw.ellipse((cx+15*s,cy-20*s,cx+31*s,cy-4*s),fill=color)
        draw.pieslice((cx-28*s,cy-2*s,cx+28*s,cy+38*s),180,360,fill=color)
        draw.pieslice((cx-43*s,cy-1*s,cx-6*s,cy+31*s),180,360,fill=color)
        draw.pieslice((cx+6*s,cy-1*s,cx+43*s,cy+31*s),180,360,fill=color)
    elif kind=="heart":
        # outline heart
        pts=[]
        for t in range(0,361,5):
            a=math.radians(t)
            x=16*s*math.sin(a)**3
            y=-(13*s*math.cos(a)-5*s*math.cos(2*a)-2*s*math.cos(3*a)-math.cos(4*a))
            pts.append((cx+x,cy+y))
        draw.line(pts+[pts[0]],fill=color,width=max(2,int(4*s)),joint="curve")
    elif kind=="growth":
        draw.rectangle((cx-25*s,cy+5*s,cx-8*s,cy+25*s),fill=color)
        draw.rectangle((cx-3*s,cy-8*s,cx+14*s,cy+25*s),fill=color)
        draw.rectangle((cx+19*s,cy-25*s,cx+36*s,cy+25*s),fill=color)
        draw.line((cx-27*s,cy-2*s,cx+28*s,cy-32*s),fill=YELLOW,width=max(2,int(5*s)))
        draw.polygon([(cx+20*s,cy-30*s),(cx+31*s,cy-32*s),(cx+29*s,cy-21*s)],fill=YELLOW)
    elif kind=="check":
        draw.ellipse((cx-24*s,cy-24*s,cx+24*s,cy+24*s),outline=color,width=max(2,int(4*s)))
        draw.line((cx-12*s,cy,cx-2*s,cy+11*s),fill=color,width=max(2,int(5*s)))
        draw.line((cx-2*s,cy+11*s,cx+15*s,cy-10*s),fill=color,width=max(2,int(5*s)))

def load_rows(path):
    path=Path(path)
    if path.suffix.lower()==".csv":
        with open(path,encoding="utf-8-sig",newline="") as f:
            return list(csv.DictReader(f))
    if path.suffix.lower()==".xlsx":
        if load_workbook is None: raise RuntimeError("Install openpyxl to read XLSX files.")
        wb=load_workbook(path,data_only=True)
        ws=wb.active
        vals=list(ws.values)
        if not vals: return []
        headers=[str(x or "").strip() for x in vals[0]]
        return [dict(zip(headers,row)) for row in vals[1:] if any(v is not None for v in row)]
    raise ValueError("Calendar must be .csv or .xlsx")

def val(row,key,default=""):
    v=row.get(key,default)
    if v is None: return default
    # Allow spreadsheet/CSV cells to use either real line breaks or literal \n notation.
    return str(v).replace("\\n", "\n")

def generate(row, assets_dir, out_path, font_dir=None, size=1254):
    global W,H
    W=H=size
    S=size/1254.0
    sc=lambda n:int(n*S)
    canvas=Image.new("RGB",(W,H),WHITE)
    d=ImageDraw.Draw(canvas)

    # --- HERO / TOP SECTION ---
    hero_path=val(row,"image").strip()
    if hero_path:
        hp=Path(hero_path)
        if not hp.is_absolute(): hp=Path(assets_dir)/hero_path
    else:
        hp=Path(assets_dir)/"hero_default.jpg"
    hero=Image.open(hp).convert("RGB") if hp.exists() else Image.new("RGB",(600,500),(220,225,230))
    hero=fit_crop(hero,(sc(760),0,W,sc(505)))
    hero=ImageEnhance.Contrast(hero).enhance(1.02)
    canvas.paste(hero,(sc(760),0))

    # white right strip
    d.rectangle((sc(1032),0,W,sc(505)),fill="#FFFDF8")
    # Navy curved/angled panel
    pts=[(0,0),(sc(795),0),(sc(770),sc(105)),(sc(735),sc(230)),(sc(695),sc(350)),(sc(645),sc(430)),(sc(560),sc(480)),(sc(390),sc(500)),(0,sc(500))]
    d.polygon(pts,fill=NAVY)

    # logo
    logo_path=Path(assets_dir)/"serenehealth_logo.png"
    if logo_path.exists():
        lg=Image.open(logo_path).convert("RGBA")
        # keep aspect, fit 350x85
        ratio=min(sc(350)/lg.width,sc(85)/lg.height)
        lg=lg.resize((int(lg.width*ratio),int(lg.height*ratio)),Image.Resampling.LANCZOS)
        canvas.paste(lg,(sc(42),sc(18)),lg)
    else:
        d.text((sc(45),sc(38)),"SereneHealth",font=F(sc(34),"bold",font_dir),fill=WHITE)

    # sub-brand line
    d.text((sc(132),sc(77)),val(row,"brand_subtitle","QUALITY MEDICAL VIRTUAL ASSISTANT SERVICES"),
           font=F(sc(15),"semibold",font_dir),fill="#A7A8D5",anchor="lm")

    # day pill
    day=val(row,"day","WEDNESDAY").upper()
    rounded(d,(sc(536),sc(28),sc(754),sc(68)),sc(22),fill="#E9E9FF")
    draw_icon(d,sc(570),sc(48),"calendar",.38,NAVY)
    d.text((sc(600),sc(49)),day,font=F(sc(17),"bold",font_dir),fill=NAVY,anchor="lm")

    # right slogan
    slogan=val(row,"right_slogan","SUPPORTING\nHEALTHIER\nPRACTICES\nTOGETHER")
    d.multiline_text((sc(1100),sc(39)),slogan,font=F(sc(14),"semibold",font_dir),fill=NAVY,
                     anchor="ma",align="center",spacing=sc(4))
    d.rounded_rectangle((sc(1106),sc(126),sc(1190),sc(132)),radius=sc(3),fill=YELLOW)

    # headline
    hw=val(row,"headline_white","A Fuller Schedule.")
    hy=val(row,"headline_yellow","Fewer Missed Visits.")
    f1=fit_text(d,hw,sc(690),sc(54),"bold",font_dir,30)
    f2=fit_text(d,hy,sc(700),sc(54),"bold",font_dir,30)
    d.text((sc(48),sc(145)),hw,font=f1,fill=WHITE)
    d.text((sc(48),sc(213)),hy,font=f2,fill=YELLOW)

    sub=val(row,"subheadline","Appointment coordination made easier.")
    fs=fit_text(d,sub,sc(690),sc(34),"bold",font_dir,22)
    d.text((sc(50),sc(297)),sub,font=fs,fill=WHITE)
    d.rounded_rectangle((sc(50),sc(357),sc(168),sc(364)),radius=sc(3),fill=YELLOW)

    body=val(row,"body","SereneHealth helps support your front-end workflow so your day runs smoother.")
    fb=F(sc(25),"regular",font_dir)
    text_block(d,(sc(50),sc(387)),body,fb,WHITE,sc(655),spacing=sc(7),max_lines=2)

    # handwritten-style message
    hand=val(row,"handwritten","Your\nPractice\nOur Support.")
    d.multiline_text((sc(1165),sc(180)),hand,font=F(sc(23),"script",font_dir),fill=NAVY,
                     anchor="ma",align="center",spacing=sc(3))
    d.rounded_rectangle((sc(1110),sc(333),sc(1200),sc(338)),radius=sc(3),fill=YELLOW)

    # floating badge
    rounded(d,(sc(995),sc(380),sc(1230),sc(487)),sc(19),fill=CREAM)
    draw_icon(d,sc(1032),sc(433),"heart",.95,YELLOW)
    badge=val(row,"badge","More Time\nfor What\nMatters Most.")
    d.multiline_text((sc(1075),sc(404)),badge,font=F(sc(17),"bold",font_dir),fill=NAVY,spacing=sc(1))

    # --- SUPPORT CARD ---
    rounded(d,(sc(29),sc(515),sc(1224),sc(893)),sc(25),fill=CARD,outline="#E5E6F0",width=sc(2))
    title=val(row,"support_title","SereneHealth Supports:")
    rounded(d,(sc(48),sc(518),sc(574),sc(582)),sc(20),fill=YELLOW)
    d.text((sc(82),sc(550)),title,font=F(sc(29),"bold",font_dir),fill="#05052D",anchor="lm")

    items=[]
    for i in range(1,4):
        items.append({
            "title":val(row,f"item{i}_title",["Scheduling","Reminders","Confirmations"][i-1]),
            "desc":val(row,f"item{i}_desc",["Book appointments and manage your calendar.","Send patient reminders to reduce no-shows.","Follow up and confirm upcoming appointments."][i-1]),
            "icon":val(row,f"item{i}_icon",["calendar","bell","people"][i-1])
        })
    centers=[sc(215),sc(626),sc(1008)]
    for idx,(cx,item) in enumerate(zip(centers,items)):
        if idx<2:
            d.line((sc(429+idx*380),sc(602),sc(429+idx*380),sc(872)),fill=LINE,width=sc(1))
        d.ellipse((cx-sc(72),sc(592),cx+sc(72),sc(736)),fill=CREAM)
        draw_icon(d,cx,sc(664),item["icon"],1.0,NAVY)
        ft=fit_text(d,item["title"],sc(270),sc(31),"bold",font_dir,20)
        d.text((cx,sc(754)),item["title"],font=ft,fill=NAVY,anchor="ma")
        fd=F(sc(24),"regular",font_dir)
        text_block(d,(cx,sc(788)),item["desc"],fd,NAVY,sc(260),spacing=sc(4),anchor="ma",max_lines=3)

    # --- OUTCOME STRIP ---
    rounded(d,(sc(25),sc(908),sc(1228),sc(1044)),sc(22),fill=LAVENDER)
    draw_icon(d,sc(173),sc(976),"growth",.75,NAVY)
    d.line((sc(292),sc(926),sc(292),sc(1026)),fill="#BDBFE1",width=sc(2))
    outcome=val(row,"outcome","Improve workflow and\nkeep your calendar moving.")
    d.multiline_text((sc(357),sc(944)),outcome,font=F(sc(32),"bold",font_dir),fill=NAVY,spacing=sc(0))
    hand2=val(row,"outcome_script","Better Care.\nBrighter Days.")
    d.multiline_text((sc(1090),sc(952)),hand2,font=F(sc(21),"script",font_dir),fill=NAVY,anchor="ma",align="center",spacing=sc(2))
    d.rounded_rectangle((sc(1088),sc(1010),sc(1193),sc(1016)),radius=sc(3),fill=YELLOW)

    # --- FOOTER ---
    d.rectangle((0,sc(1051),W,H),fill=NAVY)
    d.rectangle((0,sc(1049),W,sc(1053)),fill=YELLOW)
    rounded(d,(sc(48),sc(1090),sc(688),sc(1184)),sc(45),fill=YELLOW)
    cta=val(row,"cta","Let’s support your practice")
    d.text((sc(100),sc(1137)),cta,font=fit_text(d,cta,sc(510),sc(29),"bold",font_dir,18),fill="#05052D",anchor="lm")
    d.text((sc(648),sc(1137)),"›",font=F(sc(44),"regular",font_dir),fill="#05052D",anchor="mm")
    d.line((sc(730),sc(1080),sc(730),sc(1220)),fill=YELLOW,width=sc(2))

    website=val(row,"website","SereneHealthVA.com")
    phone=val(row,"phone","(630) 394-2838")
    email=val(row,"email","Contact@serenehealthva.com")
    # simple footer icons
    for y,kind in [(1092,"globe"),(1139,"phone"),(1184,"mail")]:
        if kind=="phone":
            d.arc((sc(784),sc(y),sc(820),sc(y+36)),80,280,fill=YELLOW,width=sc(6))
        elif kind=="mail":
            d.rectangle((sc(782),sc(y+4),sc(824),sc(y+30)),outline=YELLOW,width=sc(4))
            d.line((sc(782),sc(y+4),sc(803),sc(y+20),sc(824),sc(y+4)),fill=YELLOW,width=sc(3))
        else:
            d.ellipse((sc(782),sc(y),sc(824),sc(y+42)),outline=WHITE,width=sc(3))
            d.arc((sc(787),sc(y+5),sc(819),sc(y+37)),90,270,fill=WHITE,width=sc(2))
            d.line((sc(803),sc(y),sc(803),sc(y+42)),fill=WHITE,width=sc(2))
    d.text((sc(854),sc(1109)),website,font=F(sc(23),"bold",font_dir),fill=WHITE,anchor="lm")
    d.text((sc(854),sc(1155)),phone,font=F(sc(23),"bold",font_dir),fill=WHITE,anchor="lm")
    d.text((sc(854),sc(1201)),email,font=F(sc(21),"bold",font_dir),fill=WHITE,anchor="lm")

    # output
    canvas.save(out_path,quality=96)
    return out_path

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--calendar",required=True)
    ap.add_argument("--output",default="output")
    ap.add_argument("--assets",default="assets")
    ap.add_argument("--font-dir",default="fonts")
    ap.add_argument("--size",type=int,default=1254)
    args=ap.parse_args()
    rows=load_rows(args.calendar)
    os.makedirs(args.output,exist_ok=True)
    if not rows: raise SystemExit("No rows found in calendar.")
    for i,row in enumerate(rows,1):
        slug=val(row,"filename").strip()
        if not slug:
            date=val(row,"date",str(i))
            slug=re.sub(r"[^A-Za-z0-9_-]+","-",date).strip("-") or f"graphic-{i:03d}"
        slug=Path(slug).stem
        slug=Path(slug).stem
        out=Path(args.output)/(slug+".jpg")
        generate(row,args.assets,out,args.font_dir,args.size)
        print("Created",out)

if __name__=="__main__":
    main()
