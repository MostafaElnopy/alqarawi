
from flask import Flask, render_template_string, request, jsonify, redirect, session, send_file
import json, os, glob, urllib.parse
from datetime import datetime, timedelta
import math
try:
 import openpyxl
 HAS_EXCEL=True
except:
 HAS_EXCEL=False

app = Flask(__name__, static_folder="static")
app.secret_key = "alqarawi123"
UPLOAD_FOLDER = "uploads/branches"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
DB_FILE = "alqarawi_db.json"

DEFAULT = {
  "branches": [
  {"name": "بريدة الرئيسي", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى9:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966536405559", "loc": "https://maps.app.goo.gl/2As6CtLgDfHn9eif6", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "الشارع التجارى", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "phone": "966533265553", "loc": "https://maps.app.goo.gl/kadYqCgimJiHbQ1Z7", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "شارع الخبيب", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "phone": "966553285559", "loc": "https://maps.app.goo.gl/RxVEChesgGBTmRir6", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "صناعية السليم1", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966553120500", "loc": "https://maps.app.goo.gl/MP8j7TyGdi7U9ZpM8", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "الزلفى", "city": "الزلفى", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966558685557", "loc": "https://maps.app.goo.gl/qJvWtaBy651mtKnW8", "img": "", "rating": 4.8, "lat": 26.3, "lng": 44.81},
  {"name": "الرس", "city": "الرس", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966558265557", "loc": "https://maps.app.goo.gl/sLsaG8JUVevvxryr7", "img": "", "rating": 4.8, "lat": 25.869, "lng": 43.498},
  {"name": "عنيزة", "city": "عنيزة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966558605551", "loc": "https://maps.app.goo.gl/5qYfzDkYvcSZJaxt9", "img": "", "rating": 4.8, "lat": 26.089, "lng": 43.993},
  {"name": "صناعية السليم2", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966538465558", "loc": "https://maps.app.goo.gl/dps5gZbMW6uDSPHL9", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "حفر الباطن", "city": "حفر الباطن", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966558615559", "loc": "https://maps.app.goo.gl/g3n2T5f1Vx3NSc75A", "img": "", "rating": 4.8, "lat": 28.432, "lng": 45.96},
  {"name": "صناعية الرواف", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966553405558", "loc": "https://maps.app.goo.gl/iDWhMc51rv2jepfH6", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "حائل صناعية الجربوع", "city": "حائل", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966507105558", "loc": "https://maps.app.goo.gl/na3TAJZHkHd71psW7", "img": "", "rating": 4.8, "lat": 27.52, "lng": 41.69},
  {"name": "العزيزية", "city": "المدينة المنورة", "hours1": "من 8 صباحا إلى الساعة 2:00 مساء", "hours2": "من بعد آذان العصر بنصف ساعة إلى اذان العشاء", "friday": "من بعد آذان العصر بنصف ساعة إلى اذان العشاء", "phone": "966558425557", "loc": "https://maps.app.goo.gl/3KeaTkvXTft66xyLA", "img": "", "rating": 4.8, "lat": 24.46, "lng": 39.61},
  {"name": "الكيلو 14", "city": "جدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "phone": "966550565553", "loc": "https://maps.app.goo.gl/oGDsNVQraZBwJniv5", "img": "", "rating": 4.8, "lat": 21.5433, "lng": 39.1728},
  {"name": "جدة قويزة", "city": "جدة", "hours1": "من 8 صباحا إلى آذان العشاء", "hours2": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966532036555", "loc": "https://maps.app.goo.gl/9odgwFct5Ny6M2Yz7", "img": "", "rating": 4.8, "lat": 21.5433, "lng": 39.1728},
  {"name": "المجمعة", "city": "المجمعة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966532087555", "loc": "https://maps.app.goo.gl/1qzD4VfNaFaLXN3eA", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "عروة", "city": "المدينة المنورة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "phone": "966532064555", "loc": "https://maps.app.goo.gl/6xcRJr6gTKb6Ath2A", "img": "", "rating": 4.8, "lat": 24.46, "lng": 39.61},
  {"name": "مكة", "city": "مكة المكرمة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "friday": "من بعد آذان العصر بنصف ساعة إلى8:30 مساء", "phone": "966533941555", "loc": "https://maps.app.goo.gl/9UU79jGiGUhDkRoRA", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "سكاكا", "city": "الجوف", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966532042555", "loc": "https://maps.app.goo.gl/ApgofeGqCEiJwiGQ9", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "الطائف", "city": "الطائف", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966532037555", "loc": "https://maps.app.goo.gl/Jcmhty6ETF7SQ5sb8", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "تبوك", "city": "تبوك", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966558682121", "loc": "", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "محايل عسير", "city": "محايل عسير", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966535482555", "loc": "https://maps.app.goo.gl/N6TK7T2VKyMo4eJX6", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "خميس مشيط", "city": "خمبس مشيط", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى آذان العشاء", "friday": "من الساعة2:00 مساء  إلى آذان العشاء", "phone": "966533042323", "loc": "https://maps.app.goo.gl/Crgy7YBvRHLHg8Wc9", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "النسيم1", "city": "الرياض", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "friday": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "phone": "966502382555", "loc": "https://maps.app.goo.gl/bYnV1SfrpKCvQxkA7", "img": "", "rating": 4.8, "lat": 24.7136, "lng": 46.6753},
  {"name": "ام الحمام", "city": "الرياض", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "friday": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "phone": "966536212555", "loc": "https://maps.app.goo.gl/91ar5EZyNxTN92Xo8", "img": "", "rating": 4.8, "lat": 24.7136, "lng": 46.6753},
  {"name": "النسيم2", "city": "الرياض", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "friday": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "phone": "966559842555", "loc": "", "img": "", "rating": 4.8, "lat": 24.7136, "lng": 46.6753},
  {"name": "صناعية الرمال", "city": "الرياض", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "friday": "من الساعة2:00 مساء  إلى الساعة 9:30 مساء", "phone": "966537412555", "loc": "", "img": "", "rating": 4.8, "lat": 24.7136, "lng": 46.6753}
 ],
 "products": [],
 "users": [{"name": "مصطفى", "pass": "1234", "perms": ["branches","ads","hours","add","users","marquee","reviews"]}],
 "reviews": [],
 "notify_phones": ["966163274444"],
 "main_phone": "966163274444",
 "marquee_text": "بريدة الرئيسي • الشارع التجارى • شارع الخبيب • صناعية السليم 1 و 2 • الزلفى • الرس • عنيزة • بريدة الرئيسي • دوام بعد آذان العصر بنصف ساعة إلى آذان العشاء • تقييمك يهمنا - شاركنا موقعك يجيبلك اقرب فرع • ٨ فروع - قطع أصلية"
}




def get_excel_path():
    for pat in ["*القرعاوي*.xlsx","*branches*.xlsx","*فروع*.xlsx","Alqarawi*.xlsx"]:
        for fp in glob.glob(pat):
            if not os.path.basename(fp).startswith("~$"):
                return fp
    return "branches.xlsx"

def normalize_phone(raw):
    if not raw: return ""
    s=''.join([c for c in str(raw) if c.isdigit()])
    if s.startswith('00'): s=s[2:]
    if s.startswith('0'): s='966'+s[1:]
    if not s.startswith('966') and len(s)<=10:
        if len(s)==10 and s.startswith('05'): s='966'+s[1:]
        elif len(s)==9: s='966'+s
        elif len(s)==10 and s.startswith('5'): s='966'+s
    return s or ""

def load_branches_from_excel():
    if not HAS_EXCEL: return None
    fpath=get_excel_path()
    if not os.path.exists(fpath): return None
    try:
        wb=openpyxl.load_workbook(fpath)
        ws=wb.active
        rows=list(ws.iter_rows(values_only=True))
        if not rows or len(rows)<2: return None
        branches=[]
        for r in rows[1:]:
            if not r or not r[0]: continue
            name=str(r[0] or "").strip()
            if not name: continue
            city=str(r[1] or "").strip() if len(r)>1 else "بريدة"
            phone=normalize_phone(str(r[2] or ""))
            loc=str(r[3] or "").strip() if len(r)>3 else ""
            img=str(r[4] or "").strip() if len(r)>4 else ""
            h1=str(r[5] or "").strip() if len(r)>5 else "من 8ص إلى آذان الظهر"
            h2=str(r[6] or "").strip() if len(r)>6 else "من بعد آذان العصر بنصف ساعة إلى آذان العشاء"
            fri=str(r[7] or "").strip() if len(r)>7 else h2
            # Try parse lat/lng from loc or extra columns
            lat=26.359; lng=43.968
            if "بريدة" in city: lat=26.359; lng=43.968
            if "الرياض" in city: lat=24.7136; lng=46.6753
            if "جدة" in city: lat=21.5433; lng=39.1728
            b={"name":name,"city":city,"hours1":h1,"hours2":h2,"friday":fri,"phone":phone or "966163274444","loc":loc,"img":img,"rating":4.8,"lat":lat,"lng":lng}
            branches.append(b)
        if branches: return branches
    except Exception as e:
        print(e)
    return None

def load():
    eb=load_branches_from_excel()
    if not os.path.exists(DB_FILE):
        db=json.loads(json.dumps(DEFAULT))
        if eb: db["branches"]=eb
        return db
    try:
        d=json.load(open(DB_FILE, encoding="utf-8"))
        for k in DEFAULT:
            if k not in d: d[k]=DEFAULT[k]
        if eb: d["branches"]=eb
        if "notify_phones" not in d: d["notify_phones"]=[d.get("main_phone","966163274444")]
        if "marquee_text" not in d: d["marquee_text"]=DEFAULT["marquee_text"]
        return d
    except:
        db=json.loads(json.dumps(DEFAULT))
        if eb: db["branches"]=eb
        return db

def save(db):
    json.dump(db, open(DB_FILE,"w",encoding="utf-8"), ensure_ascii=False, indent=2)

HTML = """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>شركة القرعاوي - صناعية السليم</title>
<style>
body{margin:0;font-family:Arial;background:#fff;color:#000}
header{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#fff;border-bottom:1px solid #eee;position:sticky;top:0;z-index:20}
.btn{padding:10px 18px;border-radius:999px;font-weight:800;margin:4px;text-decoration:none;display:inline-block}
.btn-g{background:#22c55e;color:#fff}.btn-b{background:#000;color:#fff}.btn-r{background:#e11d48;color:#fff}
.section{padding:20px 16px;max-width:1100px;margin:auto}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:16px}
.card{border:1px solid #ddd;border-radius:16px;overflow:hidden;background:#fff;position:relative;color:#000}
.card-body{padding:14px;line-height:1.8;color:#000;background:#fff} .card-body h3{color:#000;font-weight:900;font-size:18px;margin:8px 0} .card-body p{color:#111} .card-body b{color:#000}
.hours{font-size:13px;background:#f9fafb;border-radius:8px;padding:10px;margin:6px 0;color:#000;border:1px solid #eee} .hours small{color:#444!important}
.open-badge{position:absolute;top:10px;left:10px;padding:8px 14px;border-radius:999px;font-size:13px;font-weight:900;z-index:5;border:1px solid rgba(0,0,0,0.1);color:#fff!important;text-shadow:0 1px 1px rgba(0,0,0,0.3)}
.open{ background:#16a34a!important;color:#fff!important; border:1px solid #15803d } .closed{ background:#dc2626!important;color:#fff!important; } .calculating{ background:#d97706!important;color:#fff!important; }
.search-box{border:2px solid #000;border-radius:12px;padding:16px;margin:12px 0;background:#f9fafb;box-sizing:border-box;overflow:hidden} .search-box *{box-sizing:border-box} .search-box input{max-width:100%;margin:8px 0;display:block}
@keyframes pulse{0%{transform:scale(1)}50%{transform:scale(1.05)}100%{transform:scale(1)}} .voice-btn{background:#7c3aed;color:#fff;padding:12px 18px;border-radius:999px;border:0;font-weight:800;font-size:16px;cursor:pointer}
.review-box{border:2px solid #22c55e;border-radius:12px;padding:12px;margin:8px 0;background:#f0fdf4}
.footer{background:#0a0a0a;color:#fff;text-align:center;padding:24px}
#sidebar{position:fixed;top:0;right:-320px;width:300px;height:100vh;background:#fff;border-left:3px solid #000;z-index:200;transition:0.3s;overflow-y:auto;padding:16px;box-shadow:-4px 0 20px rgba(0,0,0,0.2)}
#sidebar.open{right:0}
#overlay{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:150;display:none} #overlay.show{display:block}
</style></head>
<body>
<div id="overlay" onclick="closeSidebar()"></div>
<div id="sidebar">
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;border-bottom:2px solid #000;padding-bottom:12px"><b>القائمة - ٨ فروع</b><button onclick="closeSidebar()" style="background:#000;color:#fff;border:0;padding:8px 14px;border-radius:999px">✕</button></div>
<div style="display:flex;flex-direction:column;gap:8px"><a href="/" style="padding:12px;border-radius:8px;background:#f9fafb;border:1px solid #ddd;text-decoration:none;color:#000;font-weight:800">🏠 الرئيسية</a><a href="#branches" onclick="closeSidebar()" style="padding:12px;border-radius:8px;background:#f9fafb;border:1px solid #ddd;text-decoration:none;color:#000;font-weight:800">📍 فروعنا - ٨ فروع</a><a href="#ads" onclick="closeSidebar()" style="padding:12px;border-radius:8px;background:#fef2f2;border:1px solid #e11d48;text-decoration:none;color:#e11d48;font-weight:800">📢 إعلاناتنا</a></div></div>

<header style="display:flex;justify-content:space-between;align-items:center;padding:12px 20px;background:linear-gradient(135deg,#0a0a0a 0%,#1a1a1a 100%);border-bottom:3px solid #e11d48;position:sticky;top:0;z-index:100;box-shadow:0 4px 20px rgba(0,0,0,0.5)"><div style="display:flex;align-items:center;gap:14px"><img src="/static/logo.jpg" style="width:72px;height:72px;border-radius:14px;border:3px solid #fff;object-fit:cover;background:#fff;box-shadow:0 4px 12px rgba(0,0,0,0.3)" onerror="this.src='/static/logo.jpg'"><div style="color:#fff"><b style="font-size:20px;color:#fff;letter-spacing:1px;display:block">شركة القرعاوي</b><small style="color:#e11d48;font-weight:800;letter-spacing:2px">TOYOTA - قطع أصلية - ٨ فروع</small></div></div><div style="display:flex;gap:8px;align-items:center"><a class="btn btn-g" href="https://api.whatsapp.com/send?phone={{main_phone}}" style="padding:10px 18px;font-size:13px;box-shadow:0 2px 8px rgba(34,197,94,0.3)">واتساب</a><button onclick="openSidebar()" style="padding:10px 16px;background:#fff;color:#000;border:0;border-radius:999px;font-weight:900;box-shadow:0 2px 8px rgba(255,255,255,0.2)">☰</button></div></header>

<div style="background:#000;color:#fff;overflow:hidden;white-space:nowrap;padding:10px 0;font-size:15px;font-weight:900;line-height:1.5;border-bottom:2px solid #222"><marquee scrollamount="6" direction="right"> {{marquee_text}} </marquee></div>

<!-- 2: VIN Search + 6: Voice + 10: Share Location -->
<div class="section" style="background:#fefce8;border-bottom:2px solid #eab308">
<h2 style="margin:0 0 12px">🔍 ابحث بذكاء - ٣ طرق</h2>
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px">

<div class="search-box" style="border-color:#e11d48;background:linear-gradient(135deg,#fef2f2 0%,#fff 100%);box-shadow:0 4px 16px rgba(225,29,72,0.15);border-width:3px">
<h3 style="margin:0 0 8px;color:#000;font-size:18px">🔍 بحث برقم الشاصي VIN (17 رقم) - التصميم الفخم</h3>
<p style="font-size:13px;margin:4px 0;color:#444">اكتب رقم الشاصي - اختار الفرع - اكتب المطلوب - صور القطعة - يرسل واتس خدمة العملاء - الشريط والإعلانات شغالين</p>

<label style="font-weight:800;font-size:13px;margin-top:10px;display:block">١- رقم الشاصي - 17 حرف انجليزي - يتحول كبير أوتوماتيك:</label>
<input id="vinInput" placeholder="JTDKB20U093456789 - 17 حرف انجليزي بالظبط" maxlength="17" pattern="[A-Za-z0-9]{17}" title="17 حرف ورقم انجليزي بالظبط" style="width:100%;padding:14px;border-radius:10px;border:3px solid #e11d48;font-size:18px;direction:ltr;text-transform:uppercase;letter-spacing:3px;font-weight:900;background:#fff;color:#000;box-shadow:inset 0 2px 8px rgba(0,0,0,0.05)" oninput="let v=this.value.toUpperCase().replace(/[^A-Z0-9]/g,''); if(this.value!==v) this.value=v; document.getElementById('vinCount').innerHTML=this.value.length+'/17 - '+(this.value.length==17?'✅ مظبوط - أحرف وأرقام انجليزي كبيرة':'❌ لازم 17 - '+(17-this.value.length)+' باقي')">

<div id="vinCount" style="font-size:13px;margin-top:8px;color:#111;font-weight:800;background:#fff;padding:8px 12px;border-radius:8px;border:2px solid #e11d48;display:inline-block;box-shadow:0 2px 6px rgba(0,0,0,0.05)">0/17 - لازم 17 بالظبط - انجليزي فقط - يكبر أوتوماتيك</div>

<label style="font-weight:800;font-size:13px;margin-top:12px;display:block">٢- اختار الفرع - قائمة منسدلة - العميل هو اللي يختار - من ملفك - 26 فرع:</label>
<select id="vinBranch" style="width:100%;padding:12px;border-radius:10px;border:2px solid #000;font-size:15px;background:#fff;color:#000;font-weight:700">
<option value="">اختار الفرع - 26 فرع - العميل يختار</option>
{% for b in branches %}<option value="{{b.name}} - {{b.city}}">{{b.name}} - {{b.city}} - {{b.phone}}</option>{% endfor %}
</select>

<label style="font-weight:800;font-size:13px;margin-top:12px;display:block">٣- العميل يكتب المطلوب:</label>
<textarea id="vinNeed" placeholder="مثال: عايز مساعد امامي يمين كامري 2018 - فلتر زيت - تيل فرامل - اكتب كل اللي محتاجه - بالتفصيل" style="width:100%;padding:12px;border-radius:10px;border:2px solid #000;min-height:80px;font-size:14px;background:#fff;color:#000"></textarea>

<label style="font-weight:800;font-size:13px;margin-top:12px;display:block">٤- صور بالكاميرا أو رفع صورة أو أكتر - صور القطعة القديمة أو الشاصي:</label>
<div style="display:flex;gap:8px;flex-wrap:wrap;margin:8px 0">
<label style="flex:1;min-width:120px;padding:12px;background:#000;color:#fff;border-radius:10px;text-align:center;cursor:pointer;font-weight:800;border:2px solid #000">📷 صور بالكاميرا<input id="vinCam" type="file" accept="image/*" capture="environment" multiple style="display:none" onchange="previewVinImages(this)"></label>
<label style="flex:1;min-width:120px;padding:12px;background:#fff;color:#000;border-radius:10px;text-align:center;cursor:pointer;font-weight:800;border:2px solid #000">🖼️ رفع صورة/صور<input id="vinImgs" type="file" accept="image/*" multiple style="display:none" onchange="previewVinImages(this)"></label>
</div>
<div id="vinPreview" style="display:flex;gap:8px;flex-wrap:wrap;margin:8px 0"></div>
<div id="vinPreviewCount" style="font-size:12px;color:#666"></div>

<button onclick="searchVIN()" style="width:100%;margin-top:12px;padding:16px;background:linear-gradient(135deg,#e11d48 0%,#991b1b 100%);color:#fff;border:0;border-radius:12px;font-weight:900;font-size:16px;box-shadow:0 4px 16px rgba(225,29,72,0.4)">📤 ارسل بالشاصي والمطلوب والصور - واتس خدمة العملاء - {{main_phone}}</button>
<div id="vinResult" style="margin-top:12px;font-size:13px;background:#fff;padding:12px;border-radius:10px;border:1px solid #eee"></div>
</div>

<div class="search-box" style="border-color:#7c3aed;background:#f5f3ff">
<h3 style="margin:0 0 8px">٦- اطلب بصوتك</h3>
<p style="font-size:13px;margin:4px 0">دوس وتكلم: "عايز مساعد امامي كامري 2018"</p>
<button id="voiceBtn" class="voice-btn" onclick="startVoice()">🎤 اضغط وتكلم - اطلب قطعتي</button>
<div id="voiceText" style="margin-top:8px;padding:8px;background:#fff;border-radius:8px;min-height:40px;border:1px dashed #7c3aed">الكلام هيظهر هنا...</div>
<button id="voiceSend" onclick="sendVoice()" style="display:none;width:100%;margin-top:8px;padding:12px;background:#7c3aed;color:#fff;border:0;border-radius:8px;font-weight:800;box-shadow:0 2px 8px rgba(124,58,237,0.3)">📤 ارسل الطلب الصوتي واتساب - {{main_phone}}</button>
</div>

<div class="search-box" style="border-color:#22c55e;background:#f0fdf4">
<h3 style="margin:0 0 8px">١٠- شاركنا موقعك - نجيبلك اقرب فرع</h3>
<p style="font-size:13px;margin:4px 0">نحسب المسافة لصناعية السليم ١ و ٢ ونوديك</p>
<button onclick="shareLocation()" style="width:100%;padding:12px;background:#22c55e;color:#fff;border:0;border-radius:8px;font-weight:800">📍 شارك موقعي - احسب اقرب فرع</button>
<div id="locationResult" style="margin-top:8px;font-size:13px"></div>
</div>

</div>
</div>

<div class="section" style="background:#fef2f2;border-bottom:2px solid #e11d48;padding:16px">
<h2 style="margin:0 0 12px">📢 إعلاناتنا - عروض قطع الغيار الأصلية</h2>
{% if products %}
<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px">
{% for p in products %}
<div style="border:2px solid #e11d48;border-radius:12px;padding:12px;background:#fff">
<b style="color:#e11d48">{{p.title}}</b><br>
<small>{{p.desc}}</small><br>
<small style="color:#666">{{p.date}}</small>
</div>
{% endfor %}
</div>
{% else %}
<p style="font-size:13px;color:#666">لا يوجد إعلانات حاليا - تابعنا - اضف اعلان من لوحة الإدارة - صلاحية الإعلانات</p>
{% endif %}
</div>
<div class="section" id="branches"><h2>فروعنا - صناعية السليم ١ و ٢ - بريدة</h2><p style="font-size:13px;color:#666">٥- مفتوح الآن - يحسب حسب آذان المدينة اللي في عمود المدينة</p>
<div class="cards">{% for b in branches %}<div class="card" data-city="{{b.city}}" data-lat="{{b.lat}}" data-lng="{{b.lng}}" data-name="{{b.name}}">
<div class="open-badge calculating" id="open-{{loop.index0}}">⏳ يحسب الآذان - {{b.city}}...</div>
<div class="card-body" style="padding-top:36px">
<h3>{{b.name}}</h3>
<p><b>{{b.city}}</b> - {{b.phone}}</p>
<div class="hours">☀️ صباحا: {{b.hours1}} <br><small style="color:#666">حتى آذان الظهر - {{b.city}}</small></div>
<div class="hours">🌙 مساء: {{b.hours2}} <br><small style="color:#666">بعد آذان العصر بنصف ساعة إلى آذان العشاء - {{b.city}}</small></div>
<div class="hours">🕌 الجمعة: {{b.friday}}</div>
<div id="prayer-times-{{loop.index0}}" style="font-size:11px;background:#000;color:#fff;border-radius:8px;padding:6px;margin:6px 0">جاري جلب آذان {{b.city}}...</div>
<div style="display:flex;gap:6px;flex-wrap:wrap;margin:8px 0"><a class="btn btn-g" href="https://api.whatsapp.com/send?phone={{b.phone}}" style="flex:1;text-align:center;min-width:80px">💬 واتساب</a><a class="btn btn-b" href="tel:+{{b.phone}}" style="flex:1;text-align:center;min-width:80px;background:#000;color:#fff">📞 اتصال</a><a class="btn btn-b" href="{{b.loc}}" style="flex:1;text-align:center;min-width:80px;background:#fff;color:#000;border:1px solid #000">📍 الموقع</a></div>
<div class="review-box">
<b>⭐ قيّم {{b.name}}</b>
<textarea id="comment-{{loop.index0}}" placeholder="رأيك يهمنا" style="width:100%;padding:8px;margin:4px 0;border-radius:8px;border:1px solid #ddd"></textarea>
<button onclick="sendReview({{loop.index0}}, '{{b.name}}')" style="width:100%;padding:10px;background:#22c55e;color:#fff;border:0;border-radius:8px">إرسال التقييم - يوصل واتساب</button>
<div id="msg-{{loop.index0}}" style="margin-top:6px;font-size:13px"></div>
</div>
</div></div>{% endfor %}</div></div>

<div class="footer">
<p style="font-size:18px;font-weight:800;margin:0">شركة القرعاوي لبيع قطع غيار تويوتا الأصلية</p>
<p style="margin:8px 0">© شركة القرعاوي - {{main_phone}} - </p>
<p id="footer-prayer" style="font-size:12px;color:#aaa">جاري حساب آذان بريدة - دوام بعد العصر بنصف ساعة إلى العشاء</p>

<div style="margin-top:16px;display:flex;justify-content:center;gap:8px;flex-wrap:wrap">
<a href="https://api.whatsapp.com/send?phone={{main_phone}}" style="background:#22c55e;color:#fff;padding:10px 18px;border-radius:999px;text-decoration:none;font-weight:800">💬 واتساب {{main_phone}}</a>
<a href="tel:+{{main_phone}}" style="background:#000;color:#fff;padding:10px 18px;border-radius:999px;text-decoration:none;border:1px solid #333">📞 اتصال</a>
</div>

<div style="margin-top:20px;padding-top:15px;border-top:1px solid #222;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">
<div style="font-size:11px;opacity:0.35;text-align:right">
<a href="/staff-login-2024" style="color:#666;text-decoration:none;font-size:11px">⚙️ دخول الموظفين</a>
<span style="margin:0 6px">|</span>
<span>نظام داخلي</span>
</div>
<div style="font-size:10px;opacity:0.35;display:flex;align-items:center;gap:6px;background:#111;padding:6px 12px;border-radius:999px;border:1px solid #222">
<span style="background:#666;color:#fff;font-weight:900;padding:2px 6px;border-radius:4px;font-size:9px">ME</span>
<span style="color:#666">تصميم</span>
<span style="color:#888;font-weight:700">مصطفى النوبي</span>
<span style="color:#444">|</span>
<span style="color:#555">Mostafa Elnopy</span>
</div>
</div>

<div style="margin-top:12px;font-size:10px;color:#444;text-align:center">
📍 مواقع دقيقة - ✔ قطع أصلية - ⭐ تقييمك يهمنا - 🔍 بحث بالشاصي 17 رقم - 🎤 اطلب بصوتك - 📱 شارك موقعك
</div>
</div>

<script>
// 5- مفتوح الآن حسب آذان المدينة اللي في عمود المدينة
async function loadPrayerTimes(city, idx){
 try{
  let apiCity = city;
  if(city.includes('بريدة')||city.includes('بريده')) apiCity='Buraidah';
  if(city.includes('الرياض')) apiCity='Riyadh';
  if(city.includes('جدة')||city.includes('جده')) apiCity='Jeddah';
  let url = `https://api.aladhan.com/v1/timingsByCity?city=${apiCity}&country=Saudi Arabia&method=4`;
  let res = await fetch(url);
  let data = await res.json();
  if(!data || !data.data) throw 'no data';
  let t = data.data.timings;
  // ت = Dhuhr, Asr, Maghrib, Isha
  document.getElementById('prayer-times-'+idx).innerHTML = `آذان ${city}: الظهر ${t.Dhuhr} - العصر ${t.Asr} - المغرب ${t.Maghrib} - العشاء ${t.Isha}`;
  document.getElementById('footer-prayer').innerHTML = `آذان ${city} اليوم: الظهر ${t.Dhuhr} - العصر ${t.Asr} - العشاء ${t.Isha} - دوامنا بعد العصر بنصف ساعة إلى العشاء`;

  // حساب مفتوح الآن
  let now = new Date();
  let nowMin = now.getHours()*60 + now.getMinutes();
  function toMin(s){ let parts=s.split(':'); return parseInt(parts[0])*60+parseInt(parts[1]); }
  let dhuhr = toMin(t.Dhuhr);
  let asr = toMin(t.Asr);
  let isha = toMin(t.Isha);
  let maghrib = toMin(t.Maghrib);
  let openMorningStart = 8*60; // 8ص
  let openMorningEnd = dhuhr;
  let openEveningStart = asr + 30; // بعد العصر بنصف ساعة
  let openEveningEnd = isha + 30; // الى العشاء + نص ساعة سماح

  let badge = document.getElementById('open-'+idx);
  let isOpen = (nowMin >= openMorningStart && nowMin <= openMorningEnd) || (nowMin >= openEveningStart && nowMin <= openEveningEnd);
  // الجمعة
  let isFriday = now.getDay()===5;
  if(isFriday){
   isOpen = (nowMin >= openEveningStart && nowMin <= openEveningEnd);
  }

  if(isOpen){
   badge.className='open-badge open';
   badge.innerHTML='🟢 مفتوح الآن - '+city;
  }else{
   badge.className='open-badge closed';
   let nextOpen = '';
   if(nowMin < openMorningStart) nextOpen='يفتح 8ص';
   else if(nowMin < openEveningStart) nextOpen='يفتح بعد العصر بنصف ساعة ('+t.Asr+')';
   else nextOpen='يفتح غدا 8ص';
   badge.innerHTML='🔴 مغلق - '+nextOpen+' - '+city;
  }
 }catch(e){
  document.getElementById('prayer-times-'+idx).innerHTML='آذان '+city+': من 8ص إلى الظهر - بعد العصر بنصف ساعة إلى العشاء';
  document.getElementById('open-'+idx).className='open-badge calculating';
  document.getElementById('open-'+idx).innerHTML='⏰ دوام '+city+': بعد العصر بنصف ساعة إلى العشاء';
 }
}

// حمل لكل فرع حسب المدينة اللي في عمود المدينة
document.querySelectorAll('.card').forEach((card,i)=>{
 let city = card.dataset.city || 'بريدة';
 loadPrayerTimes(city, i);
});

// 10- شاركنا موقعك - نجيبلك اقرب فرع
function haversine(lat1,lon1,lat2,lon2){
 const R=6371;
 const dLat=(lat2-lat1)*Math.PI/180;
 const dLon=(lon2-lon1)*Math.PI/180;
 const a=Math.sin(dLat/2)*Math.sin(dLat/2)+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)*Math.sin(dLon/2);
 const c=2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
 return R*c;
}
function shareLocation(){
 const result=document.getElementById('locationResult');
 if(!navigator.geolocation){ result.innerHTML='المتصفح لا يدعم الموقع'; return; }
 result.innerHTML='⏳ جاري تحديد موقعك...';
 navigator.geolocation.getCurrentPosition(pos=>{
  let myLat=pos.coords.latitude; let myLng=pos.coords.longitude;
  let branches=[];
  document.querySelectorAll('.card').forEach(card=>{
   let lat=parseFloat(card.dataset.lat); let lng=parseFloat(card.dataset.lng); let name=card.dataset.name;
   if(!isNaN(lat)&&!isNaN(lng)){
    let dist=haversine(myLat,myLng,lat,lng);
    branches.push({name:name, lat:lat, lng:lng, dist:dist});
   }
  });
  branches.sort((a,b)=>a.dist-b.dist);
  if(branches.length===0){ result.innerHTML='لا يوجد فروع باحداثيات - اضف احداثيات في الاكسيل'; return; }
  let nearest=branches[0];
  result.innerHTML=`📍 انت هنا: ${myLat.toFixed(4)}, ${myLng.toFixed(4)}<br>🏆 اقرب فرع: <b>${nearest.name}</b> - ${nearest.dist.toFixed(2)} كم<br><a href="https://www.google.com/maps/dir/${myLat},${myLng}/${nearest.lat},${nearest.lng}" target="_blank" style="background:#22c55e;color:#fff;padding:8px 12px;border-radius:999px;text-decoration:none;display:inline-block;margin-top:6px">🧭 اتجاهات خرائط جوجل - ${nearest.dist.toFixed(2)} كم</a><br><br>كل الفروع:<br>${branches.map(b=>`${b.name}: ${b.dist.toFixed(2)} كم - <a href="https://www.google.com/maps/dir/${myLat},${myLng}/${b.lat},${b.lng}" target="_blank">اتجاهات</a>`).join('<br>')}`;
 }, err=>{
  result.innerHTML='خطأ: '+err.message+' - فعل GPS واسمح للموقع';
 });
}

// 2- بحث برقم الشاصي VIN
function openSidebar(){document.getElementById('sidebar').classList.add('open');document.getElementById('overlay').classList.add('show')}
function closeSidebar(){document.getElementById('sidebar').classList.remove('open');document.getElementById('overlay').classList.remove('show')}
let vinFiles=[];
function previewVinImages(input){
 let preview=document.getElementById('vinPreview');
 let count=document.getElementById('vinPreviewCount');
 vinFiles=Array.from(input.files);
 preview.innerHTML='';
 vinFiles.forEach((file,i)=>{
  let reader=new FileReader();
  reader.onload=function(e){
   let div=document.createElement('div');
   div.style.cssText='position:relative;width:80px;height:80px;border:2px solid #000;border-radius:8px;overflow:hidden';
   div.innerHTML=`<img src="${e.target.result}" style="width:100%;height:100%;object-fit:cover"><span style="position:absolute;top:2px;right:2px;background:#000;color:#fff;border-radius:50%;width:20px;height:20px;display:flex;align-items:center;justify-content:center;font-size:12px;cursor:pointer" onclick="this.parentElement.remove()">✕</span>`;
   preview.appendChild(div);
  };
  reader.readAsDataURL(file);
 });
 count.innerHTML=`✅ تم اختيار ${vinFiles.length} صورة - ستُرسل مع رسالة الواتساب - بعد فتح الواتساب ارفق الصور من المعرض`;
}

function searchVIN(){
 let vin=document.getElementById('vinInput').value.trim().toUpperCase();
 let branch=document.getElementById('vinBranch').value;
 let need=document.getElementById('vinNeed').value.trim();
 let res=document.getElementById('vinResult');
 if(vin.length<10){ res.innerHTML='❌ اكتب رقم الشاصي كامل - 17 حرف ورقم انجليزي'; return; }
 if(vin.length!==17){ res.innerHTML='❌ رقم الشاصي لازم 17 حرف ورقم - انت كاتب '+vin.length+' - باقي '+(17-vin.length); return; }
 if(!branch){ res.innerHTML='❌ اختار الفرع من القائمة المنسدلة - ٨ فروع'; return; }
 if(!need){ res.innerHTML='❌ اكتب المطلوب - العميل يكتب المطلوب - مثال: مساعد امامي كامري 2018'; return; }
 let hasImages = vinFiles && vinFiles.length>0;
 let msg=`🚗 طلب قطع برقم الشاصي - شركة القرعاوي الفخمة
━━━━━━━━━━━━━━━━━━
🔢 رقم الشاصي VIN: ${vin}
📍 الفرع المختار: ${branch}
📝 المطلوب: ${need}
📸 الصور المرفقة: ${hasImages ? vinFiles.length+' صورة - انظر المرفقات' : 'لا يوجد - العميل لم يرفق صور'}
━━━━━━━━━━━━━━━━━━
🏢 شركة القرعاوي - قطع غيار تويوتا أصلية
📞 خدمة العملاء: {{main_phone}}
🌐 https://alqarawi.onrender.com
تم عبر الموقع - بحث الشاصي + كتابة + صور - تصميم فخم`;
 let enc=encodeURIComponent(msg);
 let wa=`https://api.whatsapp.com/send?phone={{main_phone}}&text=${enc}`;
 res.innerHTML=`✅ تم تجهيز الطلب الفخم:<br>🔢 الشاصي: <b style="letter-spacing:2px;background:#000;color:#fff;padding:4px 8px;border-radius:6px">${vin}</b><br>📍 الفرع: <b>${branch}</b><br>📝 المطلوب: ${need}<br>📸 الصور: ${hasImages ? vinFiles.length+' صورة جاهزة' : 'بدون صور'}<br><br>
 <div style="background:#fefce8;border:2px solid #eab308;border-radius:10px;padding:10px;margin:8px 0;font-size:12px">
 ⚠️ مهم: بعد ما تدوس واتساب - افتح الواتساب - الصور اللي اخترتها (${hasImages ? vinFiles.length : 0}) - لازم ترفقها يدوي من زر المرفقات 📎 - عشان واتساب ما يسمحش ارسال صور اوتوماتيك من الموقع
 </div>
 <a href="${wa}" target="_blank" style="background:linear-gradient(135deg,#e11d48 0%,#991b1b 100%);color:#fff;padding:14px 24px;border-radius:999px;text-decoration:none;display:inline-block;margin-top:8px;font-weight:900;box-shadow:0 4px 16px rgba(225,29,72,0.4)">📤 افتح واتساب خدمة العملاء - ارسل الشاصي والكتابة - وارفق الصور</a>
 <div style="margin-top:12px"><small>الشعار الرخامي - خلفية سودة - شركة القرعاوي - تصميم فخم</small></div>`;
 // Try Web Share API with files if available - for mobile
 if(navigator.share && hasImages && navigator.canShare && navigator.canShare({files: vinFiles})){
   navigator.share({title: 'طلب قطع '+vin, text: msg, files: vinFiles}).then(()=>{ res.innerHTML+='<br>✅ تمت المشاركة بالصور - واتساب'; }).catch(()=>{ setTimeout(()=>{window.open(wa,'_blank');},800); });
 } else {
   setTimeout(()=>{window.open(wa,'_blank');},800);
 }
}

// 6- اطلب بصوتك
let recognition;
let voiceTranscript='';
let isRecording=false;
function startVoice(){
 let btn=document.getElementById('voiceBtn');
 let textDiv=document.getElementById('voiceText');
 let sendBtn=document.getElementById('voiceSend');
 if(!('webkitSpeechRecognition' in window || 'SpeechRecognition' in window)){
  textDiv.innerHTML='❌ المتصفح لا يدعم التعرف على الصوت - استخدم Chrome على الجوال - أو اكتب طلبك واتساب<br>💡 حل: افتح الموقع بـ Chrome - مش فيسبوك - دوس 3 نقط - افتح في Chrome';
  return;
 }
 if(isRecording && recognition){
  try{ recognition.stop(); }catch(e){}
  isRecording=false;
  btn.innerHTML='🎤 اضغط وتكلم - اطلب قطعتي';
  btn.style.background='#7c3aed';
  return;
 }
 let SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
 try{
  recognition = new SpeechRecognition();
  recognition.lang='ar-SA';
  recognition.lang='ar';
  recognition.interimResults=true;
  recognition.continuous=false;
  recognition.maxAlternatives=3;
  btn.innerHTML='🔴 اسمعك... تكلم الآن - دوس تاني للوقف';
  btn.style.background='#ef4444';
  btn.style.animation='pulse 1s infinite';
  textDiv.innerHTML='🎤 اسمعك... قول: عايز مساعد امامي كامري 2018 - اتكلم بوضوح - قرب الموبايل منك';
  textDiv.style.background='#fef2f2';
  textDiv.style.border='2px solid #ef4444';
  recognition.start();
  isRecording=true;
  recognition.onstart = function(){
   textDiv.innerHTML='🎙️ بدأ التسجيل... اتكلم الآن - قول طلبك';
   btn.innerHTML='🔴 بسجل... دوس للوقف';
  };
  recognition.onresult = function(event){
   let transcript='';
   for(let i=0;i<event.results.length;i++){
    transcript+=event.results[i][0].transcript+' ';
   }
   voiceTranscript = transcript.trim();
   textDiv.innerHTML='✅ سمعتك: <b style="font-size:16px;color:#000;background:#fff;padding:4px 8px;border-radius:6px">'+voiceTranscript+'</b><br><small>لو صح دوس ارسال واتساب - لو غلط دوس تاني وتكلم</small>';
   textDiv.style.background='#f0fdf4';
   textDiv.style.border='2px solid #22c55e';
   if(event.results[0].isFinal){
    sendBtn.style.display='block';
    btn.innerHTML='🎤 اضغط وتكلم مرة اخرى - لو عايز تغير';
    btn.style.background='#7c3aed';
    btn.style.animation='none';
    isRecording=false;
   }
  };
  recognition.onerror = function(e){
   console.log(e);
   let msg='خطأ: '+e.error;
   if(e.error==='not-allowed' || e.error==='permission-denied'){
    msg='❌ السماح مرفوض - دوس على علامة القفل فوق - اسمح للميكروفون - أو افتح الإعدادات - اسمح للموقع بالميكروفون - وبعدين دوس تاني';
   } else if(e.error==='no-speech'){
    msg='❌ ما سمعتش حاجة - اتكلم بصوت أعلى - قرب الموبايل - دوس تاني وتكلم';
   } else if(e.error==='audio-capture'){
    msg='❌ مفيش ميكروفون - تأكد الميكروفون شغال - أو اكتب طلبك يدوي';
   } else if(e.error==='network'){
    msg='❌ مشكلة شبكة - تأكد النت شغال - جرب تاني';
   }
   textDiv.innerHTML=msg+' - حاول مرة اخرى - أو اكتب يدوي واتساب';
   textDiv.style.background='#fef2f2';
   textDiv.style.border='2px solid #e11d48';
   btn.innerHTML='🎤 اضغط وتكلم - جرب تاني';
   btn.style.background='#7c3aed';
   btn.style.animation='none';
   isRecording=false;
  };
  recognition.onend = function(){
   if(isRecording && !voiceTranscript){
    textDiv.innerHTML='⏹ وقف التسجيل - ما سمعتش - دوس تاني وتكلم بصوت أوضح';
   }
   btn.innerHTML='🎤 اضغط وتكلم - اطلب قطعتي';
   btn.style.background='#7c3aed';
   btn.style.animation='none';
   isRecording=false;
  };
  recognition.onspeechend = function(){
   // auto stop after speech ends
  };
 }catch(err){
  textDiv.innerHTML='❌ خطأ: '+err.message+' - جرب Chrome - أو اكتب طلبك واتساب';
  btn.innerHTML='🎤 اضغط وتكلم';
  btn.style.background='#7c3aed';
  isRecording=false;
 }
}
function sendVoice(){

 if(!voiceTranscript){ alert('اتكلم اولاً'); return; }
 let msg=`طلب صوتي - صناعية السليم\nالطلب: ${voiceTranscript}\n---\nتم عبر الموقع - اطلب بصوتك`;
 let enc=encodeURIComponent(msg);
 let wa=`https://api.whatsapp.com/send?phone={{main_phone}}&text=${enc}`;
 window.open(wa,'_blank');
 document.getElementById('voiceText').innerHTML+='<br>✅ تم الارسال واتساب';
}

// تقييم
async function sendReview(idx, branch){
 let comment=document.getElementById('comment-'+idx).value;
 if(!comment){alert('اكتب تعليقك');return;}
 let data={branch:branch, stars:5, name:'عميل', comment:comment};
 let res=await fetch('/api/review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
 let j=await res.json();
 if(j.ok){
  document.getElementById('msg-'+idx).innerHTML='تم - شكرا - سيتم فتح واتساب';
  if(j.wa_links && j.wa_links.length>0) setTimeout(()=>{window.open(j.wa_links[0],'_blank');},500);
 }
}
</script>
</body></html>
"""

@app.route("/")
def home():
    db=load()
    return render_template_string(HTML, branches=db["branches"], main_phone=db.get("main_phone",""), marquee_text=db.get("marquee_text",""), products=db.get("products",[]))

@app.route("/api/data")
def api_data(): return jsonify(load())

@app.route("/api/review", methods=["POST"])
def api_review():
    db=load()
    data=request.get_json() or {}
    review={
        "branch":data.get("branch",""),
        "stars":int(data.get("stars",5)),
        "name":data.get("name","عميل"),
        "comment":data.get("comment",""),
        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "timestamp": datetime.now().isoformat()
    }
    db.setdefault("reviews",[]).append(review)
    if len(db["reviews"])>200: db["reviews"]=db["reviews"][-200:]
    save(db)
    wa_links=[]
    for phone in db.get("notify_phones",[]) or [db.get("main_phone","")]:
        phone=normalize_phone(phone)
        if not phone: continue
        msg=f"⭐ تقييم جديد - {review['branch']}\n👤 {review['name']} - {review['stars']} نجوم\n💬 {review['comment']}\n📅 {review['date']}"
        enc=urllib.parse.quote(msg)
        wa_links.append(f"https://api.whatsapp.com/send?phone={phone}&text={enc}")
    return jsonify({"ok":True, "wa_links":wa_links})

@app.route("/staff-login-2024", methods=["GET","POST"])
@app.route("/admin", methods=["GET","POST"])
def admin():
    db=load()
    if request.method=="POST":
        u=request.form.get("u"); p=request.form.get("p")
        for user in db["users"]:
            if user["name"]==u and user["pass"]==p:
                session["user"]=u; return redirect("/admin/panel")
        return "خطأ"
    return '<form method="post" style="max-width:400px;margin:50px auto;padding:20px;border:1px solid #ddd;border-radius:12px"><h3>⚙️ دخول الموظفين</h3><input name="u" placeholder="اسم المستخدم" autocomplete="off" style="width:100%;padding:12px;margin:4px 0"><input name="p" type="password" placeholder="كلمة السر" autocomplete="new-password" style="width:100%;padding:10px;margin:4px 0"><button style="width:100%;padding:12px;background:#000;color:#fff;border:0;border-radius:8px">دخول</button></form>'

@app.route("/admin/panel", methods=["GET","POST"])
def panel():
    if "user" not in session: return redirect("/admin")
    db=load()
    if request.method=="POST":
        if "save_marquee" in request.form:
            db["marquee_text"]=request.form.get("marquee_text","").strip() or db["marquee_text"]
            save(db); return redirect("/admin/panel")
        if "save_phones" in request.form:
            main=normalize_phone(request.form.get("main_phone","")) or db.get("main_phone","")
            db["main_phone"]=main
            raw=request.form.get("notify_phones_text","")
            phones=[]
            for line in raw.replace(',', '\n').splitlines():
                for p in line.split(','):
                    pp=normalize_phone(p.strip())
                    if pp and pp not in phones: phones.append(pp)
            if main and main not in phones: phones=[main]+phones
            elif main in phones: phones.remove(main); phones=[main]+phones
            db["notify_phones"]=phones or [main]
            save(db); return redirect("/admin/panel")
        if "save_ad" in request.form:
            title=request.form.get("ad_title","").strip()
            desc=request.form.get("ad_desc","").strip()
            if title:
                db.setdefault("products",[]).append({"title":title,"desc":desc,"date": datetime.now().strftime("%Y-%m-%d %H:%M")})
                save(db)
            return redirect("/admin/panel")
        if "delete_ad" in request.form:
            idx=int(request.form.get("delete_ad"))
            if 0<=idx<len(db.get("products",[])): db["products"].pop(idx); save(db)
            return redirect("/admin/panel")
        if "delete_review" in request.form:
            idx=int(request.form.get("delete_review"))
            if 0<=idx<len(db.get("reviews",[])): db["reviews"].pop(idx); save(db)
            return redirect("/admin/panel")
    notify_text="\n".join(db.get("notify_phones",[]))
    reviews_html=""
    for i,r in enumerate(reversed(db.get("reviews",[]))):
        real_idx=len(db.get("reviews",[]))-1-i
        reviews_html+=f"<div style='border:1px solid #ddd;padding:8px;margin:4px 0;border-radius:8px'><b>{r.get('branch')} - {r.get('stars')} stars</b> - {r.get('name')} - {r.get('date')}<br>{r.get('comment')}<form method='post' style='display:inline'><button name='delete_review' value='{real_idx}' style='background:#e11d48;color:#fff;border:0;padding:4px 8px;border-radius:4px'>حذف</button></form></div>"
    marquee_current=db.get("marquee_text","")
    return f"""
    <div style="padding:12px;font-family:Arial;max-width:900px;margin:auto" dir="rtl">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <h2>الادارة - لوحة التحكم</h2>
    <div style="background:#000;color:#fff;border:2px solid #000;border-radius:12px;padding:12px;margin:12px 0">
    <h3>🎯 الشريط المتحرك - مظبوط مش مايل</h3>
    <p style="font-size:13px">الصلاحية: تعديل الشريط - تقدر تغير الكلام</p>
    <form method="post">
    <textarea name="marquee_text" style="width:100%;height:80px;padding:12px;border-radius:8px;background:#111;color:#fff">{marquee_current}</textarea>
    <button name="save_marquee" value="1" style="width:100%;padding:12px;background:#fff;color:#000;border:0;border-radius:8px;font-weight:800;margin-top:8px">💾 حفظ الشريط</button>
    </form>
    </div>
    <div style="background:#f0fdf4;border:2px solid #22c55e;border-radius:12px;padding:12px;margin:12px 0">
    <h3>📱 ارقام استلام التقييمات</h3>
    <form method="post">
    <label>رقم خدمة العملاء الاساسي:</label>
    <input name="main_phone" value="{db.get('main_phone','')}" style="width:100%;padding:12px;margin:4px 0;border:2px solid #22c55e;border-radius:8px">
    <textarea name="notify_phones_text" style="width:100%;height:100px;padding:12px;margin:4px 0;border-radius:8px">{notify_text}</textarea>
    <button name="save_phones" value="1" style="width:100%;padding:12px;background:#22c55e;color:#fff;border:0;border-radius:8px;font-weight:800">💾 حفظ ارقام الواتساب</button>
    </form>
    </div>
    <div style="background:#fef2f2;border:2px solid #e11d48;border-radius:12px;padding:12px;margin:12px 0">
    <h3>📢 الإعلانات - وِلّ في الإعلانات - اعلانات قطع الغيار</h3>
    <p style="font-size:13px">الصلاحية: تعديل الإعلانات - اضف اعلان جديد يظهر للعميل في الموقع</p>
    <form method="post" style="background:#fff;padding:12px;border-radius:8px;margin:8px 0">
    <input name="ad_title" placeholder="عنوان الاعلان - مثال: خصم 20% على فلاتر الزيت" style="width:100%;padding:10px;margin:4px 0;border:1px solid #e11d48;border-radius:8px">
    <textarea name="ad_desc" placeholder="تفاصيل الاعلان - مثال: العرض ساري حتى نهاية الشهر - صناعية السليم ١ و ٢" style="width:100%;height:60px;padding:10px;margin:4px 0;border:1px solid #ddd;border-radius:8px"></textarea>
    <button name="save_ad" value="1" style="width:100%;padding:12px;background:#e11d48;color:#fff;border:0;border-radius:8px;font-weight:800">➕ اضف اعلان جديد</button>
    </form>
    <div style="margin-top:12px">
    {"".join([f"<div style='border:1px solid #ddd;padding:8px;margin:4px 0;border-radius:8px;background:#fff'><b>{p.get('title')}</b><br><small>{p.get('desc')} - {p.get('date')}</small><form method='post' style='display:inline;float:left'><button name='delete_ad' value='{i}' style='background:#e11d48;color:#fff;border:0;padding:4px 8px;border-radius:4px'>حذف</button></form></div>" for i,p in enumerate(db.get("products",[]))]) or "<p>لا يوجد اعلانات - اضف اول اعلان</p>"}
    </div>
    </div>
    
    <div style="background:#fff;border:2px solid #000;border-radius:12px;padding:12px;margin:12px 0">
    <h3>📍 الفروع - ٨ فروع بس - اتفقنا ٨</h3>
    <p style="font-size:12px;color:#666">اقصى ٨ فروع - تعديل من اكسيل branches.xlsx - عمود المدينة يحسب الآذان - رقم الجوال Text</p>
    <div style="max-height:300px;overflow-y:auto;border:1px solid #ddd;border-radius:8px;padding:8px">
    {"".join([f"<div style='border-bottom:1px solid #eee;padding:6px;display:flex;justify-content:space-between'><span>{b.get('name')} - {b.get('city')} - {b.get('phone')}</span><span style='font-size:11px;background:#000;color:#fff;padding:2px 8px;border-radius:999px'>{(b.get('hours1') or '')[:20]}</span></div>" for b in db.get("branches",[])])}
    </div>
    <p style="font-size:11px">لتعديل الفروع: عدل ملف branches.xlsx وارفعه - ٨ فروع بس - لما تكتب بعد العصر يظهر تمام - Text</p>
    </div>
    <div style="background:#f0f0ff;border:2px solid #6366f1;border-radius:12px;padding:12px;margin:12px 0">
    <h3>👥 المستخدمين - الصلاحيات</h3>
    <p style="font-size:12px">مصطفى - 1234 - كل الصلاحيات - لا تظهر جاهزة - لازم كتابة</p>
    <div>{"".join([f"<div style='border:1px solid #ddd;padding:6px;margin:4px 0;border-radius:6px'><b>{u.get('name')}</b> - {','.join(u.get('perms',[]))}</div>" for u in db.get("users",[])])}</div>
    </div>

    <div style="background:#fffbeb;border:1px solid #f59e0b;border-radius:12px;padding:12px;margin:12px 0">
    <h3>⭐ التقييمات</h3>
    {reviews_html or '<p>لا يوجد</p>'}
    </div>
    <a href="/" style="padding:10px 18px;background:#000;color:#fff;border-radius:999px;text-decoration:none">رجوع للموقع</a>
    </div>
    """

if __name__=="__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
