
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
  {"name": "صناعية السليم ١", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966163274444", "loc": "https://share.google/KVhZpBRi4vOvGJcD", "img": "", "rating": 4.8, "lat": 26.359, "lng": 43.968},
  {"name": "صناعية السليم ٢", "city": "بريدة", "hours1": "من 8 صباحا إلى آذان الظهر", "hours2": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "friday": "من بعد آذان العصر بنصف ساعة إلى آذان العشاء", "phone": "966163274444", "loc": "https://maps.app.goo.gl/Qvbomn7VyQEhRNDA", "img": "", "rating": 4.8, "lat": 26.36, "lng": 43.97}
 ],
 "products": [],
 "users": [{"name": "مصطفى", "pass": "1234", "perms": ["branches","ads","hours","add","users","marquee","reviews"]}],
 "reviews": [],
 "notify_phones": ["966163274444"],
 "main_phone": "966163274444",
 "marquee_text": "صناعية السليم ١ • صناعية السليم ٢ • بريدة الرئيسي • دوام بعد آذان العصر بنصف ساعة إلى آذان العشاء • تقييمك يهمنا - شاركنا موقعك يجيبلك اقرب فرع"
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
.voice-btn{background:#7c3aed;color:#fff;padding:12px 18px;border-radius:999px;border:0;font-weight:800;font-size:16px;cursor:pointer}
.review-box{border:2px solid #22c55e;border-radius:12px;padding:12px;margin:8px 0;background:#f0fdf4}
.footer{background:#0a0a0a;color:#fff;text-align:center;padding:24px}
</style></head>
<body>
<header style="display:flex;justify-content:space-between;align-items:center;padding:8px 16px;background:#fff;border-bottom:2px solid #000;position:sticky;top:0;z-index:20"><div style="display:flex;align-items:center;gap:10px"><img src="/static/logo.jpg" style="width:58px;height:58px;border-radius:12px;border:2px solid #000;object-fit:cover;background:#fff" onerror="this.src='/static/logo.jpg'"><b style="font-size:16px;color:#000">شركة القرعاوي</b></div><div style="display:flex;gap:6px;align-items:center"><a class="btn btn-g" href="https://api.whatsapp.com/send?phone={{main_phone}}" style="padding:8px 14px;font-size:13px">واتساب</a><a class="btn btn-b" href="tel:+{{main_phone}}" style="padding:8px 14px;font-size:13px;background:#000;color:#fff;border:1px solid #000">📞</a></div></header>

<div style="background:#000;color:#fff;overflow:hidden;white-space:nowrap;padding:10px 0;font-size:15px;font-weight:900;line-height:1.5;border-bottom:2px solid #222"><marquee scrollamount="6" direction="right"> {{marquee_text}} </marquee></div>

<!-- 2: VIN Search + 6: Voice + 10: Share Location -->
<div class="section" style="background:#fefce8;border-bottom:2px solid #eab308">
<h2 style="margin:0 0 12px">🔍 ابحث بذكاء - ٣ طرق</h2>
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px">

<div class="search-box" style="border-color:#e11d48;background:#fef2f2">
<h3 style="margin:0 0 8px">٢- بحث برقم الشاصي VIN (17 رقم)</h3>
<p style="font-size:13px;margin:4px 0">اكتب رقم الشاصي - نطلع لك القطع المناسبة لسيارتك بالظبط</p>
<input id="vinInput" placeholder="JTDKB20U093456789 - 17 حرف انجليزي بالظبط - لا اكتر ولا اقل" maxlength="17" pattern="[A-Za-z0-9]{17}" title="17 حرف ورقم انجليزي بالظبط" style="width:100%;padding:12px;border-radius:8px;border:2px solid #e11d48;font-size:16px;direction:ltr;text-transform:uppercase;letter-spacing:2px;font-weight:800" oninput="let v=this.value.toUpperCase().replace(/[^A-Z0-9]/g,''); if(this.value!==v) this.value=v; document.getElementById('vinCount').innerHTML=this.value.length+'/17 - '+(this.value.length==17?'✅ مظبوط - أحرف وأرقام انجليزي':'❌ لازم 17 - '+(17-this.value.length)+' باقي')+ ' - '+(this.value!==this.value.toLowerCase()?'كبير ✅':'')" style="width:100%;padding:12px;border-radius:8px;border:1px solid #e11d48;font-size:16px">
<div id="vinCount" style="font-size:12px;margin-top:8px;color:#111;font-weight:700;background:#fff;padding:6px 10px;border-radius:6px;border:1px solid #eee;display:inline-block">0/17 - لازم 17 بالظبط - انجليزي فقط</div><button onclick="searchVIN()" style="width:100%;margin-top:8px;padding:12px;background:#e11d48;color:#fff;border:0;border-radius:8px;font-weight:800">🔍 ابحث بقطع الشاصي - يرسل واتساب</button>
<div id="vinResult" style="margin-top:8px;font-size:13px"></div>
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
function searchVIN(){
 let vin=document.getElementById('vinInput').value.trim().toUpperCase();
 let res=document.getElementById('vinResult');
 if(vin.length<10){ res.innerHTML='اكتب رقم الشاصي كامل - 17 حرف ورقم'; return; }
 if(vin.length!==17){ res.innerHTML='رقم الشاصي لازم 17 حرف ورقم - انت كاتب '+vin.length; }
 let msg=`طلب قطع برقم الشاصي VIN\nالشاصي: ${vin}\nالفرع: صناعية السليم\nارجو تزويدي بالقطع المناسبة لهذا الشاصي - تويوتا اصلية`;
 let enc=encodeURIComponent(msg);
 let wa=`https://api.whatsapp.com/send?phone={{main_phone}}&text=${enc}`;
 res.innerHTML=`✅ الشاصي: <b>${vin}</b><br>جاري تحويلك لواتساب خدمة العملاء - نطلع لك القطع المناسبة...<br><a href="${wa}" target="_blank" style="background:#e11d48;color:#fff;padding:10px 18px;border-radius:999px;text-decoration:none;display:inline-block;margin-top:8px">📤 ارسل الشاصي واتساب - {{main_phone}}</a>`;
 setTimeout(()=>{window.open(wa,'_blank');},800);
}

// 6- اطلب بصوتك
let recognition;
let voiceTranscript='';
function startVoice(){
 let btn=document.getElementById('voiceBtn');
 let textDiv=document.getElementById('voiceText');
 let sendBtn=document.getElementById('voiceSend');
 if(!('webkitSpeechRecognition' in window || 'SpeechRecognition' in window)){
  textDiv.innerHTML='المتصفح لا يدعم التعرف على الصوت - استخدم Chrome على الجوال - او اكتب طلبك واتساب';
  return;
 }
 let SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
 recognition = new SpeechRecognition();
 recognition.lang='ar-SA';
 recognition.interimResults=false;
 recognition.maxAlternatives=1;
 btn.innerHTML='🎙️ اسمعك... تكلم الآن';
 btn.style.background='#ef4444';
 textDiv.innerHTML='🎤 اسمعك... قول: عايز مساعد امامي كامري 2018';
 recognition.start();
 recognition.onresult = function(event){
  voiceTranscript = event.results[0][0].transcript;
  textDiv.innerHTML='✅ سمعتك: <b>'+voiceTranscript+'</b>';
  sendBtn.style.display='block';
  btn.innerHTML='🎤 اضغط وتكلم مرة اخرى';
  btn.style.background='#7c3aed';
 };
 recognition.onerror = function(e){
  textDiv.innerHTML='خطأ: '+e.error+' - حاول مرة اخرى - او اكتب يدوي';
  btn.innerHTML='🎤 اضغط وتكلم';
  btn.style.background='#7c3aed';
 };
 recognition.onend = function(){
  btn.innerHTML='🎤 اضغط وتكلم';
  btn.style.background='#7c3aed';
 };
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
    return '<form method="post" style="max-width:400px;margin:50px auto;padding:20px;border:1px solid #ddd;border-radius:12px"><h3>⚙️ دخول الموظفين</h3><input name="u" placeholder="اسم" style="width:100%;padding:10px;margin:4px 0"><input name="p" type="password" placeholder="كلمة السر" style="width:100%;padding:10px;margin:4px 0"><button style="width:100%;padding:12px;background:#000;color:#fff;border:0;border-radius:8px">دخول</button></form>'

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
    <div style="background:#fffbeb;border:1px solid #f59e0b;border-radius:12px;padding:12px;margin:12px 0">
    <h3>⭐ التقييمات</h3>
    {reviews_html or '<p>لا يوجد</p>'}
    </div>
    <a href="/" style="padding:10px 18px;background:#000;color:#fff;border-radius:999px;text-decoration:none">رجوع للموقع</a>
    </div>
    """

if __name__=="__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
