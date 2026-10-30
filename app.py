
from flask import Flask, render_template_string, request, jsonify, redirect, session
import json, os
app = Flask(__name__)
app.secret_key = "alqarawi123"
DB_FILE = "alqarawi_db.json"
DEFAULT = {
 "branches": [
  {"name": "فرع بريدة - السليمانية 1", "city": "بريدة", "hours": "السبت - الخميس 8ص - 10م", "phone": "966163274444", "loc": "https://maps.google.com/?q=بريدة"},
  {"name": "فرع الرياض - السلي", "city": "الرياض", "hours": "السبت - الخميس 8ص - 10م", "phone": "966163274444", "loc": "https://maps.google.com/?q=الرياض"},
  {"name": "فرع جدة", "city": "جدة", "hours": "السبت - الخميس 8ص - 10م", "phone": "966163274444", "loc": "https://maps.google.com/?q=جدة"}
 ],
 "products": [
  {"name": "تويوتا كامري", "model": "Camry", "price": "249", "img": "https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=400", "desc": "فلتر وترمبات أصلية"},
  {"name": "تويوتا كورولا", "model": "Corolla", "price": "189", "img": "https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=400", "desc": "قطع أصلية"},
  {"name": "تويوتا هايلوكس", "model": "Hilux", "price": "329", "img": "https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=400", "desc": "تشغيلة 2025"},
  {"name": "تويوتا برادو", "model": "Prado", "price": "399", "img": "https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=400", "desc": "قطع أصلية برادو"},
  {"name": "تويوتا يارس", "model": "Yaris", "price": "159", "img": "https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=400", "desc": "اقتصادية وأصلية"},
  {"name": "تويوتا لاندكروزر", "model": "Land", "price": "489", "img": "https://images.unsplash.com/photo-1621007947382-bb3c3994e3fb?w=400", "desc": "أصلية لاندكروزر"}
 ],
 "users": [{"name": "مصطفى", "pass": "1234", "perms": ["branches","ads","hours","add","users"]}]
}
def load():
    if not os.path.exists(DB_FILE): return json.loads(json.dumps(DEFAULT))
    return json.load(open(DB_FILE, encoding="utf-8"))
def save(db): json.dump(db, open(DB_FILE,"w",encoding="utf-8"), ensure_ascii=False, indent=2)
HTML = open("templates.html",encoding="utf-8").read() if os.path.exists("templates.html") else """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>شركة القرعاوي</title>
<style>body{margin:0;font-family:Arial;background:#fff}header{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#fff;border-bottom:1px solid #eee}nav{display:flex;gap:16px;font-weight:700}.hero{background:#0a0a0a;color:#fff;text-align:center;padding:32px 16px}.btn{display:inline-block;padding:12px 20px;border-radius:999px;font-weight:800;margin:6px}.btn-g{background:#22c55e;color:#fff}.section{padding:24px 16px;max-width:1100px;margin:auto}.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px}.card{border:1px solid #eee;border-radius:16px;overflow:hidden}.card img{width:100%;height:160px;object-fit:cover}.card-body{padding:14px}.red{background:#e11d48;color:#fff}.footer{background:#0a0a0a;color:#fff;text-align:center;padding:24px}</style></head>
<body><header><b>شركة القرعاوي - Q</b><nav><a href="/">الرئيسية</a><a href="#branches">فروعنا</a><a href="#products">المنتجات</a><a href="#contact">تواصل معنا</a></nav><a href="/admin" class="btn" style="background:#000;color:#fff">الإدارة</a></header>
<div class="hero"><h1>شركة القرعاوي - ثقتك في قطع غيار تويوتا الأصلية</h1><p>شركة القرعاوي</p><p>فروعنا في جميع أنحاء المملكة</p><p>بأسعار تنافسية من خلال توفيرنا لقطع الغيار للعملاء على مدار الساعة</p><p>بأسعار تنافسية من خلال توفيرنا لقطع الغيار للعملاء على مدار الساعة</p><a class="btn btn-g" href="https://api.whatsapp.com/send?phone=966163274444">اتواصل معنا - 0163274444</a></div>
<div class="section"><h2>جديد القرعاوي - لا يفوتك</h2><div class="card" style="background:#e11d48;color:#fff"><div class="card-body"><h3>افتتاح فرع بريدة - السليمانية 1</h3></div></div></div>
<div class="section red" id="branches"><h2>فروعنا - اختر أقرب فرع لك</h2><div class="cards">{% for b in branches %}<div class="card"><div class="card-body"><h3>{{b.name}}</h3><p>{{b.city}} - {{b.hours}}</p><a class="btn btn-g" href="https://api.whatsapp.com/send?phone={{b.phone}}">واتساب</a> <a class="btn" style="background:#000;color:#fff" href="{{b.loc}}">📍 الموقع</a></div></div>{% endfor %}</div></div>
<div class="section" id="products"><h2>المنتجات</h2><div class="cards">{% for p in products %}<div class="card"><img src="{{p.img}}"><div class="card-body"><b>TOYOTA GENUINE</b><h3>{{p.name}}</h3><p>{{p.model}} - {{p.desc}}</p><b>ابتداءً من {{p.price}} ر.س</b><br><a class="btn btn-g" href="https://api.whatsapp.com/send?phone=966163274444">اطلب القطعة واتساب</a></div></div>{% endfor %}</div></div>
<div class="footer" id="contact"><h3>شركة القرعاوي لبيع قطع غيار تويوتا</h3><p>شركة القرعاوي</p><p>فروعنا في جميع أنحاء المملكة</p><a class="btn btn-g" href="https://api.whatsapp.com/send?phone=966163274444">اتواصل معنا</a></div>
<script>setInterval(()=>{fetch("/api/data").then(r=>r.json()).then(d=>{ if(JSON.stringify(d)!==localStorage.getItem("last")){ localStorage.setItem("last",JSON.stringify(d)); location.reload(); } })},10000);</script></body></html>
"""
@app.route("/")
def home():
    db=load()
    return render_template_string(HTML, branches=db["branches"], products=db["products"])
@app.route("/api/data")
def api_data(): return jsonify(load())
@app.route("/admin", methods=["GET","POST"])
def admin():
    db=load()
    if request.method=="POST":
        u=request.form.get("u"); p=request.form.get("p")
        for user in db["users"]:
            if user["name"]==u and user["pass"]==p:
                session["user"]=u; return redirect("/admin/panel")
        return "خطأ - الاسم أو الباسورد غلط"
    return """<form method="post"><h3>الإدارة</h3><input name="u" placeholder="اسم المستخدم"><input name="p" type="password" placeholder="كلمة السر"><button>دخول</button></form>"""
@app.route("/admin/panel", methods=["GET","POST"])
def panel():
    if "user" not in session: return redirect("/admin")
    db=load()
    if request.method=="POST":
        # تعديل فرع
        if "branch_idx" in request.form:
            i=int(request.form["branch_idx"]); db["branches"][i]["name"]=request.form["name"]; db["branches"][i]["hours"]=request.form["hours"]; save(db)
        if "prod_idx" in request.form:
            i=int(request.form["prod_idx"]); db["products"][i]["name"]=request.form["pname"]; db["products"][i]["price"]=request.form["price"]; db["products"][i]["img"]=request.form["img"]; save(db)
        if "new_user" in request.form:
            db["users"].append({"name": request.form["new_user"], "pass": request.form["new_pass"], "perms": request.form.getlist("perms")}); save(db)
        if "my_pass" in request.form:
            for u in db["users"]:
                if u["name"]==session["user"]: u["pass"]=request.form["my_pass"]
            save(db)
        return redirect("/admin/panel")
    opts="".join([f'<option value="{i}">{b["name"]}</option>' for i,b in enumerate(db["branches"])])
    popts="".join([f'<option value="{i}">{p["name"]}</option>' for i,p in enumerate(db["products"])])
    return f"""
    <h2>الإدارة - {session["user"]}</h2>
    <h3>تعديل فرع - قايمة منسدلة</h3><form method="post"><select name="branch_idx">{opts}</select><input name="name" placeholder="اسم الفرع"><input name="hours" placeholder="الدوام"><button>حفظ الفرع</button></form>
    <h3>تعديل منتج بالصورة</h3><form method="post"><select name="prod_idx">{popts}</select><input name="pname" placeholder="اسم المنتج"><input name="price" placeholder="السعر"><input name="img" placeholder="رابط صورة"><button>حفظ المنتج</button></form>
    <h3>إضافة مستخدم وصلاحيات</h3><form method="post"><input name="new_user" placeholder="اسم الجديد"><input name="new_pass" placeholder="باسورد"><label><input type="checkbox" name="perms" value="branches"> تعديل الفروع</label><label><input type="checkbox" name="perms" value="ads"> تعديل الإعلانات</label><label><input type="checkbox" name="perms" value="hours"> تعديل الدوام</label><label><input type="checkbox" name="perms" value="add"> إضافة فرع / إعلان</label><label><input type="checkbox" name="perms" value="users"> إدارة المستخدمين</label><button>إضافة</button></form>
    <h3>حسابي - تغيير كلمة سري</h3><form method="post"><input name="my_pass" placeholder="كلمة السر الجديدة"><button>تغيير</button></form>
    <p>✅ تعديلك هيسمع عند كل الناس فورا - من غير رفع ملف</p>
    <a href="/">رجوع للموقع</a>
    """
if __name__=="__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
