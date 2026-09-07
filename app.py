
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import sqlite3, os, io
from datetime import datetime
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "smart-opsh-phone-change-me")
DB = os.path.join(os.path.dirname(__file__), "smart_opsh_phone.db")

DEV_PASSWORD = "moaazGamal111"
USER_PASSWORD = "1234"

def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS repairs(
      id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE, customer TEXT, phone TEXT,
      device TEXT, part TEXT, price REAL DEFAULT 0, fault TEXT, status TEXT DEFAULT 'جاري الصيانة',
      created_at TEXT, updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS inventory(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, category TEXT, qty INTEGER DEFAULT 0,
      buy_price REAL DEFAULT 0, sell_price REAL DEFAULT 0, min_qty INTEGER DEFAULT 3
    );
    CREATE TABLE IF NOT EXISTS sales(
      id INTEGER PRIMARY KEY AUTOINCREMENT, invoice TEXT, customer TEXT, total REAL DEFAULT 0,
      created_at TEXT
    );
    CREATE TABLE IF NOT EXISTS sale_items(
      id INTEGER PRIMARY KEY AUTOINCREMENT, sale_id INTEGER, name TEXT, qty INTEGER, price REAL
    );
    CREATE TABLE IF NOT EXISTS shifts(
      id INTEGER PRIMARY KEY AUTOINCREMENT, shift_date TEXT, opening REAL DEFAULT 0,
      closing REAL DEFAULT 0, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS used_phones(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, condition TEXT, price REAL, details TEXT
    );
    CREATE TABLE IF NOT EXISTS software(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, price REAL
    );
    """)
    c.commit()
    c.close()

init_db()

def login_required():
    return "user" in session

@app.route("/", methods=["GET"])
def home():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    stats={
      "repairs": c.execute("SELECT COUNT(*) n FROM repairs").fetchone()["n"],
      "today": c.execute("SELECT COUNT(*) n FROM repairs WHERE substr(created_at,1,10)=?",(datetime.now().strftime("%Y-%m-%d"),)).fetchone()["n"],
      "inventory": c.execute("SELECT COUNT(*) n FROM inventory").fetchone()["n"],
      "low": c.execute("SELECT COUNT(*) n FROM inventory WHERE qty < min_qty").fetchone()["n"],
      "sales": c.execute("SELECT COALESCE(SUM(total),0) n FROM sales WHERE substr(created_at,1,10)=?",(datetime.now().strftime("%Y-%m-%d"),)).fetchone()["n"]
    }
    c.close()
    return render_template("index.html", stats=stats, user=session["user"])

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method=="POST":
        p=request.form.get("password","")
        if p==DEV_PASSWORD: session["user"]="developer"
        elif p==USER_PASSWORD: session["user"]="user"
        else:
            flash("كلمة المرور غير صحيحة. للدعم الفني عبر واتساب: 01050122542","danger")
            return redirect(url_for("login"))
        return redirect(url_for("home"))
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login"))

@app.route("/repairs", methods=["GET","POST"])
def repairs():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    if request.method=="POST":
        code="SP-"+datetime.now().strftime("%y%m%d%H%M%S")
        c.execute("""INSERT INTO repairs(code,customer,phone,device,part,price,fault,status,created_at,updated_at)
          VALUES(?,?,?,?,?,?,?,?,?,?)""",(code,request.form["customer"],request.form["phone"],request.form["device"],
          request.form.get("part",""),float(request.form.get("price") or 0),request.form.get("fault",""),
          request.form.get("status","جاري الصيانة"),datetime.now().isoformat(timespec="seconds"),datetime.now().isoformat(timespec="seconds")))
        c.commit(); flash(f"تم تسجيل الجهاز. كود المتابعة: {code}","success")
        return redirect(url_for("repairs"))
    rows=c.execute("SELECT * FROM repairs ORDER BY id DESC").fetchall()
    parts=c.execute("SELECT * FROM inventory ORDER BY name").fetchall()
    c.close()
    return render_template("repairs.html", rows=rows, parts=parts)

@app.route("/repair/<int:rid>/status", methods=["POST"])
def repair_status(rid):
    if not login_required(): return redirect(url_for("login"))
    c=db(); c.execute("UPDATE repairs SET status=?,updated_at=? WHERE id=?",(request.form["status"],datetime.now().isoformat(timespec="seconds"),rid)); c.commit(); c.close()
    return redirect(url_for("repairs"))

@app.route("/inventory", methods=["GET","POST"])
def inventory():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO inventory(name,category,qty,buy_price,sell_price,min_qty) VALUES(?,?,?,?,?,?)",
          (request.form["name"],request.form["category"],int(request.form["qty"] or 0),float(request.form["buy_price"] or 0),float(request.form["sell_price"] or 0),int(request.form.get("min_qty") or 3)))
        c.commit(); flash("تم حفظ الصنف","success"); return redirect(url_for("inventory"))
    rows=c.execute("SELECT * FROM inventory ORDER BY category,name").fetchall(); c.close()
    return render_template("inventory.html", rows=rows, developer=session["user"]=="developer")

@app.route("/inventory/<int:iid>/delete", methods=["POST"])
def inventory_delete(iid):
    if session.get("user")!="developer": return redirect(url_for("inventory"))
    c=db(); c.execute("DELETE FROM inventory WHERE id=?",(iid,)); c.commit(); c.close(); return redirect(url_for("inventory"))

@app.route("/sales", methods=["GET","POST"])
def sales():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    if request.method=="POST":
        ids=request.form.getlist("item_id"); qtys=request.form.getlist("qty")
        customer=request.form.get("customer","عميل نقدي")
        items=[]; total=0
        for iid,q in zip(ids,qtys):
            q=int(q or 0)
            if q<=0: continue
            it=c.execute("SELECT * FROM inventory WHERE id=?",(iid,)).fetchone()
            if not it or it["qty"]<q: flash(f"المخزون غير كافٍ للصنف: {it['name'] if it else iid}","danger"); c.close(); return redirect(url_for("sales"))
            total += it["sell_price"]*q; items.append((it,q))
        inv="INV-"+datetime.now().strftime("%Y%m%d%H%M%S")
        cur=c.execute("INSERT INTO sales(invoice,customer,total,created_at) VALUES(?,?,?,?)",(inv,customer,total,datetime.now().isoformat(timespec="seconds")))
        sid=cur.lastrowid
        for it,q in items:
            c.execute("INSERT INTO sale_items(sale_id,name,qty,price) VALUES(?,?,?,?)",(sid,it["name"],q,it["sell_price"]))
            c.execute("UPDATE inventory SET qty=qty-? WHERE id=?",(q,it["id"]))
        c.commit(); c.close()
        return redirect(url_for("invoice", sid=sid))
    rows=c.execute("SELECT * FROM inventory WHERE qty>0 ORDER BY name").fetchall()
    recent=c.execute("SELECT * FROM sales ORDER BY id DESC LIMIT 10").fetchall(); c.close()
    return render_template("sales.html", rows=rows, recent=recent)

@app.route("/invoice/<int:sid>")
def invoice(sid):
    if not login_required(): return redirect(url_for("login"))
    c=db(); sale=c.execute("SELECT * FROM sales WHERE id=?",(sid,)).fetchone(); items=c.execute("SELECT * FROM sale_items WHERE sale_id=?",(sid,)).fetchall(); c.close()
    return render_template("invoice.html",sale=sale,items=items)

@app.route("/shifts", methods=["GET","POST"])
def shifts():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO shifts(shift_date,opening,closing,notes) VALUES(?,?,?,?)",(request.form["shift_date"],float(request.form["opening"] or 0),float(request.form["closing"] or 0),request.form.get("notes","")))
        c.commit(); return redirect(url_for("shifts"))
    rows=c.execute("SELECT * FROM shifts ORDER BY id DESC").fetchall(); c.close()
    return render_template("shifts.html",rows=rows)

@app.route("/used", methods=["GET","POST"])
def used():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO used_phones(name,condition,price,details) VALUES(?,?,?,?)",(request.form["name"],request.form["condition"],float(request.form["price"] or 0),request.form.get("details",""))); c.commit(); return redirect(url_for("used"))
    rows=c.execute("SELECT * FROM used_phones ORDER BY id DESC").fetchall(); c.close(); return render_template("used.html",rows=rows)

@app.route("/software", methods=["GET","POST"])
def software():
    if not login_required(): return redirect(url_for("login"))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO software(name,price) VALUES(?,?)",(request.form["name"],float(request.form["price"] or 0))); c.commit(); return redirect(url_for("software"))
    rows=c.execute("SELECT * FROM software ORDER BY name").fetchall(); c.close(); return render_template("software.html",rows=rows)

@app.route("/export")
def export():
    if not login_required(): return redirect(url_for("login"))
    c=db(); wb=Workbook(); wb.remove(wb.active)
    for title,sql in [("الصيانة","SELECT * FROM repairs"),("المخزون","SELECT * FROM inventory"),("المبيعات","SELECT * FROM sales"),("الشفتات","SELECT * FROM shifts"),("المستعمل","SELECT * FROM used_phones"),("السوفت وير","SELECT * FROM software")]:
        ws=wb.create_sheet(title); rows=c.execute(sql).fetchall()
        if rows:
            ws.append(rows[0].keys())
            for r in rows: ws.append(list(r))
    c.close(); bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio,as_attachment=True,download_name="Smart_Opsh_Phone_Backup.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)))
