import os, json, sqlite3, threading
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from flask import Flask, request, jsonify, render_template, send_file
from openpyxl import Workbook

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup, Update, WebAppInfo
from telegram.constants import ParseMode
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

# ================== ENV SETTINGS ==================
BOT_TOKEN = os.getenv("BOT_TOKEN","").strip()
WEBAPP_URL = os.getenv("WEBAPP_URL","").strip()  # https://<service>.onrender.com/
ADMINS = [int(x) for x in os.getenv("ADMINS","").split(",") if x.strip().isdigit()]
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD","1234").strip()
AUTO_APPROVE_MINUTES = int(os.getenv("AUTO_APPROVE_MINUTES","30"))

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "bookings.sqlite")
XLSX_PATH = os.path.join(DATA_DIR, "bookings.xlsx")

ROOMS = {
    "GO_3":  "Переговорная ГО (3 этаж)",
    "MINOR": "Кабинет офис Минор",
}

app = Flask(__name__)

# ================== DB ==================
def db():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con = db()
    con.execute("""
    CREATE TABLE IF NOT EXISTS bookings(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      created_at TEXT NOT NULL,
      user_id INTEGER,
      chat_id INTEGER,
      username TEXT,
      full_name TEXT,
      phone TEXT,
      room_id TEXT NOT NULL,
      room_name TEXT NOT NULL,
      date TEXT NOT NULL,
      start_time TEXT NOT NULL,
      end_time TEXT NOT NULL,
      purpose TEXT,
      participants TEXT,
      status TEXT NOT NULL
    )
    """)
    con.commit()
    con.close()

def tmin(t: str) -> int:
    h,m = t.split(":")
    return int(h)*60+int(m)

def overlaps(a1,a2,b1,b2): return a1 < b2 and b1 < a2

def has_conflict(room_id: str, date_s: str, start_t: str, end_t: str, ignore_id: Optional[int]=None) -> bool:
    s0,e0 = tmin(start_t), tmin(end_t)
    con = db()
    rows = con.execute(
        "SELECT id,start_time,end_time FROM bookings WHERE room_id=? AND date=? AND status IN ('pending','approved')",
        (room_id, date_s)
    ).fetchall()
    con.close()
    for r in rows:
        if ignore_id and int(r["id"]) == int(ignore_id): 
            continue
        if overlaps(s0,e0,tmin(r["start_time"]), tmin(r["end_time"])):
            return True
    return False

def list_bookings() -> List[Dict[str,Any]]:
    con = db()
    rows = con.execute("SELECT * FROM bookings ORDER BY date, start_time, id").fetchall()
    con.close()
    return [dict(r) for r in rows]

def get_booking(bid: int) -> Optional[Dict[str,Any]]:
    con = db()
    r = con.execute("SELECT * FROM bookings WHERE id=?", (int(bid),)).fetchone()
    con.close()
    return dict(r) if r else None

def export_excel():
    rows = list_bookings()
    wb = Workbook()
    ws = wb.active
    ws.title = "Bookings"
    ws.append(["id","created_at","status","room_id","room_name","date","start_time","end_time","purpose","participants","user_id","username","full_name","phone","chat_id"])
    for r in rows:
        ws.append([
            r.get("id",""), r.get("created_at",""), r.get("status",""),
            r.get("room_id",""), r.get("room_name",""),
            r.get("date",""), r.get("start_time",""), r.get("end_time",""),
            r.get("purpose",""), r.get("participants",""),
            r.get("user_id",""), r.get("username",""), r.get("full_name",""),
            r.get("phone",""), r.get("chat_id","")
        ])
    wb.save(XLSX_PATH)

def create_booking(payload: Dict[str,Any]) -> Tuple[int, Dict[str,Any]]:
    for k in ("room_id","date","start_time","end_time","purpose"):
        if not payload.get(k):
            return 400, {"error": f"missing {k}"}
    room_id = str(payload["room_id"])
    if room_id not in ROOMS:
        return 400, {"error":"unknown room"}
    date_s = str(payload["date"])
    start_t = str(payload["start_time"])
    end_t = str(payload["end_time"])

    try:
        datetime.strptime(date_s, "%Y-%m-%d")
        if tmin(end_t) <= tmin(start_t):
            return 400, {"error":"invalid time range"}
    except Exception:
        return 400, {"error":"invalid date/time"}

    # запрет прошлых дат
    try:
        d = datetime.strptime(date_s, "%Y-%m-%d").date()
        if d < datetime.now().date():
            return 400, {"error":"past date not allowed"}
    except Exception:
        return 400, {"error":"invalid date"}

    if has_conflict(room_id, date_s, start_t, end_t):
        return 409, {"error":"conflict"}

    con = db()
    cur = con.execute("""
      INSERT INTO bookings(created_at,user_id,chat_id,username,full_name,phone,room_id,room_name,date,start_time,end_time,purpose,participants,status)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
      int(payload.get("user_id") or 0) or None,
      int(payload.get("chat_id") or 0) or None,
      str(payload.get("username") or ""),
      str(payload.get("full_name") or ""),
      str(payload.get("phone") or ""),
      room_id,
      ROOMS[room_id],
      date_s, start_t, end_t,
      str(payload.get("purpose") or ""),
      str(payload.get("participants") or ""),
      "pending"
    ))
    bid = cur.lastrowid
    con.commit()
    con.close()
    export_excel()
    b = get_booking(bid)
    if b:
        notify_admins_new_booking(b)
    return 200, b or {"error":"create failed"}

def set_status(bid: int, status: str) -> Tuple[int, Dict[str,Any]]:
    if status not in ("pending","approved","rejected"):
        return 400, {"error":"invalid status"}
    b = get_booking(bid)
    if not b:
        return 404, {"error":"not found"}
    if status == "approved":
        if has_conflict(b["room_id"], b["date"], b["start_time"], b["end_time"], ignore_id=bid):
            return 409, {"error":"conflict"}
    con = db()
    con.execute("UPDATE bookings SET status=? WHERE id=?", (status, int(bid)))
    con.commit()
    con.close()
    export_excel()
    b2 = get_booking(bid) or b
    notify_user_status_change(b2)
    return 200, b2

# ================== Flask routes ==================
@app.get("/")
def page():
    return render_template("index.html")

@app.get("/api/bookings")
def api_list():
    return jsonify(list_bookings())

@app.post("/api/bookings")
def api_create():
    payload = request.get_json(force=True, silent=True) or {}
    # if request comes from Telegram WebApp, initData isn't validated here; acceptable for internal tool
    return jsonify(*create_booking(payload)[1:]) if False else (lambda code_data: (jsonify(code_data[1]), code_data[0]))(create_booking(payload))

@app.post("/api/bookings/<int:bid>/status")
def api_status(bid: int):
    payload = request.get_json(force=True, silent=True) or {}
    if str(payload.get("admin_password","")) != ADMIN_PASSWORD:
        return jsonify({"error":"bad admin password"}), 403
    return (lambda code_data: (jsonify(code_data[1]), code_data[0]))(set_status(bid, str(payload.get("status",""))))

@app.get("/excel")
def excel():
    if not os.path.exists(XLSX_PATH):
        export_excel()
    return send_file(XLSX_PATH, as_attachment=True, download_name="bookings.xlsx")

# ================== Telegram bot (optional) ==================
BOT_APP: Optional[Application] = None

def kb_webapp():
    url = WEBAPP_URL.rstrip("/") or ""
    return ReplyKeyboardMarkup([[KeyboardButton("📅 Забронировать переговорку", web_app=WebAppInfo(url=url))]], resize_keyboard=True)

def kb_admin(bid: int):
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ ПОДТВЕРДИТЬ", callback_data=f"approve_{bid}"),
        InlineKeyboardButton("❌ ОТКЛОНИТЬ", callback_data=f"reject_{bid}"),
    ]])

def fmt_booking(b: Dict[str,Any]) -> str:
    return (
      f"🏢 {b.get('room_name')}\n"
      f"📅 {b.get('date')}  ⏰ {b.get('start_time')}–{b.get('end_time')}\n"
      f"👤 {b.get('full_name') or b.get('username') or b.get('user_id')}\n"
      f"👥 Участников: {b.get('participants','')}\n"
      f"📝 {b.get('purpose','')}"
    )

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.effective_message.reply_text("Нажмите кнопку для бронирования 👇", reply_markup=kb_webapp())

async def cb_admin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.from_user.id not in ADMINS:
        await q.edit_message_text("❌ Нет доступа")
        return
    action, bid_s = q.data.split("_",1)
    status = "approved" if action=="approve" else "rejected"
    code, data = set_status(int(bid_s), status)
    if code != 200:
        msg = "🚫 Конфликт" if data.get("error")=="conflict" else "❌ Ошибка"
        await q.edit_message_text(msg)
        return
    await q.edit_message_text(("✅ " if status=="approved" else "❌ ") + f"Заявка #{bid_s} обновлена")

def notify_admins_new_booking(b: Dict[str,Any]):
    if BOT_APP is None or not ADMINS:
        return
    async def _send():
        txt = "📋 *НОВАЯ ЗАЯВКА*\n\n" + fmt_booking(b)
        for aid in ADMINS:
            try:
                await BOT_APP.bot.send_message(aid, txt, parse_mode=ParseMode.MARKDOWN, reply_markup=kb_admin(int(b["id"])))
            except Exception:
                pass
    BOT_APP.create_task(_send())

def notify_user_status_change(b: Dict[str,Any]):
    if BOT_APP is None:
        return
    chat_id = b.get("chat_id")
    if not chat_id:
        return
    icon = "✅" if b.get("status")=="approved" else "❌"
    label = "ПОДТВЕРЖДЕНО" if b.get("status")=="approved" else "ОТКЛОНЕНО"
    async def _send():
        try:
            await BOT_APP.bot.send_message(int(chat_id), f"{icon} Ваша заявка #{b['id']} — {label}\n\n{fmt_booking(b)}")
        except Exception:
            pass
    BOT_APP.create_task(_send())

def run_bot():
    global BOT_APP
    if not BOT_TOKEN:
        return
    BOT_APP = Application.builder().token(BOT_TOKEN).build()
    BOT_APP.add_handler(CommandHandler("start", cmd_start))
    BOT_APP.add_handler(CallbackQueryHandler(cb_admin, pattern=r"^(approve|reject)_\d+$"))
    BOT_APP.run_polling(drop_pending_updates=True)

def main():
    init_db()
    export_excel()
    if BOT_TOKEN:
        t = threading.Thread(target=run_bot, daemon=True)
        t.start()
    port = int(os.getenv("PORT","8000"))
    app.run(host="0.0.0.0", port=port)

if __name__ == "__main__":
    main()
