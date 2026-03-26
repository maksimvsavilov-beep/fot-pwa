from fastapi import FastAPI, UploadFile, File, HTTPException, Depends, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import sqlite3, hashlib, os, io, json
from datetime import datetime, timedelta
import pandas as pd
import secrets

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
app.mount("/static", StaticFiles(directory="static"), name="static")

DB = "fot.db"

STORE_DATA = {
    15159: {"name": "Пушкино Парк",    "plan": 7149435, "sr": 60000, "ar": 70000, "dr": 100000, "dc": 1, "ac": 2, "sc": 3, "nc": 0,   "cl": 30000, "ld": 6000},
    15003: {"name": "Чехов Карнавал",  "plan": 3989304, "sr": 55000, "ar": 65000, "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0.5, "cl": 15000, "ld": 7000},
    15023: {"name": "Раменское",       "plan": 3876988, "sr": 55000, "ar": 65000, "dr": 95000,  "dc": 1, "ac": 2, "sc": 2, "nc": 0.5, "cl": 18000, "ld": 11250},
    15171: {"name": "Видное Галерея",  "plan": 3222628, "sr": 60000, "ar": 65000, "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 27000, "ld": 9000},
    15141: {"name": "Жуковский",       "plan": 3017359, "sr": 60000, "ar": 60000, "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 15000, "ld": 9000},
    15147: {"name": "Янтарь",          "plan": 2626992, "sr": 60000, "ar": 70000, "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 27000, "ld": 9000},
    15222: {"name": "ТЦ Круг",         "plan": 2596560, "sr": 60000, "ar": 70000, "dr": 95000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 30000, "ld": 9000},
    15145: {"name": "Серпухов Атлас",  "plan": 2276176, "sr": 50000, "ar": 60000, "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 25000, "ld": 9000},
    15170: {"name": "Коломна КАДО",    "plan": 2214011, "sr": 50000, "ar": 65000, "dr": 85000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 30000, "ld": 9000},
    15191: {"name": "ТЦ Облака",       "plan": 2004742, "sr": 60000, "ar": 70000, "dr": 95000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 29000, "ld": 9000},
    15203: {"name": "Ивантеевка Твид", "plan": 1542018, "sr": 60000, "ar": 75000, "dr": 90000,  "dc": 1, "ac": 0, "sc": 2, "nc": 0,   "cl": 0,     "ld": 0},
    15221: {"name": "Кузьминки Молл",  "plan": 1512887, "sr": 60000, "ar": 75000, "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 29000, "ld": 9000},
}


def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            pin_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'director',
            store_id INTEGER,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL,
            expires_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS motivation_data (
            id INTEGER PRIMARY KEY,
            upload_date TEXT NOT NULL,
            file_date TEXT NOT NULL,
            store_id INTEGER NOT NULL,
            data_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS summary_data (
            id INTEGER PRIMARY KEY,
            upload_date TEXT NOT NULL,
            file_date TEXT NOT NULL,
            data_json TEXT NOT NULL
        );
    """)
    # Создаём дефолтного админа если нет
    existing = c.execute("SELECT id FROM users WHERE role='admin'").fetchone()
    if not existing:
        pin_hash = hashlib.sha256("admin1234".encode()).hexdigest()
        c.execute("INSERT INTO users (username, pin_hash, role) VALUES (?, ?, 'admin')",
                  ("admin", pin_hash))
    conn.commit()
    conn.close()


def hash_pin(pin: str) -> str:
    return hashlib.sha256(pin.encode()).hexdigest()


def create_session(user_id: int) -> str:
    token = secrets.token_urlsafe(32)
    expires = (datetime.now() + timedelta(days=30)).isoformat()
    conn = get_db()
    conn.execute("INSERT INTO sessions (token, user_id, expires_at) VALUES (?, ?, ?)",
                 (token, user_id, expires))
    conn.commit()
    conn.close()
    return token


def get_user_by_token(token: str):
    conn = get_db()
    row = conn.execute("""
        SELECT u.* FROM users u
        JOIN sessions s ON s.user_id = u.id
        WHERE s.token = ? AND s.expires_at > datetime('now')
    """, (token,)).fetchone()
    conn.close()
    return dict(row) if row else None


def plan_coef(pct: float) -> float:
    if pct < 80:  return 0.70
    if pct < 90:  return 0.80
    if pct < 95:  return 0.90
    if pct < 100: return 0.95
    if pct <= 105: return 1.00
    if pct <= 110: return 1.05
    if pct <= 120: return 1.10
    return 1.15


def parse_motivation_xlsx(content: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(content), header=None)
    header_row = None
    for i, row in df.iterrows():
        if str(row.iloc[0]).strip() == "Дата отчета":
            header_row = i
            break
    if header_row is None:
        raise ValueError("Не найден заголовок")

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = ["date", "division", "store_id", "login", "tab_num",
                  "role", "name", "to_val", "pct_to", "income", "fdm",
                  "avg_chek", "shvch", "rassr", "bezshk_share", "fdm_dist"]

    file_date = str(df["date"].iloc[0]).split(" ")[0] if len(df) > 0 else datetime.now().strftime("%d.%m.%Y")

    stores = {}
    for _, row in df.iterrows():
        try:
            sid = int(float(str(row["store_id"])))
        except:
            continue
        if sid not in STORE_DATA:
            continue
        if sid not in stores:
            stores[sid] = {"total": None, "staff": []}

        login = str(row["login"]).strip()
        role = str(row["role"]).strip()
        name = str(row["name"]).strip()

        try:
            to_val = float(row["to_val"]) if pd.notna(row["to_val"]) else 0
            income = float(row["income"]) if pd.notna(row["income"]) else 0
            fdm = float(row["fdm"]) if pd.notna(row["fdm"]) else 0
        except:
            to_val, income, fdm = 0, 0, 0

        if login == "Total":
            stores[sid]["total"] = {"to": to_val, "income": income, "fdm": fdm}
        elif login not in ("БезШК",) and role and role != "НетДолжности":
            stores[sid]["staff"].append({"name": name, "role": role, "income": income})

    # Считаем прогноз для каждого магазина
    results = {}
    for sid, s in stores.items():
        if not s["total"]:
            continue
        ref = STORE_DATA[sid]
        fact_to = s["total"]["to"]
        fact_fot = s["total"]["income"] + s["total"]["fdm"]

        results[sid] = {
            "store_id": sid,
            "store_name": ref["name"],
            "file_date": file_date,
            "fact_to": fact_to,
            "fact_fot": fact_fot,
            "plan": ref["plan"],
            "fact_plan_pct": round(fact_to / ref["plan"] * 100, 1) if ref["plan"] else 0,
            "staff": s["staff"],
            "ref": ref,
        }

    return {"file_date": file_date, "stores": results}


# ─── ROUTES ────────────────────────────────────────────────────────────────

@app.on_event("startup")
def startup():
    init_db()


@app.get("/")
def index():
    return FileResponse("static/index.html")


@app.get("/admin")
def admin():
    return FileResponse("static/admin.html")


# AUTH
@app.post("/api/login")
async def login(username: str = Form(...), pin: str = Form(...)):
    conn = get_db()
    user = conn.execute(
        "SELECT * FROM users WHERE username=? AND pin_hash=?",
        (username.strip(), hash_pin(pin.strip()))
    ).fetchone()
    conn.close()
    if not user:
        raise HTTPException(status_code=401, detail="Неверный логин или пин-код")
    token = create_session(user["id"])
    return {"token": token, "role": user["role"], "store_id": user["store_id"], "username": user["username"]}


@app.post("/api/logout")
async def logout(token: str = Form(...)):
    conn = get_db()
    conn.execute("DELETE FROM sessions WHERE token=?", (token,))
    conn.commit()
    conn.close()
    return {"ok": True}


# ADMIN: управление пользователями
@app.get("/api/admin/users")
def get_users(token: str):
    user = get_user_by_token(token)
    if not user or user["role"] != "admin":
        raise HTTPException(403)
    conn = get_db()
    users = conn.execute("SELECT id, username, role, store_id FROM users").fetchall()
    conn.close()
    return [dict(u) for u in users]


@app.post("/api/admin/users")
async def create_user(token: str = Form(...), username: str = Form(...),
                       pin: str = Form(...), role: str = Form(...),
                       store_id: str = Form(None)):
    user = get_user_by_token(token)
    if not user or user["role"] != "admin":
        raise HTTPException(403)
    sid = int(store_id) if store_id and store_id.strip() else None
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO users (username, pin_hash, role, store_id) VALUES (?,?,?,?)",
            (username.strip(), hash_pin(pin.strip()), role, sid)
        )
        conn.commit()
    except sqlite3.IntegrityError:
        raise HTTPException(400, "Такой логин уже существует")
    finally:
        conn.close()
    return {"ok": True}


@app.delete("/api/admin/users/{user_id}")
def delete_user(user_id: int, token: str):
    user = get_user_by_token(token)
    if not user or user["role"] != "admin":
        raise HTTPException(403)
    conn = get_db()
    conn.execute("DELETE FROM users WHERE id=? AND role!='admin'", (user_id,))
    conn.commit()
    conn.close()
    return {"ok": True}


# ADMIN: загрузка файла мотивации
@app.post("/api/admin/upload")
async def upload_file(token: str = Form(...), file: UploadFile = File(...)):
    user = get_user_by_token(token)
    if not user or user["role"] != "admin":
        raise HTTPException(403)

    content = await file.read()
    parsed = parse_motivation_xlsx(content)
    file_date = parsed["file_date"]
    upload_date = datetime.now().isoformat()

    conn = get_db()
    # Удаляем старые данные за ту же дату
    conn.execute("DELETE FROM motivation_data WHERE file_date=?", (file_date,))
    conn.execute("DELETE FROM summary_data WHERE file_date=?", (file_date,))

    for sid, store_data in parsed["stores"].items():
        conn.execute(
            "INSERT INTO motivation_data (upload_date, file_date, store_id, data_json) VALUES (?,?,?,?)",
            (upload_date, file_date, sid, json.dumps(store_data, ensure_ascii=False))
        )

    # Сводка по всем магазинам
    conn.execute(
        "INSERT INTO summary_data (upload_date, file_date, data_json) VALUES (?,?,?)",
        (upload_date, file_date, json.dumps(parsed["stores"], ensure_ascii=False, default=str))
    )
    conn.commit()
    conn.close()
    return {"ok": True, "file_date": file_date, "stores_count": len(parsed["stores"])}


@app.get("/api/admin/uploads")
def get_uploads(token: str):
    user = get_user_by_token(token)
    if not user or user["role"] != "admin":
        raise HTTPException(403)
    conn = get_db()
    rows = conn.execute(
        "SELECT DISTINCT file_date, upload_date FROM summary_data ORDER BY file_date DESC LIMIT 30"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# DATA: получение данных (с фильтрацией по роли)
@app.get("/api/data")
def get_data(token: str, days_total: int = 31, day_report: int = None,
             forecast_pct: float = 100, mode: str = "avg"):
    user = get_user_by_token(token)
    if not user:
        raise HTTPException(401)

    conn = get_db()
    # Берём последние загруженные данные
    latest = conn.execute(
        "SELECT file_date FROM summary_data ORDER BY upload_date DESC LIMIT 1"
    ).fetchone()

    if not latest:
        conn.close()
        return {"error": "Нет загруженных данных", "stores": {}}

    file_date = latest["file_date"]

    if user["role"] == "admin":
        rows = conn.execute(
            "SELECT store_id, data_json FROM motivation_data WHERE file_date=?", (file_date,)
        ).fetchall()
    else:
        if not user["store_id"]:
            conn.close()
            return {"error": "Магазин не назначен", "stores": {}}
        rows = conn.execute(
            "SELECT store_id, data_json FROM motivation_data WHERE file_date=? AND store_id=?",
            (file_date, user["store_id"])
        ).fetchall()

    conn.close()

    # Авто-определяем день отчёта из даты файла
    if day_report is None:
        try:
            day_report = int(file_date.split(".")[0])
        except:
            day_report = 24

    ratio = days_total / day_report if day_report > 0 else 1
    result = {}

    for row in rows:
        sd = json.loads(row["data_json"])
        sid = row["store_id"]
        ref = STORE_DATA.get(sid, {})
        if not ref:
            continue

        fact_to = sd["fact_to"]
        fact_fot = sd["fact_fot"]
        plan = ref["plan"]

        forecast_to = fact_to * ratio if mode == "avg" else plan * forecast_pct / 100
        plan_pct = forecast_to / plan * 100 if plan else 0
        coef = plan_coef(plan_pct)

        fot_dm = ref["dr"] * ref["dc"] * coef
        fot_am = ref["ar"] * ref["ac"] * coef
        fot_s = ref["sr"] * (ref["sc"] + ref["nc"]) * coef
        fot_wages = fot_dm + fot_am + fot_s
        extra = ref["cl"] + ref["ld"]
        forecast_fot = fot_wages + extra
        rest_fot = forecast_fot - fact_fot - extra * (day_report / days_total)

        # Прогноз ЗП по каждому сотруднику
        staff_forecast = []
        for emp in sd.get("staff", []):
            role = emp["role"]
            if "Директор" in role:
                rate = ref["dr"]
            elif "Администратор" in role:
                rate = ref["ar"]
            else:
                rate = ref["sr"]
            staff_forecast.append({
                "name": emp["name"],
                "role": role,
                "rate": rate,
                "forecast_salary": round(rate * coef),
                "fact_income": round(emp["income"]),
            })

        result[sid] = {
            "store_id": sid,
            "store_name": ref["name"],
            "file_date": file_date,
            "fact_to": round(fact_to),
            "fact_fot": round(fact_fot),
            "forecast_to": round(forecast_to),
            "forecast_fot": round(forecast_fot),
            "plan": plan,
            "fact_plan_pct": round(fact_to / plan * 100, 1) if plan else 0,
            "forecast_plan_pct": round(plan_pct, 1),
            "coef": coef,
            "fot_dm": round(fot_dm),
            "fot_am": round(fot_am),
            "fot_s": round(fot_s),
            "fot_wages": round(fot_wages),
            "extra": round(extra),
            "rest_fot": round(max(0, rest_fot)),
            "fot_eff": round(forecast_fot / forecast_to * 100, 1) if forecast_to else 0,
            "staff": staff_forecast,
        }

    return {
        "file_date": file_date,
        "day_report": day_report,
        "days_total": days_total,
        "stores": result,
    }


@app.get("/api/stores")
def get_stores():
    return [{"id": k, "name": v["name"]} for k, v in STORE_DATA.items()]


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", 8000)))
