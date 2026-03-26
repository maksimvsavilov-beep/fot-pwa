from fastapi import FastAPI, HTTPException, Depends, UploadFile, File, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import sqlite3, hashlib, secrets, os, io
from datetime import datetime, timedelta
from typing import Optional
import pandas as pd

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_PATH = "fot.db"
ADMIN_PIN = os.environ.get("ADMIN_PIN", "admin1234")

# ─── Справочник магазинов ────────────────────────────────────────────────────
STORE_DATA = {
    15159: {"name": "Пушкино Парк",    "plan": 7149435, "sr": 60000, "ar": 70000,  "dr": 100000, "dc": 1, "ac": 2, "sc": 3, "nc": 0,   "cl": 30000, "ld": 6000},
    15003: {"name": "Чехов Карнавал",  "plan": 3989304, "sr": 55000, "ar": 65000,  "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0.5, "cl": 15000, "ld": 7000},
    15023: {"name": "Раменское",       "plan": 3876988, "sr": 55000, "ar": 65000,  "dr": 95000,  "dc": 1, "ac": 2, "sc": 2, "nc": 0.5, "cl": 18000, "ld": 11250},
    15171: {"name": "Видное Галерея",  "plan": 3222628, "sr": 60000, "ar": 65000,  "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 27000, "ld": 9000},
    15141: {"name": "Жуковский",       "plan": 3017359, "sr": 60000, "ar": 60000,  "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 15000, "ld": 9000},
    15147: {"name": "Янтарь",          "plan": 2626992, "sr": 60000, "ar": 70000,  "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 27000, "ld": 9000},
    15222: {"name": "ТЦ Круг",         "plan": 2596560, "sr": 60000, "ar": 70000,  "dr": 95000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 30000, "ld": 9000},
    15145: {"name": "Серпухов Атлас",  "plan": 2276176, "sr": 50000, "ar": 60000,  "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 25000, "ld": 9000},
    15170: {"name": "Коломна КАДО",    "plan": 2214011, "sr": 50000, "ar": 65000,  "dr": 85000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 30000, "ld": 9000},
    15191: {"name": "ТЦ Облака",       "plan": 2004742, "sr": 60000, "ar": 70000,  "dr": 95000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 29000, "ld": 9000},
    15203: {"name": "Ивантеевка Твид", "plan": 1542018, "sr": 60000, "ar": 75000,  "dr": 90000,  "dc": 1, "ac": 0, "sc": 2, "nc": 0,   "cl": 0,     "ld": 0},
    15221: {"name": "Кузьминки Молл",  "plan": 1512887, "sr": 60000, "ar": 75000,  "dr": 90000,  "dc": 1, "ac": 1, "sc": 2, "nc": 0,   "cl": 29000, "ld": 9000},
}

# ─── БД ─────────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            pin_hash TEXT NOT NULL,
            store_id INTEGER,
            role TEXT DEFAULT 'director',
            token TEXT,
            token_expires TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS motivation_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            store_id INTEGER NOT NULL,
            report_date TEXT NOT NULL,
            login TEXT,
            role TEXT,
            name TEXT,
            to_fact REAL DEFAULT 0,
            pct_to REAL DEFAULT 0,
            income REAL DEFAULT 0,
            fdm REAL DEFAULT 0,
            is_total INTEGER DEFAULT 0,
            is_bezshk INTEGER DEFAULT 0,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS upload_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            report_date TEXT,
            rows_count INTEGER,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
    """)
    conn.commit()
    # Создаём admin если нет
    cur = conn.execute("SELECT id FROM users WHERE role='admin'")
    if not cur.fetchone():
        pin_hash = hashlib.sha256(ADMIN_PIN.encode()).hexdigest()
        conn.execute("INSERT INTO users (name, pin_hash, store_id, role) VALUES (?, ?, NULL, 'admin')",
                     ("Администратор", pin_hash))
        conn.commit()
    conn.close()

init_db()

# ─── Утилиты ─────────────────────────────────────────────────────────────────
def hash_pin(pin: str) -> str:
    return hashlib.sha256(pin.strip().encode()).hexdigest()

def plan_coef(pct: float) -> float:
    if pct < 80:  return 0.70
    if pct < 90:  return 0.80
    if pct < 95:  return 0.90
    if pct < 100: return 0.95
    if pct <= 105: return 1.00
    if pct <= 110: return 1.05
    if pct <= 120: return 1.10
    return 1.15

def get_user_by_token(token: str):
    if not token:
        return None
    conn = get_db()
    user = conn.execute(
        "SELECT * FROM users WHERE token=? AND token_expires > ?",
        (token, datetime.now().isoformat())
    ).fetchone()
    conn.close()
    return user

def require_auth(authorization: Optional[str] = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Не авторизован")
    token = authorization.split(" ")[1]
    user = get_user_by_token(token)
    if not user:
        raise HTTPException(status_code=401, detail="Сессия истекла")
    return dict(user)

def require_admin(user=Depends(require_auth)):
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Нет доступа")
    return user

# ─── Auth ────────────────────────────────────────────────────────────────────
@app.post("/api/login")
def login(body: dict):
    pin = body.get("pin", "").strip()
    if not pin:
        raise HTTPException(status_code=400, detail="Введите пин-код")
    pin_hash = hash_pin(pin)
    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE pin_hash=?", (pin_hash,)).fetchone()
    if not user:
        conn.close()
        raise HTTPException(status_code=401, detail="Неверный пин-код")
    token = secrets.token_hex(32)
    expires = (datetime.now() + timedelta(days=30)).isoformat()
    conn.execute("UPDATE users SET token=?, token_expires=? WHERE id=?",
                 (token, expires, user["id"]))
    conn.commit()
    conn.close()
    return {
        "token": token,
        "name": user["name"],
        "role": user["role"],
        "store_id": user["store_id"],
        "store_name": STORE_DATA.get(user["store_id"], {}).get("name") if user["store_id"] else None
    }

@app.post("/api/logout")
def logout(user=Depends(require_auth)):
    conn = get_db()
    conn.execute("UPDATE users SET token=NULL WHERE id=?", (user["id"],))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.get("/api/me")
def me(user=Depends(require_auth)):
    return {
        "name": user["name"],
        "role": user["role"],
        "store_id": user["store_id"],
        "store_name": STORE_DATA.get(user["store_id"], {}).get("name") if user["store_id"] else None
    }

# ─── Загрузка файла мотивации (только admin) ─────────────────────────────────
@app.post("/api/upload")
async def upload_motivation(file: UploadFile = File(...), user=Depends(require_admin)):
    content = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(content), header=None)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Ошибка чтения файла: {e}")

    # Ищем строку заголовка
    header_row = None
    for i, row in df.iterrows():
        if str(row[0]).strip().startswith("Дата"):
            header_row = i
            break
    if header_row is None:
        raise HTTPException(status_code=400, detail="Не найден заголовок 'Дата отчета'")

    data_df = df.iloc[header_row + 1:].reset_index(drop=True)
    conn = get_db()
    rows_inserted = 0
    report_date = None

    for _, row in data_df.iterrows():
        try:
            store_id = int(float(str(row[2]))) if str(row[2]) not in ["nan", ""] else None
            if not store_id:
                continue
            login = str(row[3] or "").strip()
            role = str(row[5] or "").strip()
            name = str(row[6] or "").strip()
            to_fact = float(row[7]) if str(row[7]) not in ["nan", ""] else 0
            pct_to = float(row[8]) if str(row[8]) not in ["nan", ""] else 0
            income = float(row[9]) if str(row[9]) not in ["nan", ""] else 0
            fdm_val = float(row[10]) if str(row[10]) not in ["nan", ""] else 0

            # Дата
            raw_date = row[0]
            if hasattr(raw_date, "strftime"):
                rd = raw_date.strftime("%d.%m.%Y")
            elif isinstance(raw_date, (int, float)):
                from xlrd import xldate_as_datetime
                rd = xldate_as_datetime(raw_date, 0).strftime("%d.%m.%Y")
            else:
                rd = str(raw_date).split(" ")[0].strip()
            if report_date is None:
                report_date = rd

            is_total = 1 if login == "Total" else 0
            is_bezshk = 1 if login == "БезШК" else 0

            conn.execute("""
                INSERT INTO motivation_data
                (store_id, report_date, login, role, name, to_fact, pct_to, income, fdm, is_total, is_bezshk)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """, (store_id, rd, login, role, name, to_fact, pct_to, income, fdm_val, is_total, is_bezshk))
            rows_inserted += 1
        except Exception:
            continue

    conn.execute("INSERT INTO upload_log (filename, report_date, rows_count) VALUES (?,?,?)",
                 (file.filename, report_date, rows_inserted))
    conn.commit()
    conn.close()
    return {"ok": True, "rows": rows_inserted, "date": report_date}

# ─── Данные магазина ──────────────────────────────────────────────────────────
@app.get("/api/store/{store_id}")
def get_store_data(store_id: int, days_total: int = 31, forecast_pct: int = 100, mode: str = "avg", user=Depends(require_auth)):
    # Директор видит только свой магазин
    if user["role"] == "director" and user["store_id"] != store_id:
        raise HTTPException(status_code=403, detail="Нет доступа к этому магазину")

    ref = STORE_DATA.get(store_id)
    if not ref:
        raise HTTPException(status_code=404, detail="Магазин не найден")

    conn = get_db()
    # Берём последние данные по этому магазину
    latest = conn.execute(
        "SELECT MAX(report_date) as d FROM motivation_data WHERE store_id=?", (store_id,)
    ).fetchone()
    if not latest or not latest["d"]:
        conn.close()
        return {"store_id": store_id, "name": ref["name"], "no_data": True}

    report_date = latest["d"]
    day_report = int(report_date.split(".")[0])

    rows = conn.execute(
        "SELECT * FROM motivation_data WHERE store_id=? AND report_date=?",
        (store_id, report_date)
    ).fetchall()
    conn.close()

    total_row = next((r for r in rows if r["is_total"]), None)
    staff = [r for r in rows if not r["is_total"] and not r["is_bezshk"] and r["role"] and r["role"] != "НетДолжности"]

    fact_to = total_row["to_fact"] if total_row else 0
    fact_fot = (total_row["income"] + total_row["fdm"]) if total_row else 0

    ratio = days_total / day_report
    forecast_to = fact_to * ratio if mode == "avg" else ref["plan"] * forecast_pct / 100
    plan_pct = (forecast_to / ref["plan"] * 100) if ref["plan"] > 0 else 0
    fact_plan_pct = (fact_to / ref["plan"] * 100) if ref["plan"] > 0 else 0
    coef = plan_coef(plan_pct)

    fot_dm = ref["dr"] * ref["dc"] * coef
    fot_am = ref["ar"] * ref["ac"] * coef
    fot_s = ref["sr"] * (ref["sc"] + ref["nc"]) * coef
    fot_wages = fot_dm + fot_am + fot_s
    extra = ref["cl"] + ref["ld"]
    forecast_fot = fot_wages + extra
    rest_fot = max(0, forecast_fot - fact_fot - extra * (day_report / days_total))

    def get_rate(role):
        if "Директор" in role: return ref["dr"]
        if "Администратор" in role: return ref["ar"]
        return ref["sr"]

    staff_list = sorted([{
        "name": r["name"],
        "role": r["role"],
        "rate": get_rate(r["role"]),
        "fact_income": r["income"],
        "forecast_salary": get_rate(r["role"]) * coef,
    } for r in staff], key=lambda x: (
        0 if "Директор" in x["role"] else 1 if "Администратор" in x["role"] else 2
    ))

    return {
        "store_id": store_id,
        "name": ref["name"],
        "report_date": report_date,
        "plan": ref["plan"],
        "fact_to": fact_to,
        "fact_fot": fact_fot,
        "fact_plan_pct": round(fact_plan_pct, 1),
        "forecast_to": round(forecast_to),
        "forecast_fot": round(forecast_fot),
        "forecast_plan_pct": round(plan_pct, 1),
        "coef": coef,
        "fot_wages": round(fot_wages),
        "extra": extra,
        "rest_fot": round(rest_fot),
        "fot_to_pct": round(forecast_fot / forecast_to * 100, 1) if forecast_to > 0 else 0,
        "staff": staff_list,
    }

# ─── Сводка по всем магазинам (только admin) ──────────────────────────────────
@app.get("/api/summary")
def get_summary(days_total: int = 31, forecast_pct: int = 100, mode: str = "avg", user=Depends(require_auth)):
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Нет доступа")
    result = []
    for sid in STORE_DATA:
        try:
            data = get_store_data(sid, days_total, forecast_pct, mode, user)
            if not data.get("no_data"):
                result.append(data)
        except Exception:
            continue
    result.sort(key=lambda x: x.get("forecast_to", 0), reverse=True)
    return result

# ─── Управление пользователями (только admin) ─────────────────────────────────
@app.get("/api/users")
def get_users(user=Depends(require_admin)):
    conn = get_db()
    users = conn.execute("SELECT id, name, store_id, role, created_at FROM users").fetchall()
    conn.close()
    return [{"id": u["id"], "name": u["name"], "store_id": u["store_id"],
             "role": u["role"], "store_name": STORE_DATA.get(u["store_id"], {}).get("name") if u["store_id"] else None,
             "created_at": u["created_at"]} for u in users]

@app.post("/api/users")
def create_user(body: dict, user=Depends(require_admin)):
    name = body.get("name", "").strip()
    pin = body.get("pin", "").strip()
    store_id = body.get("store_id")
    role = body.get("role", "director")
    if not name or not pin:
        raise HTTPException(status_code=400, detail="Имя и пин-код обязательны")
    if len(pin) < 4:
        raise HTTPException(status_code=400, detail="Пин-код минимум 4 символа")
    pin_hash = hash_pin(pin)
    conn = get_db()
    existing = conn.execute("SELECT id FROM users WHERE pin_hash=?", (pin_hash,)).fetchone()
    if existing:
        conn.close()
        raise HTTPException(status_code=400, detail="Такой пин-код уже существует")
    conn.execute("INSERT INTO users (name, pin_hash, store_id, role) VALUES (?,?,?,?)",
                 (name, pin_hash, store_id, role))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.delete("/api/users/{user_id}")
def delete_user(user_id: int, user=Depends(require_admin)):
    conn = get_db()
    target = conn.execute("SELECT role FROM users WHERE id=?", (user_id,)).fetchone()
    if not target:
        conn.close()
        raise HTTPException(status_code=404, detail="Пользователь не найден")
    if target["role"] == "admin":
        conn.close()
        raise HTTPException(status_code=400, detail="Нельзя удалить администратора")
    conn.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.put("/api/users/{user_id}/pin")
def change_pin(user_id: int, body: dict, user=Depends(require_admin)):
    new_pin = body.get("pin", "").strip()
    if len(new_pin) < 4:
        raise HTTPException(status_code=400, detail="Пин-код минимум 4 символа")
    conn = get_db()
    conn.execute("UPDATE users SET pin_hash=? WHERE id=?", (hash_pin(new_pin), user_id))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.get("/api/upload_log")
def upload_log(user=Depends(require_admin)):
    conn = get_db()
    logs = conn.execute("SELECT * FROM upload_log ORDER BY id DESC LIMIT 10").fetchall()
    conn.close()
    return [dict(l) for l in logs]

@app.get("/api/stores")
def get_stores_list(user=Depends(require_admin)):
    return [{"id": k, "name": v["name"]} for k, v in STORE_DATA.items()]

# ─── Отдаём PWA ──────────────────────────────────────────────────────────────
app.mount("/static", StaticFiles(directory="../frontend/static"), name="static")

@app.get("/{full_path:path}")
def serve_spa(full_path: str):
    return FileResponse("../frontend/index.html")
