from fastapi import FastAPI, HTTPException, Depends, UploadFile, File, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
import psycopg2
import psycopg2.extras
import hashlib, secrets, os, io, json
from datetime import datetime, timedelta
from typing import Optional, List
from calendar import monthrange, weekday as cal_weekday
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATABASE_URL = os.environ.get("DATABASE_URL", "")
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

# ─── Вспомогательные функции для данных магазина ─────────────────────────────
def get_store_ref_from_conn(store_id: int, conn) -> dict:
    """Возвращает плановые данные из fot_plan (последний месяц) или STORE_DATA"""
    row = conn.execute(
        "SELECT * FROM fot_plan WHERE store_id=? ORDER BY month DESC LIMIT 1",
        (store_id,)
    ).fetchone()
    if row:
        base = STORE_DATA.get(store_id, {})
        return {
            "name": base.get("name", f"Магазин {store_id}"),
            "plan": int(row["plan"] or 0),
            "sr":   int(row["sr"]   or 0),
            "ar":   int(row["ar"]   or 0),
            "dr":   int(row["dr"]   or 0),
            "dc":   float(row["dc"] or 0),
            "ac":   float(row["ac"] or 0),
            "sc":   float(row["sc"] or 0),
            "nc":   float(row["nc"] or 0),
            "cl":   int(row["cl"]   or 0),
            "ld":   int(row["ld"]   or 0),
        }
    return STORE_DATA.get(store_id)

# ─── БД (PostgreSQL) ─────────────────────────────────────────────────────────
class DBConn:
    """Wrapper making psycopg2 behave like sqlite3 for minimal code changes."""
    def __init__(self, dsn):
        self._conn = psycopg2.connect(dsn)

    def execute(self, sql, params=None):
        sql = sql.replace("?", "%s")
        cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params) if params else cur.execute(sql)
        return cur

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()

def get_db() -> DBConn:
    return DBConn(DATABASE_URL)

def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            pin_hash TEXT NOT NULL,
            store_id INTEGER,
            role TEXT DEFAULT 'director',
            token TEXT,
            token_expires TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_sessions (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL,
            token TEXT NOT NULL UNIQUE,
            token_expires TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS motivation_data (
            id SERIAL PRIMARY KEY,
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
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS upload_log (
            id SERIAL PRIMARY KEY,
            filename TEXT,
            report_date TEXT,
            rows_count INTEGER,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS fot_plan (
            id SERIAL PRIMARY KEY,
            store_id INTEGER NOT NULL,
            month TEXT NOT NULL,
            plan INTEGER DEFAULT 0,
            sr INTEGER DEFAULT 0,
            ar INTEGER DEFAULT 0,
            dr INTEGER DEFAULT 0,
            dc REAL DEFAULT 0,
            ac REAL DEFAULT 0,
            sc REAL DEFAULT 0,
            nc REAL DEFAULT 0,
            cl INTEGER DEFAULT 0,
            ld INTEGER DEFAULT 0,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS fot_upload_log (
            id SERIAL PRIMARY KEY,
            filename TEXT,
            month TEXT,
            stores_count INTEGER,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS work_schedule (
            id SERIAL PRIMARY KEY,
            store_id INTEGER NOT NULL,
            month TEXT NOT NULL,
            employee_name TEXT NOT NULL,
            employee_role TEXT DEFAULT '',
            days TEXT DEFAULT '[]',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(store_id, month, employee_name)
        )
    """)
    conn.commit()
    # Добавляем колонку first_login если нет (миграция)
    try:
        conn.execute("ALTER TABLE users ADD COLUMN first_login INTEGER DEFAULT 1")
        conn.commit()
    except Exception:
        conn.rollback()  # Сбрасываем упавшую транзакцию (колонка уже существует)

    # Создаём аккаунт администратора если нет
    if not conn.execute("SELECT id FROM users WHERE role='admin'").fetchone():
        pin_hash = hashlib.sha256(ADMIN_PIN.encode()).hexdigest()
        conn.execute(
            "INSERT INTO users (name, pin_hash, store_id, role, first_login) VALUES (?,?,NULL,'admin',0)",
            ("Администратор", pin_hash)
        )
        conn.commit()

    # Авто-создаём аккаунты директоров для всех магазинов
    default_pin_hash = hashlib.sha256("1111".encode()).hexdigest()
    for sid, sdata in STORE_DATA.items():
        if not conn.execute("SELECT id FROM users WHERE store_id=? AND role='director'", (sid,)).fetchone():
            conn.execute(
                "INSERT INTO users (name, pin_hash, store_id, role, first_login) VALUES (?,?,?,'director',1)",
                (f"Магазин {sid}", default_pin_hash, sid)
            )
    conn.commit()
    conn.close()

init_db()

# ─── Утилиты ─────────────────────────────────────────────────────────────────
def hash_pin(pin: str) -> str:
    return hashlib.sha256(pin.strip().encode()).hexdigest()

def plan_coef(pct: float) -> float:
    coef = pct / 100.0
    return min(round(coef, 4), 1.10)

def get_user_by_token(token: str):
    if not token:
        return None
    try:
        conn = get_db()
        # Ищем в таблице сессий (новый механизм)
        row = conn.execute("""
            SELECT u.* FROM users u
            JOIN user_sessions s ON s.user_id = u.id
            WHERE s.token=? AND s.token_expires > ?
        """, (token, datetime.now().isoformat())).fetchone()
        if not row:
            # Fallback: старый механизм (один токен в users)
            row = conn.execute(
                "SELECT * FROM users WHERE token=? AND token_expires > ?",
                (token, datetime.now().isoformat())
            ).fetchone()
        conn.close()
        return row
    except Exception:
        return None

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
    store_id_raw = body.get("store_id")
    if not pin:
        raise HTTPException(status_code=400, detail="Введите пин-код")
    pin_hash = hash_pin(pin)
    conn = get_db()
    if store_id_raw:
        # Вход директора: по store_id + PIN
        try:
            sid = int(store_id_raw)
        except (ValueError, TypeError):
            conn.close()
            raise HTTPException(status_code=400, detail="Неверный номер магазина")
        user = conn.execute(
            "SELECT * FROM users WHERE store_id=? AND role='director' AND pin_hash=?",
            (sid, pin_hash)
        ).fetchone()
        if not user:
            conn.close()
            raise HTTPException(status_code=401, detail="Неверный номер магазина или пин-код")
    else:
        # Вход администратора: только PIN
        user = conn.execute(
            "SELECT * FROM users WHERE role='admin' AND pin_hash=?",
            (pin_hash,)
        ).fetchone()
        if not user:
            conn.close()
            raise HTTPException(status_code=401, detail="Неверный пин-код")
    token = secrets.token_hex(32)
    expires = (datetime.now() + timedelta(days=90)).isoformat()
    conn.execute(
        "INSERT INTO user_sessions (user_id, token, token_expires) VALUES (?,?,?)",
        (user["id"], token, expires)
    )
    conn.execute("UPDATE users SET token=?, token_expires=? WHERE id=?",
                 (token, expires, user["id"]))
    conn.commit()
    first_login = bool(user.get("first_login", 0))
    conn.close()
    return {
        "token": token,
        "name": user["name"],
        "role": user["role"],
        "store_id": user["store_id"],
        "first_login": first_login,
        "store_name": STORE_DATA.get(user["store_id"], {}).get("name") if user["store_id"] else None
    }

@app.post("/api/change_pin")
def change_pin(body: dict, user=Depends(require_auth)):
    new_pin = body.get("new_pin", "").strip()
    if not new_pin or len(new_pin) < 4:
        raise HTTPException(status_code=400, detail="Пин-код минимум 4 цифры")
    if not new_pin.isdigit():
        raise HTTPException(status_code=400, detail="Пин-код должен состоять только из цифр")
    new_hash = hash_pin(new_pin)
    conn = get_db()
    conn.execute("UPDATE users SET pin_hash=?, first_login=0 WHERE id=?", (new_hash, user["id"]))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.post("/api/reset_pin/{user_id}")
def reset_pin(user_id: int, admin=Depends(require_admin)):
    default_hash = hashlib.sha256("1111".encode()).hexdigest()
    conn = get_db()
    row = conn.execute("SELECT id FROM users WHERE id=? AND role='director'", (user_id,)).fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Директор не найден")
    conn.execute("UPDATE users SET pin_hash=?, first_login=1 WHERE id=?", (default_hash, user_id))
    conn.commit()
    conn.close()
    return {"ok": True}

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
    cleared_stores = set()  # stores whose old data has been deleted

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

            # При первом появлении магазина — удаляем его старые данные (обновление без дублей)
            if store_id not in cleared_stores:
                conn.execute("DELETE FROM motivation_data WHERE store_id=?", (store_id,))
                cleared_stores.add(store_id)

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
    if user["role"] == "director" and user["store_id"] != store_id:
        raise HTTPException(status_code=403, detail="Нет доступа к этому магазину")

    conn = get_db()
    ref = get_store_ref_from_conn(store_id, conn)
    if not ref:
        conn.close()
        raise HTTPException(status_code=404, detail="Магазин не найден")
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

    # Загружаем часы из графика работы за текущий месяц
    STANDARD_HOURS = 168
    parts = report_date.split(".")
    month_str = f"{parts[2]}-{parts[1]}"  # YYYY-MM
    sched_rows = conn.execute(
        "SELECT employee_name, days FROM work_schedule WHERE store_id=? AND month=?",
        (store_id, month_str)
    ).fetchall()
    conn.close()

    hours_map = {}
    for srow in sched_rows:
        try:
            days_list = json.loads(srow["days"] or "[]")
            total_h = sum(int(d) for d in days_list if str(d).isdigit() and int(d) > 0)
            if total_h > 0:
                hours_map[srow["employee_name"]] = total_h
        except Exception:
            pass

    def get_hours_coef(name):
        h = hours_map.get(name, 0)
        return round(h / STANDARD_HOURS, 4) if h > 0 else 1.0

    total_row = next((r for r in rows if r["is_total"]), None)
    staff = [r for r in rows if not r["is_total"] and not r["is_bezshk"] and r["role"] and r["role"] != "НетДолжности"]

    fact_to = total_row["to_fact"] if total_row else 0
    fact_fot = (total_row["income"] + total_row["fdm"]) if total_row else 0

    ratio = days_total / day_report
    forecast_to = fact_to * ratio if mode == "avg" else ref["plan"] * forecast_pct / 100
    plan_pct = (forecast_to / ref["plan"] * 100) if ref["plan"] > 0 else 0
    fact_plan_pct = (fact_to / ref["plan"] * 100) if ref["plan"] > 0 else 0
    coef = plan_coef(plan_pct)

    extra = ref["cl"] + ref["ld"]

    def get_rate(role):
        if "Директор" in role: return ref["dr"]
        if "Администратор" in role: return ref["ar"]
        return ref["sr"]

    staff_list = sorted([{
        "name": r["name"],
        "role": r["role"],
        "rate": get_rate(r["role"]),
        "fact_income": r["income"],
        "hours": hours_map.get(r["name"], 0),
        "hours_coef": get_hours_coef(r["name"]),
        "forecast_salary": round(get_rate(r["role"]) * coef * get_hours_coef(r["name"])),
    } for r in staff], key=lambda x: (
        0 if "Директор" in x["role"] else 1 if "Администратор" in x["role"] else 2
    ))

    # Плановый ФОТ бюджет (коэф = 1.0, часы = 100%) — из ставок сотрудников + допрасходы
    budget_fot_wages = sum(get_rate(r["role"]) for r in staff)
    budget_fot = budget_fot_wages + extra

    # Максимально допустимый ФОТ = бюджет × коэф плана (макс 1.10)
    # При плане < 100% — не более бюджета; при 100–110% — пропорционально; выше 110% — 1.10
    max_fot = round(budget_fot * coef)

    # Расчётный ФОТ = сумма начислений сотрудников + допрасходы
    fot_wages = sum(s["forecast_salary"] for s in staff_list)
    forecast_fot_raw = fot_wages + extra

    # Применяем лимит
    fot_capped = forecast_fot_raw > max_fot
    forecast_fot = min(forecast_fot_raw, max_fot)

    rest_fot = max(0, forecast_fot - fact_fot - extra * (day_report / days_total))

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
        "budget_fot": round(budget_fot),
        "max_fot": round(max_fot),
        "fot_capped": fot_capped,
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

@app.get("/api/admin/uploads")
def get_summary_compat(days_total: int = 31, forecast_pct: int = 100, mode: str = "avg", user=Depends(require_auth)):
    """Compatibility alias for old frontend"""
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Нет доступа")
    result = []
    for sid in STORE_DATA:
        try:
            data = get_store_data(sid, days_total, forecast_pct, mode, user)
            if not data.get("no_data"):
                data["store_name"] = data.get("name", "")
                result.append(data)
        except Exception:
            continue
    result.sort(key=lambda x: x.get("forecast_to", 0), reverse=True)
    return result

# ─── Управление пользователями (только admin) ─────────────────────────────────
@app.get("/api/users")
def get_users(user=Depends(require_admin)):
    conn = get_db()
    users = conn.execute("SELECT id, name, store_id, role, first_login, created_at FROM users").fetchall()
    conn.close()
    return [{"id": u["id"], "name": u["name"], "store_id": u["store_id"],
             "role": u["role"], "first_login": bool(u.get("first_login", 0)),
             "store_name": STORE_DATA.get(u["store_id"], {}).get("name") if u["store_id"] else None,
             "created_at": str(u["created_at"])} for u in users]

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

# ─── Загрузка файла ФОТ (месячный план) ──────────────────────────────────────
@app.post("/api/fot_upload")
async def upload_fot(file: UploadFile = File(...), month: str = "", user=Depends(require_admin)):
    if not month:
        raise HTTPException(status_code=400, detail="Укажите месяц (YYYY-MM)")
    content = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(content), header=None)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Ошибка чтения файла: {e}")

    # Находим столбец со store_id (5-значный числовой ID магазина)
    known_ids = set(STORE_DATA.keys())
    sid_col = None
    for ci in range(min(6, len(df.columns))):
        for val in df.iloc[:, ci]:
            try:
                v = int(float(str(val)))
                if v in known_ids:
                    sid_col = ci
                    break
            except Exception:
                pass
        if sid_col is not None:
            break
    if sid_col is None:
        raise HTTPException(status_code=400, detail="Не найдены ID магазинов. Проверьте формат файла.")

    def safe_int(val):
        try: return int(float(str(val)))
        except: return 0

    def safe_float(val):
        try: return float(str(val))
        except: return 0.0

    conn = get_db()
    stores_inserted = 0

    for _, row in df.iterrows():
        try:
            sid = int(float(str(row.iloc[sid_col])))
            if sid not in known_ids:
                continue
        except Exception:
            continue

        # Удаляем старую запись за этот месяц для данного магазина
        conn.execute("DELETE FROM fot_plan WHERE store_id=? AND month=?", (sid, month))

        conn.execute("""
            INSERT INTO fot_plan
            (store_id, month, dc, ac, sc, nc, plan, sr, ar, dr, cl, ld)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            sid, month,
            safe_float(row.iloc[sid_col + 4]),   # Директор
            safe_float(row.iloc[sid_col + 5]),   # Администратор
            safe_float(row.iloc[sid_col + 6]),   # Продавец штат
            safe_float(row.iloc[sid_col + 7]),   # Продавец наймикс
            safe_int(row.iloc[sid_col + 8]),     # бюджет ТО
            safe_int(row.iloc[sid_col + 9]),     # Средняя ЗП Продавца
            safe_int(row.iloc[sid_col + 10]),    # Зарплата АМ
            safe_int(row.iloc[sid_col + 11]),    # Зарплата ДМ
            safe_int(row.iloc[sid_col + 18]),    # Клининг
            safe_int(row.iloc[sid_col + 20]),    # Разгрузка погрузка
        ))
        stores_inserted += 1

    conn.execute(
        "INSERT INTO fot_upload_log (filename, month, stores_count) VALUES (?,?,?)",
        (file.filename, month, stores_inserted)
    )
    conn.commit()
    conn.close()
    return {"ok": True, "stores_count": stores_inserted, "month": month}

@app.get("/api/fot_log")
def fot_log(user=Depends(require_admin)):
    conn = get_db()
    logs = conn.execute("SELECT * FROM fot_upload_log ORDER BY id DESC LIMIT 10").fetchall()
    conn.close()
    return [dict(l) for l in logs]

@app.get("/api/fot_active_month")
def fot_active_month(user=Depends(require_admin)):
    conn = get_db()
    row = conn.execute("SELECT month FROM fot_plan ORDER BY month DESC LIMIT 1").fetchone()
    conn.close()
    return {"month": row["month"] if row else None}

@app.get("/api/stores")
def get_stores_list(user=Depends(require_admin)):
    return [{"id": k, "name": v["name"]} for k, v in STORE_DATA.items()]

@app.get("/api/debug/store_ids")
def debug_store_ids():
    conn = get_db()
    rows = conn.execute(
        "SELECT DISTINCT store_id, report_date, COUNT(*) as cnt FROM motivation_data GROUP BY store_id, report_date"
    ).fetchall()
    conn.close()
    known = list(STORE_DATA.keys())
    return {
        "in_db": [{"store_id": r["store_id"], "date": r["report_date"], "rows": r["cnt"], "in_store_data": r["store_id"] in known} for r in rows],
        "store_data_ids": known
    }

# ─── График работы ────────────────────────────────────────────────────────────
@app.get("/api/schedule/{store_id}/{month}")
def get_schedule(store_id: int, month: str, user=Depends(require_auth)):
    if user["role"] == "director" and user["store_id"] != store_id:
        raise HTTPException(status_code=403, detail="Нет доступа")
    try:
        year, mon = map(int, month.split("-"))
        _, days_in_month = monthrange(year, mon)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат месяца (YYYY-MM)")

    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT * FROM work_schedule WHERE store_id=? AND month=? ORDER BY id",
            (store_id, month)
        ).fetchall()
        # PostgreSQL: DISTINCT + ORDER BY CASE требует подзапрос
        employees_db = conn.execute("""
            SELECT name, role FROM (
                SELECT DISTINCT name, role
                FROM motivation_data
                WHERE store_id=?
                  AND is_total=0 AND is_bezshk=0
                  AND name IS NOT NULL AND name NOT IN ('', 'nan')
                  AND role IS NOT NULL AND role NOT IN ('', 'НетДолжности', 'nan')
            ) sub
            ORDER BY
                CASE
                    WHEN role LIKE '%%Директор%%' THEN 1
                    WHEN role LIKE '%%Администратор%%' THEN 2
                    ELSE 3
                END, name
        """, (store_id,)).fetchall()
        conn.close()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"DB error: {str(e)}")

    schedule_map = {r["employee_name"]: json.loads(r["days"] or "[]") for r in rows}
    result = []
    for emp in employees_db:
        saved = schedule_map.get(emp["name"])
        if saved and len(saved) >= days_in_month:
            days = saved[:days_in_month]
        else:
            days = [""] * days_in_month
        result.append({"name": emp["name"], "role": emp["role"], "days": days})

    weekdays = [cal_weekday(year, mon, d + 1) for d in range(days_in_month)]

    return {
        "store_id": store_id,
        "month": month,
        "days_in_month": days_in_month,
        "weekdays": weekdays,
        "employees": result
    }


@app.post("/api/schedule/{store_id}/{month}")
def save_schedule(store_id: int, month: str, body: dict, user=Depends(require_auth)):
    if user["role"] == "director" and user["store_id"] != store_id:
        raise HTTPException(status_code=403, detail="Нет доступа")
    employees = body.get("employees", [])
    conn = get_db()
    saved = 0
    for emp in employees:
        name = str(emp.get("name", "")).strip()
        role = str(emp.get("role", ""))
        days = json.dumps(emp.get("days", []))
        if not name:
            continue
        conn.execute("""
            INSERT INTO work_schedule (store_id, month, employee_name, employee_role, days, updated_at)
            VALUES (%s,%s,%s,%s,%s,%s)
            ON CONFLICT (store_id, month, employee_name)
            DO UPDATE SET employee_role=EXCLUDED.employee_role,
                          days=EXCLUDED.days,
                          updated_at=EXCLUDED.updated_at
        """, (store_id, month, name, role, days, datetime.now().isoformat()))
        saved += 1
    conn.commit()
    conn.close()
    return {"ok": True, "saved": saved}


@app.get("/api/schedule/{store_id}/{month}/export")
def export_schedule(store_id: int, month: str, user=Depends(require_auth)):
    if user["role"] == "director" and user["store_id"] != store_id:
        raise HTTPException(status_code=403, detail="Нет доступа")

    data = get_schedule(store_id, month, user)
    store_name = STORE_DATA.get(store_id, {}).get("name", str(store_id))
    year, mon = map(int, month.split("-"))
    days_in_month = data["days_in_month"]
    weekdays = data["weekdays"]

    month_name = ["Январь","Февраль","Март","Апрель","Май","Июнь",
                  "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"][mon - 1]

    wb = Workbook()
    ws = wb.active
    ws.title = "График"

    # Fills
    fill_weekend = PatternFill("solid", fgColor="D9D9D9")
    fill_work    = PatternFill("solid", fgColor="C6EFCE")
    fill_vac     = PatternFill("solid", fgColor="FFEB9C")
    fill_sick    = PatternFill("solid", fgColor="FFC7CE")
    fill_header  = PatternFill("solid", fgColor="1A56C0")
    fill_subhdr  = PatternFill("solid", fgColor="BDD7EE")

    bold_white = Font(bold=True, color="FFFFFF")
    bold_dark  = Font(bold=True)
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # Title
    ws.merge_cells(f"A1:{get_column_letter(days_in_month + 4)}1")
    title_cell = ws["A1"]
    title_cell.value = f"График работы — {store_name} — {month_name} {year}"
    title_cell.font = bold_white
    title_cell.fill = fill_header
    title_cell.alignment = center
    ws.row_dimensions[1].height = 22

    # Column headers
    ws["A2"] = "№"
    ws["B2"] = "Сотрудник"
    ws["C2"] = "Должность"
    for d in range(1, days_in_month + 1):
        col = get_column_letter(d + 3)
        ws[f"{col}2"] = d
    ws[f"{get_column_letter(days_in_month + 4)}2"] = "Итого (ч)"

    for col in range(1, days_in_month + 5):
        cell = ws.cell(row=2, column=col)
        cell.fill = fill_subhdr
        cell.font = bold_dark
        cell.alignment = center
        cell.border = border
        # Weekend highlight in header
        if col > 3 and col < days_in_month + 4:
            if weekdays[col - 4] >= 5:
                cell.fill = fill_weekend

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    for d in range(1, days_in_month + 1):
        ws.column_dimensions[get_column_letter(d + 3)].width = 3.5
    ws.column_dimensions[get_column_letter(days_in_month + 4)].width = 6

    # Data rows
    STATUS_COLOR = {"О": fill_vac, "Б": fill_sick}
    for i, emp in enumerate(data["employees"]):
        row = i + 3
        ws.cell(row=row, column=1, value=i + 1).border = border
        ws.cell(row=row, column=2, value=emp["name"]).border = border
        ws.cell(row=row, column=3, value=emp["role"]).border = border
        total_hours = 0
        for d in range(days_in_month):
            col = d + 4
            status = emp["days"][d] if d < len(emp["days"]) else ""
            # Числовые часы
            try:
                hours = int(status)
                cell = ws.cell(row=row, column=col, value=hours)
                cell.fill = fill_work
                total_hours += hours
            except (ValueError, TypeError):
                cell = ws.cell(row=row, column=col, value=status)
                if status == "В" or (not status and weekdays[d] >= 5):
                    cell.fill = fill_weekend
                elif status in STATUS_COLOR:
                    cell.fill = STATUS_COLOR[status]
            cell.alignment = center
            cell.border = border
        tot_cell = ws.cell(row=row, column=days_in_month + 4, value=total_hours)
        tot_cell.font = bold_dark
        tot_cell.alignment = center
        tot_cell.border = border
        ws.row_dimensions[row].height = 18

    # Legend
    leg_row = len(data["employees"]) + 4
    ws.cell(row=leg_row, column=1, value="Обозначения:").font = bold_dark
    for col, (code, label, fill) in enumerate([
        ("Р","Рабочий день", fill_work),
        ("В","Выходной", fill_weekend),
        ("О","Отпуск", fill_vac),
        ("Б","Больничный", fill_sick),
    ], start=2):
        c = ws.cell(row=leg_row, column=col, value=f"{code} — {label}")
        c.fill = fill
        c.border = border

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"schedule_{store_id}_{month}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# ─── Отдаём PWA ──────────────────────────────────────────────────────────────
app.mount("/static", StaticFiles(directory="../frontend/static"), name="static")

@app.get("/{full_path:path}")
def serve_spa(full_path: str):
    return FileResponse("../frontend/index.html")
