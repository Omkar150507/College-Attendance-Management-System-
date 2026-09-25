from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
import os
import sqlite3
from urllib.parse import urlparse
from xml.sax.saxutils import escape as xml_escape

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    from psycopg2.pool import ThreadedConnectionPool
    POSTGRES_AVAILABLE = True
except ImportError:
    psycopg2 = None
    RealDictCursor = None
    ThreadedConnectionPool = None
    POSTGRES_AVAILABLE = False

from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from pathlib import Path
from functools import wraps
from datetime import date, datetime
from zoneinfo import ZoneInfo
import csv
import io
import json
import zipfile
import smtplib
import re
import secrets
from calendar import monthrange
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.drawing.image import Image as XLImage
    OPENPYXL_AVAILABLE = True
except ImportError:
    Workbook = None
    XLImage = None
    OPENPYXL_AVAILABLE = False
from email.message import EmailMessage
try:
    import qrcode
    QRCODE_AVAILABLE = True
except ImportError:
    qrcode = None
    QRCODE_AVAILABLE = False
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix
from config import COLLEGE_NAME, COLLEGE_TAGLINE, LOGO_FILE

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "attendance.db"
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
RESET_TOKEN_MAX_AGE = 30 * 60


def format_time_12(value):
    if value is None:
        return ''
    text=str(value).strip()
    if not text:
        return ''
    if '-' in text and text.count(':') >= 2:
        a,b=text.split('-',1)
        return f"{format_time_12(a)} – {format_time_12(b)}"
    m=re.match(r'^(\d{1,2})[:.](\d{2})$',text)
    if not m:
        return text
    h,minute=int(m.group(1)),int(m.group(2))
    if h>23 or minute>59:
        return text
    return f"{h%12 or 12}:{minute:02d} {'AM' if h<12 else 'PM'}"


# Small per-worker PostgreSQL pool. Render can occasionally recycle an idle
# PostgreSQL/TLS connection. A pooled connection that has gone stale must never
# be handed back to a request; otherwise pages can intermittently return HTTP 500
# with errors such as: "SSL error: decryption failed or bad record mac".
PG_POOL = None

def _discard_pg_connection(conn):
    """Return a bad PostgreSQL connection to the pool and force it to close."""
    if conn is None or PG_POOL is None:
        return
    try:
        PG_POOL.putconn(conn, close=True)
    except Exception:
        try:
            conn.close()
        except Exception:
            pass

def _new_pg_pool():
    if not POSTGRES_AVAILABLE:
        raise RuntimeError("PostgreSQL driver is not installed. Run: pip install psycopg2-binary")
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL is not configured")

    # TCP keepalives reduce the chance that an idle Render/PostgreSQL connection
    # survives in the pool after its network/TLS session has already expired.
    return ThreadedConnectionPool(
        1,
        int(os.environ.get("DB_POOL_MAX", "10")),
        DATABASE_URL,
        sslmode=os.environ.get("DB_SSLMODE", "require"),
        connect_timeout=10,
        keepalives=1,
        keepalives_idle=int(os.environ.get("DB_KEEPALIVES_IDLE", "30")),
        keepalives_interval=int(os.environ.get("DB_KEEPALIVES_INTERVAL", "10")),
        keepalives_count=int(os.environ.get("DB_KEEPALIVES_COUNT", "5")),
    )

def get_pg_pool():
    global PG_POOL
    if PG_POOL is None:
        PG_POOL = _new_pg_pool()
    return PG_POOL

def get_healthy_pg_connection():
    """Get a live connection; discard stale/broken pooled connections."""
    pool = get_pg_pool()
    last_error = None
    for _ in range(3):
        conn = pool.getconn()
        try:
            if getattr(conn, "closed", 0):
                _discard_pg_connection(conn)
                continue
            # A lightweight round-trip detects stale TLS/socket state before the
            # actual application query uses the connection.
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                cur.fetchone()
            conn.rollback()
            conn.autocommit = False
            return conn
        except Exception as exc:
            last_error = exc
            _discard_pg_connection(conn)
    raise RuntimeError(f"Unable to obtain a healthy PostgreSQL connection: {last_error}")

def get_smtp_settings():
    """Read SMTP settings at request time so Render environment changes are picked up after restart."""
    host = os.environ.get("SMTP_HOST", "").strip()
    username = os.environ.get("SMTP_USERNAME", "").strip()
    password = os.environ.get("SMTP_PASSWORD", "").strip().replace(" ", "")
    from_email = os.environ.get("SMTP_FROM_EMAIL", "").strip() or username
    try:
        port = int(os.environ.get("SMTP_PORT", "587") or 587)
    except ValueError:
        port = 587
    use_tls = os.environ.get("SMTP_USE_TLS", "1").strip().lower() not in ("0", "false", "no")
    return host, port, username, password, from_email, use_tls
YEAR_OPTIONS = ["1st Year", "2nd Year", "3rd Year", "4th Year"]
SECURITY_QUESTIONS = [
    "What was the name of your first school?",
    "What is the name of your hometown?",
    "What was your favorite subject in school?",
    "What is your favorite teacher's name?",
]

app = Flask(__name__)
app.jinja_env.filters['time12'] = format_time_12
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-change-this-secret-key")
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=bool(DATABASE_URL),
)

def reset_serializer():
    return URLSafeTimedSerializer(app.secret_key, salt="password-reset")


def send_password_reset_email(to_email, full_name, reset_link):
    host, port, username, password, from_email, use_tls = get_smtp_settings()
    missing = []
    if not host: missing.append("SMTP_HOST")
    if not username: missing.append("SMTP_USERNAME")
    if not password: missing.append("SMTP_PASSWORD")
    if not from_email: missing.append("SMTP_FROM_EMAIL")
    if missing:
        raise RuntimeError("Password reset email is not configured. Missing: " + ", ".join(missing))
    app.logger.info("SMTP configuration detected: host=%s port=%s username=%s from_email=%s tls=%s", host, port, username, from_email, use_tls)
    msg = EmailMessage()
    msg["Subject"] = f"{COLLEGE_NAME} - Password Reset"
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(
        f"Hello {full_name or 'User'},\n\n"
        f"A password reset was requested for your {COLLEGE_NAME} account.\n\n"
        f"Open this link within 30 minutes to set a new password:\n{reset_link}\n\n"
        "If you did not request this, you can ignore this email.\n"
    )
    if use_tls and port == 465:
        with smtplib.SMTP_SSL(host, port, timeout=20) as server:
            server.login(username, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=20) as server:
            if use_tls:
                server.starttls()
            server.login(username, password)
            server.send_message(msg)


class CompatRow(dict):
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)

class CompatCursor:
    def __init__(self, cursor, postgres=False):
        self.cursor = cursor
        self.postgres = postgres
        self.lastrowid = None

    def fetchone(self):
        row = self.cursor.fetchone()
        return CompatRow(row) if row is not None else None

    def fetchall(self):
        return [CompatRow(r) for r in self.cursor.fetchall()]

class CompatDB:
    def __init__(self):
        self.postgres = bool(DATABASE_URL)
        self.conn = None
        self.pool = None
        self._closed = False
        if self.postgres:
            self.pool = get_pg_pool()
            self.conn = get_healthy_pg_connection()
            self.conn.autocommit = False
        else:
            self.conn = sqlite3.connect(DB_PATH)
            self.conn.row_factory = sqlite3.Row
            self.conn.execute("PRAGMA foreign_keys = ON")

    def _prepare_sql(self, sql):
        sql = sql.replace("?", "%s")
        # SQLite-specific INSERT OR IGNORE -> PostgreSQL equivalent.
        sql = sql.replace("INSERT OR IGNORE INTO", "INSERT INTO")
        if "INSERT INTO subject_teachers" in sql and "ON CONFLICT" not in sql:
            sql = sql.rstrip().rstrip(";") + " ON CONFLICT (teacher_id, subject_id) DO NOTHING"
        return sql

    def _replace_broken_connection(self):
        old = self.conn
        _discard_pg_connection(old)
        self.conn = get_healthy_pg_connection()
        self.conn.autocommit = False

    def execute(self, sql, params=()):
        if self.postgres:
            sql = self._prepare_sql(sql)
            try:
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.execute(sql, params)
                return CompatCursor(cur, True)
            except (psycopg2.OperationalError, psycopg2.InterfaceError) as exc:
                # Retry only when the connection is still at the beginning of a
                # transaction. Retrying inside a partially completed transaction
                # could silently lose earlier writes. This safely fixes the common
                # stale pooled-connection case seen on read-only page loads.
                try:
                    can_retry = self.conn.get_transaction_status() == psycopg2.extensions.TRANSACTION_STATUS_IDLE
                except Exception:
                    can_retry = False
                if not can_retry:
                    raise exc
                self._replace_broken_connection()
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.execute(sql, params)
                return CompatCursor(cur, True)
        cur = self.conn.execute(sql, params)
        wrapped = CompatCursor(cur, False)
        wrapped.lastrowid = cur.lastrowid
        return wrapped

    def executemany(self, sql, seq_of_params):
        """Execute a parameterized statement for each parameter tuple."""
        if self.postgres:
            sql = self._prepare_sql(sql)
            try:
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.executemany(sql, seq_of_params)
                return CompatCursor(cur, True)
            except (psycopg2.OperationalError, psycopg2.InterfaceError) as exc:
                try:
                    can_retry = self.conn.get_transaction_status() == psycopg2.extensions.TRANSACTION_STATUS_IDLE
                except Exception:
                    can_retry = False
                if not can_retry:
                    raise exc
                self._replace_broken_connection()
                cur = self.conn.cursor(cursor_factory=RealDictCursor)
                cur.executemany(sql, seq_of_params)
                return CompatCursor(cur, True)
        cur = self.conn.executemany(sql, seq_of_params)
        wrapped = CompatCursor(cur, False)
        wrapped.lastrowid = cur.lastrowid
        return wrapped

    def executescript(self, script):
        if self.postgres:
            # PostgreSQL schema used for hosted deployment.
            statements = [x.strip() for x in script.split(";") if x.strip()]
            for stmt in statements:
                stmt = stmt.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")
                stmt = stmt.replace("CURRENT_TIMESTAMP", "CURRENT_TIMESTAMP")
                stmt = stmt.replace("TEXT DEFAULT ''", "TEXT DEFAULT ''")
                self.execute(stmt)
        else:
            self.conn.executescript(script)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        try:
            self.conn.rollback()
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        self._closed = True
        if self.postgres:
            conn = self.conn
            self.conn = None
            if conn is None:
                return
            try:
                # Always rollback before returning a connection to the pool so a
                # later request never inherits an unfinished transaction.
                conn.rollback()
            except Exception:
                pass
            try:
                if getattr(conn, "closed", 0):
                    self.pool.putconn(conn, close=True)
                else:
                    self.pool.putconn(conn)
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass
        else:
            self.conn.close()


def get_db():
    return CompatDB()


def ensure_column(db, table, column, definition):
    if db.postgres:
        exists = db.execute("SELECT 1 FROM information_schema.columns WHERE table_name=? AND column_name=?", (table, column)).fetchone()
        if not exists:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
    else:
        cols = {r[1] for r in db.execute(f"PRAGMA table_info({table})").fetchall()}
        if column not in cols:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def migrate_attendance_for_lectures(db):
    """Upgrade old attendance schema to support lecture 1/2/3 per subject per day."""
    if db.postgres:
        cols = db.execute("SELECT column_name FROM information_schema.columns WHERE table_name='attendance'").fetchall()
        names = {r["column_name"] for r in cols}
        if "lecture_no" not in names:
            db.execute("ALTER TABLE attendance ADD COLUMN lecture_no INTEGER NOT NULL DEFAULT 1")
        if "lecture_time" not in names:
            db.execute("ALTER TABLE attendance ADD COLUMN lecture_time TEXT DEFAULT ''")
        # Replace the old 3-column unique constraint with the new 4-column constraint.
        constraints = db.execute("""SELECT conname FROM pg_constraint
            WHERE conrelid='attendance'::regclass AND contype='u'""").fetchall()
        for c in constraints:
            db.execute(f'ALTER TABLE attendance DROP CONSTRAINT IF EXISTS "{c["conname"]}"')
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_attendance_student_subject_date_lecture ON attendance(student_id,subject_id,attendance_date,lecture_no)")
    else:
        cols = db.execute("PRAGMA table_info(attendance)").fetchall()
        names = {r[1] for r in cols}
        if "lecture_no" not in names:
            db.execute("ALTER TABLE attendance RENAME TO attendance_old")
            db.execute("""CREATE TABLE attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                subject_id INTEGER NOT NULL,
                attendance_date TEXT NOT NULL,
                lecture_no INTEGER NOT NULL DEFAULT 1,
                lecture_time TEXT DEFAULT '',
                status TEXT NOT NULL CHECK(status IN ('Present','Absent')),
                marked_by INTEGER,
                UNIQUE(student_id, subject_id, attendance_date, lecture_no),
                FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
                FOREIGN KEY(marked_by) REFERENCES users(id) ON DELETE SET NULL
            )""")
            db.execute("""INSERT INTO attendance(id,student_id,subject_id,attendance_date,lecture_no,lecture_time,status,marked_by)
                SELECT id,student_id,subject_id,attendance_date,1,'',status,marked_by FROM attendance_old""")
            db.execute("DROP TABLE attendance_old")
        else:
            db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_attendance_student_subject_date_lecture ON attendance(student_id,subject_id,attendance_date,lecture_no)")


def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        roll_no TEXT UNIQUE NOT NULL,
        prn TEXT DEFAULT '',
        name TEXT NOT NULL,
        course TEXT NOT NULL DEFAULT 'B.Pharm',
        year TEXT NOT NULL DEFAULT '1st Year',
        division TEXT NOT NULL DEFAULT 'A',
        batch TEXT DEFAULT '',
        blocked INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'teacher',
        full_name TEXT DEFAULT '',
        email TEXT DEFAULT '',
        mobile TEXT DEFAULT '',
        security_question TEXT DEFAULT '',
        security_answer TEXT DEFAULT '',
        approved INTEGER NOT NULL DEFAULT 1,
        student_id INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS subjects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL,
        year TEXT NOT NULL DEFAULT '1st Year'
    );

    CREATE TABLE IF NOT EXISTS subject_teachers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        UNIQUE(teacher_id, subject_id),
        FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        attendance_date TEXT NOT NULL,
        lecture_no INTEGER NOT NULL DEFAULT 1,
        lecture_time TEXT DEFAULT '',
        session_type TEXT NOT NULL DEFAULT 'Lecture',
        batch TEXT DEFAULT '',
        status TEXT NOT NULL CHECK(status IN ('Present','Absent')),
        marked_by INTEGER,
        UNIQUE(student_id, subject_id, attendance_date, lecture_no),
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
        FOREIGN KEY(marked_by) REFERENCES users(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        action TEXT NOT NULL,
        details TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS attendance_corrections (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        attendance_id INTEGER,
        student_id INTEGER NOT NULL,
        subject_id INTEGER NOT NULL,
        attendance_date TEXT NOT NULL,
        lecture_no INTEGER NOT NULL DEFAULT 1,
        old_status TEXT,
        new_status TEXT NOT NULL,
        reason TEXT NOT NULL,
        corrected_by INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(attendance_id) REFERENCES attendance(id) ON DELETE SET NULL,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
        FOREIGN KEY(corrected_by) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS app_settings (
        key TEXT PRIMARY KEY,
        value TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS departments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,
        name TEXT UNIQUE NOT NULL,
        active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS timetable (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        department_id INTEGER,
        year TEXT NOT NULL,
        day_of_week TEXT NOT NULL,
        lecture_no INTEGER NOT NULL,
        start_time TEXT NOT NULL,
        end_time TEXT NOT NULL,
        subject_id INTEGER,
        teacher_id INTEGER,
        room TEXT DEFAULT '',
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE SET NULL,
        FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        title TEXT NOT NULL,
        message TEXT NOT NULL,
        link TEXT DEFAULT '',
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS leave_applications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        from_date TEXT NOT NULL,
        to_date TEXT NOT NULL,
        reason TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'Pending',
        reviewed_by INTEGER,
        review_note TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(reviewed_by) REFERENCES users(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS qr_sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        token TEXT UNIQUE NOT NULL,
        subject_id INTEGER NOT NULL,
        year TEXT NOT NULL,
        attendance_date TEXT NOT NULL,
        lecture_no INTEGER NOT NULL,
        lecture_time TEXT DEFAULT '',
        session_type TEXT NOT NULL DEFAULT 'Lecture',
        batch TEXT DEFAULT '',
        teacher_id INTEGER NOT NULL,
        expires_at TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
        FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE
    );
    """)

    # Migrations for databases created by earlier versions.
    ensure_column(db, "users", "full_name", "TEXT DEFAULT ''")
    ensure_column(db, "users", "email", "TEXT DEFAULT ''")
    ensure_column(db, "users", "mobile", "TEXT DEFAULT ''")
    ensure_column(db, "users", "security_question", "TEXT DEFAULT ''")
    ensure_column(db, "users", "security_answer", "TEXT DEFAULT ''")
    ensure_column(db, "users", "approved", "INTEGER NOT NULL DEFAULT 1")
    ensure_column(db, "users", "student_id", "INTEGER")
    ensure_column(db, "users", "created_at", "TEXT DEFAULT CURRENT_TIMESTAMP")
    ensure_column(db, "students", "prn", "TEXT DEFAULT ''")
    ensure_column(db, "students", "department_id", "INTEGER")
    ensure_column(db, "users", "department_id", "INTEGER")
    ensure_column(db, "subjects", "department_id", "INTEGER")
    ensure_column(db, "subjects", "year", "TEXT DEFAULT '1st Year'" )
    ensure_column(db, "attendance", "marked_by", "INTEGER")
    ensure_column(db, "attendance", "academic_year", "TEXT DEFAULT ''")
    ensure_column(db, "attendance", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    ensure_column(db, "attendance", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "attendance", "scheduled_teacher_id", "INTEGER")
    ensure_column(db, "attendance", "conducted_by_teacher_id", "INTEGER")
    ensure_column(db, "attendance", "teacher_status", "TEXT DEFAULT 'Scheduled teacher present'")
    ensure_column(db, "attendance", "start_time", "TEXT DEFAULT ''")
    ensure_column(db, "attendance", "end_time", "TEXT DEFAULT ''")
    ensure_column(db, "qr_sessions", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    ensure_column(db, "qr_sessions", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "qr_sessions", "scheduled_teacher_id", "INTEGER")
    ensure_column(db, "qr_sessions", "conducted_by_teacher_id", "INTEGER")
    ensure_column(db, "qr_sessions", "teacher_status", "TEXT DEFAULT 'Scheduled teacher present'")
    ensure_column(db, "qr_sessions", "start_time", "TEXT DEFAULT ''")
    ensure_column(db, "qr_sessions", "end_time", "TEXT DEFAULT ''")
    ensure_column(db, "timetable", "session_type", "TEXT NOT NULL DEFAULT 'Lecture'")
    ensure_column(db, "timetable", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "students", "academic_year", "TEXT DEFAULT ''")
    ensure_column(db, "students", "batch", "TEXT DEFAULT ''")
    ensure_column(db, "students", "blocked", "INTEGER NOT NULL DEFAULT 0")
    history_id_def = "SERIAL PRIMARY KEY" if db.postgres else "INTEGER PRIMARY KEY AUTOINCREMENT"
    db.execute(f"""CREATE TABLE IF NOT EXISTS student_academic_history (
        id {history_id_def},
        student_id INTEGER NOT NULL,
        academic_year TEXT NOT NULL,
        year TEXT NOT NULL,
        promoted_at TEXT DEFAULT CURRENT_TIMESTAMP,
        promoted_by INTEGER,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(promoted_by) REFERENCES users(id) ON DELETE SET NULL,
        UNIQUE(student_id, academic_year)
    )""")
    current_ay = db.execute("SELECT value FROM app_settings WHERE key='academic_year'").fetchone()
    current_ay_value = (current_ay["value"] if current_ay and current_ay["value"] else "2026-27")
    db.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO NOTHING", ("academic_year", current_ay_value))
    db.execute("UPDATE students SET academic_year=? WHERE academic_year IS NULL OR academic_year=''", (current_ay_value,))
    db.execute("UPDATE attendance SET academic_year=? WHERE academic_year IS NULL OR academic_year=''", (current_ay_value,))
    # Multi-lecture attendance migration: old versions had one record per student/subject/day.
    # Rebuild the attendance table once so a subject can have up to 3 lectures per day.
    migrate_attendance_for_lectures(db)

    # Indexes for fast filtering/reporting when many users access the system together.
    db.execute("CREATE INDEX IF NOT EXISTS idx_students_year ON students(year)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_students_name ON students(name)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_students_batch ON students(year,batch)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(attendance_date)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_subject_date ON attendance(subject_id,attendance_date)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_student_date ON attendance(student_id,attendance_date)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_attendance_academic_year ON attendance(academic_year)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_student_academic_year ON students(academic_year)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_activity_created ON activity_logs(created_at)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_notifications_user ON notifications(user_id,is_read,created_at)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_leave_status ON leave_applications(status,created_at)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_qr_token ON qr_sessions(token,active)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_timetable_day ON timetable(year,day_of_week)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_correction_date ON attendance_corrections(attendance_date)")
    # One PRN can belong to only one student when it is not blank.
    db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_students_prn_unique ON students(prn) WHERE prn IS NOT NULL AND prn <> ''")

    if db.execute("SELECT COUNT(*) FROM departments").fetchone()[0] == 0:
        db.execute("INSERT INTO departments(code,name,active) VALUES(?,?,1)", ("PHARM", "Pharmacy"))
    default_dept = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()[0]
    db.execute("UPDATE students SET department_id=? WHERE department_id IS NULL", (default_dept,))
    db.execute("UPDATE subjects SET department_id=? WHERE department_id IS NULL", (default_dept,))
    db.execute("UPDATE users SET department_id=? WHERE department_id IS NULL AND role IN ('teacher','hod')", (default_dept,))

    admin_username = os.environ.get("ADMIN_USERNAME", "admin").strip()
    admin_password = os.environ.get("ADMIN_PASSWORD", "")
    admin_name = os.environ.get("ADMIN_NAME", "College Administrator").strip() or "College Administrator"
    admin_email = os.environ.get("ADMIN_EMAIL", "").strip().lower()
    if db.execute("SELECT COUNT(*) FROM users WHERE role='admin'").fetchone()[0] == 0:
        if db.postgres and not admin_password:
            db.close()
            raise RuntimeError("ADMIN_PASSWORD must be set when using PostgreSQL. Set a strong admin password in Render Environment Variables.")
        admin_password = admin_password or "admin123"
        db.execute("INSERT INTO users(username,password,role,full_name,email,approved) VALUES(?,?,?,?,?,?)",
                   (admin_username, generate_password_hash(admin_password), "admin", admin_name, admin_email, 1))
    else:
        db.execute("UPDATE users SET approved=1 WHERE role='admin'")
        if admin_email:
            db.execute("UPDATE users SET email=? WHERE role='admin' AND username=?", (admin_email, admin_username))

    # Old demo teacher is intentionally removed; teachers now register.
    db.execute("DELETE FROM users WHERE role='teacher' AND username='teacher'")

    if db.execute("SELECT COUNT(*) FROM students").fetchone()[0] == 0:
        students = [
            ("01", "01", "Rahul Sharma", "B.Pharm", "1st Year", "A"),
            ("02", "02", "Priya Patel", "B.Pharm", "1st Year", "A"),
            ("03", "03", "Ankit Yadav", "B.Pharm", "1st Year", "A"),
            ("04", "04", "Sneha Gupta", "B.Pharm", "1st Year", "A"),
            ("05", "05", "Rohan Singh", "B.Pharm", "1st Year", "A"),
        ]
        db.executemany("INSERT INTO students(roll_no,prn,name,course,year,division) VALUES(?,?,?,?,?,?)", students)

    if db.execute("SELECT COUNT(*) FROM subjects").fetchone()[0] == 0:
        subjects = [
            ("BP503T", "Pharmacology-II"),
            ("BP502T", "Pharmacognosy-II"),
            ("BP501T", "Medicinal Chemistry-II"),
        ]
        db.executemany("INSERT INTO subjects(code,name,year) VALUES(?,?,?)", [(c,n,"1st Year") for c,n in subjects])

    db.commit()
    db.close()


@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    if request.is_secure:
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response


def validate_password(password):
    if len(password) < 8:
        return "Password must be at least 8 characters."
    if not any(c.isupper() for c in password):
        return "Password must contain at least one uppercase letter."
    if not any(c.islower() for c in password):
        return "Password must contain at least one lowercase letter."
    if not any(c.isdigit() for c in password):
        return "Password must contain at least one number."
    return None


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped


def staff_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") not in ("admin", "teacher", "hod"):
            flash("This page is available only to teachers and administrators.", "error")
            return redirect(url_for("student_dashboard"))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Admin access required.", "error")
            return redirect(url_for("dashboard"))
        return view(*args, **kwargs)
    return wrapped


def student_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "student":
            flash("Student access required.", "error")
            return redirect(url_for("dashboard"))
        try:
            db = get_db()
            row = db.execute("SELECT blocked FROM students WHERE id=?", (session.get("student_id"),)).fetchone()
            db.close()
            if row and int(row["blocked"] or 0) == 1:
                session.clear()
                flash("You are not registered. Please contact the college administration.", "error")
                return redirect(url_for("login"))
        except Exception:
            try: db.close()
            except Exception: pass
        return view(*args, **kwargs)
    return wrapped


def teacher_or_admin_subjects(db):
    if session.get("role") == "admin":
        return db.execute("SELECT * FROM subjects ORDER BY code").fetchall()
    if session.get("role") == "hod":
        dept = db.execute("SELECT department_id FROM users WHERE id=?", (session["user_id"],)).fetchone()
        return db.execute("SELECT * FROM subjects WHERE department_id=? ORDER BY code", (dept["department_id"],)).fetchall() if dept and dept["department_id"] else []
    return db.execute("""
        SELECT s.* FROM subjects s
        JOIN subject_teachers st ON st.subject_id=s.id
        WHERE st.teacher_id=? ORDER BY s.code
    """, (session["user_id"],)).fetchall()


def can_use_subject(db, subject_id):
    if session.get("role") == "admin":
        return True
    if session.get("role") == "hod":
        row = db.execute("SELECT 1 FROM subjects s JOIN users u ON u.department_id=s.department_id WHERE s.id=? AND u.id=?", (subject_id, session["user_id"])).fetchone()
        return bool(row)
    row = db.execute("SELECT 1 FROM subject_teachers WHERE teacher_id=? AND subject_id=?",
                     (session["user_id"], subject_id)).fetchone()
    return bool(row)


def current_student(db):
    return db.execute("""
        SELECT s.* FROM students s
        JOIN users u ON u.student_id=s.id
        WHERE u.id=? AND u.role='student'
    """, (session["user_id"],)).fetchone()

def log_activity(action, details=""):
    try:
        db = get_db()
        db.execute("INSERT INTO activity_logs(user_id,action,details,created_at) VALUES(?,?,?,?)",
                   (session.get("user_id"), action, details, datetime.utcnow().isoformat()))
        db.commit(); db.close()
    except Exception:
        pass


@app.context_processor
def inject_globals():
    unread_notifications = 0
    if session.get("user_id"):
        try:
            db = get_db()
            unread_notifications = db.execute("SELECT COUNT(*) c FROM notifications WHERE user_id=? AND is_read=0", (session["user_id"],)).fetchone()["c"]
            db.close()
        except Exception:
            pass
    today_timetable=[]
    try:
        db=get_db()
        if session.get("role")=="student":
            st=current_student(db); today_timetable=get_timetable_rows(db, role="student", user_id=session.get("user_id"), student=st) if st else []
        elif session.get("role")=="teacher": today_timetable=get_timetable_rows(db, role="teacher", user_id=session.get("user_id"))
        elif session.get("role")=="admin": today_timetable=get_timetable_rows(db, role="admin", user_id=session.get("user_id"))
        db.close()
    except Exception:
        try: db.close()
        except Exception: pass
    return {
        "today_timetable": today_timetable,
        "today": date.today().isoformat(),
        "current_user": session.get("username"),
        "college_name": COLLEGE_NAME,
        "college_tagline": COLLEGE_TAGLINE,
        "logo_file": LOGO_FILE,
        "unread_notifications": unread_notifications,
    }


@app.get("/health")
def health():
    return {"status": "ok", "service": "Y.N.P. College of Pharmacy Attendance System"}


@app.route("/", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        if user and check_password_hash(user["password"], password):
            if user["role"] == "student":
                blocked = db.execute("SELECT blocked FROM students WHERE id=?", (user["student_id"],)).fetchone() if user["student_id"] else None
                if blocked and int(blocked["blocked"] or 0) == 1:
                    db.close()
                    flash("You are not registered. Please contact the college administration.", "error")
                    return redirect(url_for("login"))
            db.close()
            if user["approved"] == 0:
                message = "Your teacher account is waiting for admin approval." if user["role"] == "teacher" else "Your student account is waiting for approval."
                flash(message, "error")
                return redirect(url_for("login"))
            session.clear()
            session.update(user_id=user["id"], username=user["username"], role=user["role"],
                           full_name=user["full_name"] or user["username"], student_id=user["student_id"])
            if user["role"] == "student":
                return redirect(url_for("student_dashboard"))
            if user["role"] == "hod":
                return redirect(url_for("hod_dashboard"))
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "error")
    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    db0 = get_db(); departments = db0.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall(); db0.close()
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        username = request.form.get("username", "").strip()
        department_id = request.form.get("department_id", type=int)
        email = request.form.get("email", "").strip().lower()
        mobile = request.form.get("mobile", "").strip()
        security_question = request.form.get("security_question", "").strip()
        security_answer = request.form.get("security_answer", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        if not all([full_name, username, email, mobile, security_question, security_answer, password, confirm]):
            flash("All fields are required. Please fill every field.", "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS, departments=departments)
        if security_question not in SECURITY_QUESTIONS:
            flash("Please select a valid security question.", "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS, departments=departments)
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS, departments=departments)
        password_error = validate_password(password)
        if password_error:
            flash(password_error, "error")
            return render_template("register.html", security_questions=SECURITY_QUESTIONS, departments=departments)
        db = get_db()
        try:
            if db.execute("SELECT id FROM users WHERE lower(email)=lower(?)", (email,)).fetchone():
                db.close(); flash("This email is already linked to another account. Use a different email.", "error")
                return render_template("register.html", security_questions=SECURITY_QUESTIONS, departments=departments)
            if not department_id:
                department_id = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()[0]
            db.execute("""INSERT INTO users(username,password,role,full_name,email,mobile,security_question,security_answer,approved,department_id)
                          VALUES(?,?,?,?,?,?,?,?,0,?)""",
                       (username, generate_password_hash(password), "teacher", full_name, email, mobile,
                        security_question, generate_password_hash(security_answer.casefold()), department_id))
            db.commit(); db.close()
            flash("Teacher registration submitted. Ask the college admin to approve your account and assign subjects.", "success")
            return redirect(url_for("login"))
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)):
                raise
            db.close()
            flash("Username already exists. Choose another username.", "error")
    return render_template("register.html", security_questions=SECURITY_QUESTIONS, departments=departments)


@app.route("/student/register", methods=["GET", "POST"])
def student_register():
    """Create a login account for a student already added by Admin.

    Student Management is the source of truth for PRN, name, year and batch.
    Registration only creates/links the login account; it never creates a
    second student record and never asks the student to choose a batch.
    """
    if request.method == "POST":
        prn = request.form.get("prn", "").strip()
        name = " ".join(request.form.get("name", "").strip().split())
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        mobile = request.form.get("mobile", "").strip()
        security_question = request.form.get("security_question", "").strip()
        security_answer = request.form.get("security_answer", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        year = request.form.get("year", "").strip()
        context = dict(
            year_options=YEAR_OPTIONS,
            selected_year=year,
            security_questions=SECURITY_QUESTIONS,
            departments=[],
        )

        if not all([prn, name, email, username, mobile, security_question,
                    security_answer, password, confirm, year]):
            flash("All fields are required. Please fill every field and select your year.", "error")
            return render_template("student_register.html", **context)
        if year not in YEAR_OPTIONS:
            flash("Please select a valid year.", "error")
            return render_template("student_register.html", **context)
        if security_question not in SECURITY_QUESTIONS:
            flash("Please select a valid security question.", "error")
            return render_template("student_register.html", **context)
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("student_register.html", **context)
        password_error = validate_password(password)
        if password_error:
            flash(password_error, "error")
            return render_template("student_register.html", **context)

        db = get_db()
        try:
            # The Admin-created Student Management record is required.
            student = db.execute(
                "SELECT * FROM students WHERE prn=? LIMIT 1", (prn,)
            ).fetchone()
            if not student:
                db.close()
                flash("This PRN is not in Student Management. Please ask the admin to add the student first.", "error")
                return render_template("student_register.html", **context)

            if int(student["blocked"] or 0) == 1:
                db.close()
                flash("You are not registered. Please contact the college administration.", "error")
                return render_template("student_register.html", **context)

            # One student record can have only one student login account.
            linked = db.execute(
                "SELECT id FROM users WHERE student_id=? LIMIT 1", (student["id"],)
            ).fetchone()
            if linked:
                db.close()
                flash("An account for this PRN already exists. Please log in instead.", "error")
                return render_template("student_register.html", **context)

            # Registration must match the Admin-created student's identity.
            existing_name = " ".join(str(student["name"] or "").strip().split())
            if existing_name.casefold() != name.casefold():
                db.close()
                flash("The name does not match the student record. Please enter the name exactly as added by the admin.", "error")
                return render_template("student_register.html", **context)
            if str(student["year"] or "").strip() != year:
                db.close()
                flash("The selected year does not match the student record. Please select the year assigned by the admin.", "error")
                return render_template("student_register.html", **context)

            if db.execute("SELECT id FROM users WHERE lower(email)=lower(?)", (email,)).fetchone():
                db.close()
                flash("This email is already linked to another account. Use a different email.", "error")
                return render_template("student_register.html", **context)
            if db.execute("SELECT id FROM users WHERE lower(username)=lower(?)", (username,)).fetchone():
                db.close()
                flash("This username is already in use. Choose another username.", "error")
                return render_template("student_register.html", **context)

            # Reuse the student's existing internal department relationship.
            department_id = student["department_id"] if student["department_id"] is not None else None
            db.execute(
                """INSERT INTO users(username,password,role,full_name,email,mobile,
                   security_question,security_answer,approved,student_id,department_id)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (username, generate_password_hash(password), "student", student["name"],
                 email, mobile, security_question,
                 generate_password_hash(security_answer.casefold()), 1,
                 student["id"], department_id),
            )
            db.commit()
            db.close()
            flash("Student account created successfully. You can now log in.", "success")
            return redirect(url_for("login"))
        except Exception as exc:
            try:
                db.close()
            except Exception:
                pass
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)):
                raise
            flash("Username, email or account details already exist. Please use different details.", "error")

    return render_template(
        "student_register.html",
        year_options=YEAR_OPTIONS,
        selected_year="",
        security_questions=SECURITY_QUESTIONS,
        departments=[],
    )


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    # Do not carry login-page/block messages into the password-recovery page.
    # They should be shown on the login page only.
    if request.method == "GET":
        session.pop("_flashes", None)
    if request.method == "POST":
        step = request.form.get("step", "1")
        if step == "2":
            user_id = session.get("password_reset_user_id")
            if not user_id:
                flash("Recovery session expired. Please start again.", "error")
                return redirect(url_for("forgot_password"))
            answer = request.form.get("security_answer", "").strip().casefold()
            attempts = int(session.get("password_reset_attempts", 0)) + 1
            session["password_reset_attempts"] = attempts
            db = get_db()
            user = db.execute("SELECT id,security_answer,approved FROM users WHERE id=?", (user_id,)).fetchone()
            db.close()
            if attempts > 5:
                session.pop("password_reset_user_id", None); session.pop("password_reset_attempts", None)
                flash("Too many incorrect attempts. Please start password recovery again.", "error")
                return redirect(url_for("forgot_password"))
            if user and user["approved"] and user["security_answer"] and check_password_hash(user["security_answer"], answer):
                session["password_reset_verified"] = True
                return redirect(url_for("reset_password"))
            flash(f"Incorrect security answer. Attempts remaining: {max(0, 5-attempts)}", "error")
            db = get_db(); row = db.execute("SELECT security_question FROM users WHERE id=?", (user_id,)).fetchone(); db.close()
            return render_template("forgot_password.html", step=2, security_question=row["security_question"] if row else "")

        account_type = request.form.get("account_type", "student")
        identifier = request.form.get("identifier", "").strip()
        mobile = request.form.get("mobile", "").strip()
        if not identifier or not mobile:
            flash("Please enter all required details.", "error")
            return render_template("forgot_password.html", step=1)
        db = get_db()
        if account_type == "teacher":
            user = db.execute("SELECT id,security_question,approved FROM users WHERE lower(username)=lower(?) AND mobile=? AND role='teacher' LIMIT 1", (identifier, mobile)).fetchone()
        else:
            user = db.execute("""SELECT u.id,u.security_question,u.approved FROM users u
                                JOIN students s ON s.id=u.student_id
                                WHERE s.prn=? AND u.mobile=? AND u.role='student' LIMIT 1""", (identifier, mobile)).fetchone()
        db.close()
        session.pop("password_reset_verified", None)
        session.pop("password_reset_user_id", None)
        session.pop("password_reset_attempts", None)
        if not user or not user["approved"] or not user["security_question"]:
            flash("The details do not match an approved account, or security recovery is not set up.", "error")
            return render_template("forgot_password.html", step=1)
        session["password_reset_user_id"] = user["id"]
        session["password_reset_attempts"] = 0
        return render_template("forgot_password.html", step=2, security_question=user["security_question"])
    return render_template("forgot_password.html", step=1)


@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    if not session.get("password_reset_verified") or not session.get("password_reset_user_id"):
        flash("Please verify your recovery details first.", "error")
        return redirect(url_for("forgot_password"))
    user_id = session["password_reset_user_id"]
    if request.method == "POST":
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        password_error = validate_password(password)
        if password_error:
            flash(password_error, "error")
            return render_template("reset_password.html")
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("reset_password.html")
        db = get_db()
        db.execute("UPDATE users SET password=? WHERE id=?", (generate_password_hash(password), user_id))
        db.commit(); db.close()
        session.pop("password_reset_verified", None); session.pop("password_reset_user_id", None); session.pop("password_reset_attempts", None)
        flash("Password changed successfully. You can now log in with your new password.", "success")
        return redirect(url_for("login"))
    return render_template("reset_password.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def get_timetable_rows(db, role=None, user_id=None, student=None, day_name=None):
    day_name = day_name or date.today().strftime("%A")
    where=["t.day_of_week=?"]; params=[day_name]
    if student is not None:
        where.append("t.year=?"); params.append(student["year"])
        where.append("(COALESCE(t.session_type,'Lecture')='Lecture' OR COALESCE(t.batch,'')=?)"); params.append(student["batch"] or "")
    elif role == "teacher":
        where.append("t.teacher_id=?"); params.append(user_id)
    elif role == "hod":
        where.append("t.department_id=(SELECT department_id FROM users WHERE id=? )"); params.append(user_id)
    return db.execute(f"""SELECT t.*, COALESCE(t.session_type,'Lecture') session_type, COALESCE(t.batch,'') batch,
        COALESCE(sub.code,'') subject_code, COALESCE(sub.name,'') subject_name,
        COALESCE(u.full_name,u.username,'') teacher_name
        FROM timetable t LEFT JOIN subjects sub ON sub.id=t.subject_id LEFT JOIN users u ON u.id=t.teacher_id
        WHERE {' AND '.join(where)} ORDER BY t.start_time,t.lecture_no,t.id""", tuple(params)).fetchall()


def timetable_candidates(db, year, session_type, batch, subject_id, day_name):
    return db.execute("""SELECT t.*, COALESCE(u.full_name,u.username,'') teacher_name,
        COALESCE(sub.code,'') subject_code, COALESCE(sub.name,'') subject_name
        FROM timetable t LEFT JOIN users u ON u.id=t.teacher_id LEFT JOIN subjects sub ON sub.id=t.subject_id
        WHERE t.year=? AND t.day_of_week=? AND t.subject_id=?
          AND COALESCE(t.session_type,'Lecture')=?
          AND (COALESCE(t.batch,'')=? OR ?='Lecture')
        ORDER BY t.start_time,t.lecture_no,t.id""",
        (year,day_name,subject_id,session_type,batch if session_type=='Practical' else '',session_type)).fetchall()


def current_academic_year_value(db):
    row=db.execute("SELECT value FROM app_settings WHERE key='academic_year'").fetchone()
    return row['value'] if row and row['value'] else '2026-27'


def all_teacher_rows(db):
    return db.execute("SELECT id,COALESCE(full_name,username) name FROM users WHERE role='teacher' AND approved=1 ORDER BY name").fetchall()


@app.route("/timetable", methods=["GET", "POST"])
@login_required
def timetable():
    db=get_db(); role=session.get("role")
    days=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    if request.method=='POST':
        if role!='admin':
            db.close(); flash('Only administrators can manage the timetable.','error'); return redirect(url_for('timetable'))
        year=request.form.get('year','').strip(); day=request.form.get('day_of_week','').strip()
        start=request.form.get('start_time','').strip(); end=request.form.get('end_time','').strip()
        subject_id=request.form.get('subject_id',type=int); teacher_id=request.form.get('teacher_id',type=int)
        session_type=request.form.get('session_type','Lecture').strip(); batch=request.form.get('batch','').strip(); room=request.form.get('room','').strip()
        if year not in YEAR_OPTIONS or day not in days or not start or not end or not subject_id or not teacher_id or session_type not in ('Lecture','Practical'):
            flash('Please fill all timetable fields correctly.','error')
        elif session_type=='Practical' and batch not in ('Batch A','Batch B','Batch C','Batch D'):
            flash('Select a valid practical batch.','error')
        else:
            dept_row=db.execute('SELECT id FROM departments ORDER BY id LIMIT 1').fetchone(); dept=dept_row['id'] if dept_row else None
            db.execute("INSERT INTO timetable(department_id,year,day_of_week,lecture_no,start_time,end_time,subject_id,teacher_id,room,session_type,batch) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                       (dept,year,day,1,start,end,subject_id,teacher_id,room,session_type,batch if session_type=='Practical' else ''))
            db.commit(); flash('Timetable entry added successfully.','success')
    selected_year=request.args.get('year','').strip(); selected_day=request.args.get('day','').strip()
    # The timetable UI uses view=day|week. Keep the older week=1 style compatible too.
    view=request.args.get('view','').strip().lower()
    week_view=(view=='week' or request.args.get('week','').strip().lower() in ('1','true','yes','week','all'))
    if selected_day not in days: selected_day=date.today().strftime('%A')
    where=[]; params=[]
    if role=='teacher': where.append('t.teacher_id=?'); params.append(session['user_id'])
    elif role=='student':
        st=current_student(db); selected_year=st['year'] if st else selected_year
        where += ['t.year=?',"(COALESCE(t.session_type,'Lecture')='Lecture' OR COALESCE(t.batch,'')=?)"]; params += [selected_year, st['batch'] if st else '']
    elif role=='hod': where.append('t.department_id=(SELECT department_id FROM users WHERE id=?)'); params.append(session['user_id'])
    if selected_year and role in ('admin','hod'): where.append('t.year=?'); params.append(selected_year)
    if not week_view:
        where.append('t.day_of_week=?'); params.append(selected_day)
    rows=db.execute(f"""SELECT t.*,COALESCE(t.session_type,'Lecture') session_type,COALESCE(t.batch,'') batch,
        COALESCE(sub.code,'') subject_code,COALESCE(sub.name,'') subject_name,COALESCE(u.full_name,u.username,'') teacher_name
        FROM timetable t LEFT JOIN subjects sub ON sub.id=t.subject_id LEFT JOIN users u ON u.id=t.teacher_id
        WHERE {' AND '.join(where)} ORDER BY CASE t.day_of_week WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3 WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 WHEN 'Saturday' THEN 6 ELSE 7 END, t.year,t.start_time,t.lecture_no,t.id""",tuple(params)).fetchall()
    subjects=db.execute('SELECT id,code,name,year FROM subjects ORDER BY year,code').fetchall(); teachers=all_teacher_rows(db)
    db.close()
    return render_template('timetable.html',rows=rows,subjects=subjects,teachers=teachers,years=YEAR_OPTIONS,selected_year=selected_year,selected_day=selected_day,days=days,week_view=week_view,batch_options=['Batch A','Batch B','Batch C','Batch D'])


@app.route('/timetable/export.pdf')
@login_required
def export_timetable_pdf():
    """Export the timetable as a PDF for the selected day or full Monday-Saturday week."""
    db = get_db()
    role = session.get('role')
    days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']

    # New UI uses view=day|week. Keep week=1 compatible with older links.
    view = request.args.get('view', '').strip().lower()
    week = view == 'week' or request.args.get('week', '').strip().lower() in ('1','true','yes','week','all')

    selected_day = request.args.get('day', '').strip()
    if selected_day not in days:
        selected_day = date.today().strftime('%A')
        if selected_day not in days:
            selected_day = 'Monday'

    selected_year = request.args.get('year', '').strip()
    where = []
    params = []

    if role == 'teacher':
        where.append('t.teacher_id=?')
        params.append(session['user_id'])
    elif role == 'student':
        st = current_student(db)
        if not st:
            db.close()
            return Response('Student record not found.', status=404)
        # Student timetable is always limited to the logged-in student's year and batch.
        selected_year = st['year'] or selected_year
        where += [
            't.year=?',
            "(COALESCE(t.session_type,'Lecture')='Lecture' OR COALESCE(t.batch,'')=?)"
        ]
        params += [selected_year, st['batch'] or '']
    elif role == 'hod':
        where.append('t.department_id=(SELECT department_id FROM users WHERE id=?)')
        params.append(session['user_id'])
    elif role == 'admin':
        if selected_year and selected_year in YEAR_OPTIONS:
            where.append('t.year=?')
            params.append(selected_year)
    else:
        db.close()
        return Response('Not authorized.', status=403)

    if not week:
        where.append('t.day_of_week=?')
        params.append(selected_day)

    # Always have a valid WHERE clause because every non-admin role has a role filter,
    # while admin can legitimately export the complete timetable.
    where_sql = ' AND '.join(where) if where else '1=1'
    rows = db.execute(f"""
        SELECT t.*, COALESCE(t.session_type,'Lecture') session_type,
            COALESCE(t.batch,'') batch, COALESCE(sub.code,'') subject_code,
            COALESCE(sub.name,'') subject_name,
            COALESCE(u.full_name,u.username,'') teacher_name
        FROM timetable t
        LEFT JOIN subjects sub ON sub.id=t.subject_id
        LEFT JOIN users u ON u.id=t.teacher_id
        WHERE {where_sql}
        ORDER BY CASE t.day_of_week
            WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3
            WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 WHEN 'Saturday' THEN 6 ELSE 7 END,
            t.start_time,t.lecture_no,t.id
    """, tuple(params)).fetchall()
    db.close()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=7*mm, leftMargin=7*mm,
        topMargin=7*mm, bottomMargin=8*mm
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle('TTCollegeTitleFixed', parent=styles['Title'], alignment=TA_CENTER, fontSize=17, leading=20, spaceAfter=1)
    sub = ParagraphStyle('TTSubFixed', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10, leading=12, spaceAfter=1)
    meta = ParagraphStyle('TTMetaFixed', parent=styles['Normal'], alignment=TA_CENTER, fontSize=8.5, leading=10.5)
    cell = ParagraphStyle('TTCellFixed', parent=styles['Normal'], fontSize=6.8, leading=8, alignment=TA_LEFT)
    head = ParagraphStyle('TTHeadFixed', parent=cell, fontName='Helvetica-Bold', alignment=TA_CENTER)

    def safe(value):
        return xml_escape(str(value if value is not None else '—'))

    story = []
    logo_path = BASE_DIR / 'static' / LOGO_FILE
    logo = Image(str(logo_path), width=20*mm, height=20*mm) if logo_path.exists() else ''
    class_text = selected_year if selected_year else 'All Years'
    view_text = 'Full Week (Monday – Saturday)' if week else f'Day: {selected_day}'
    header_text = [
        Paragraph(safe(COLLEGE_NAME), title),
        Paragraph('Academic Timetable', sub),
        Paragraph('Affiliated by Dr. Babasaheb Ambedkar Technological University, Lonere, Raigad', meta),
        Paragraph(f'<b>{safe(view_text)}</b> &nbsp;&nbsp; <b>Year:</b> {safe(class_text)}', meta)
    ]
    hdr = Table([[logo, header_text]], colWidths=[24*mm, 262*mm])
    hdr.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
        ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0)
    ]))
    story += [hdr, Spacer(1, 3*mm)]

    columns = ['Day','Year','Time','Session','Batch','Subject','Teacher','Room / Lab'] if week else ['Year','Time','Session','Batch','Subject','Teacher','Room / Lab']
    data = [[Paragraph(safe(c), head) for c in columns]]
    for r in rows:
        subject = f"{r['subject_code']} — {r['subject_name']}" if r['subject_code'] else (r['subject_name'] or '—')
        values = ([r['day_of_week']] if week else []) + [
            r['year'] or '—',
            f"{format_time_12(r['start_time'])} – {format_time_12(r['end_time'])}",
            r['session_type'] or 'Lecture',
            r['batch'] or 'All',
            subject,
            r['teacher_name'] or '—',
            r['room'] or '—'
        ]
        data.append([Paragraph(safe(v), cell) for v in values])

    if not rows:
        data.append([Paragraph('No timetable entries found for the selected filter.', cell)] + ['']*(len(columns)-1))

    if week:
        widths = [23*mm,17*mm,30*mm,23*mm,22*mm,63*mm,47*mm,38*mm]
    else:
        widths = [20*mm,33*mm,24*mm,22*mm,67*mm,48*mm,43*mm]

    table = Table(data, repeatRows=1, colWidths=widths)
    table.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0b3974')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('GRID',(0,0),(-1,-1),0.35,colors.HexColor('#b8c7d9')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
        ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('ALIGN',(0,1),(-1,-1),'CENTER'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f5f9fd')]),
    ]))
    story.append(table)
    story.append(Spacer(1, 3*mm))
    generated = datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%d-%m-%Y %I:%M %p')
    story.append(Paragraph(f'Generated on: {safe(generated)} IST &nbsp;&nbsp; | &nbsp;&nbsp; Y.N.P. College of Pharmacy', meta))
    doc.build(story)
    buffer.seek(0)

    suffix = 'full_week' if week else selected_day.lower()
    filename = f"timetable_{(selected_year or 'all').replace(' ','_')}_{suffix}.pdf"
    return Response(buffer.getvalue(), mimetype='application/pdf', headers={
        'Content-Disposition': f'attachment; filename={filename}'
    })


@app.post('/timetable/delete/<int:timetable_id>')
@login_required
@admin_required
def delete_timetable(timetable_id):
    db=get_db(); db.execute('DELETE FROM timetable WHERE id=?',(timetable_id,)); db.commit(); db.close(); flash('Timetable entry deleted.','success'); return redirect(url_for('timetable'))

@app.route("/dashboard")
@login_required
def dashboard():
    if session.get("role") == "student":
        return redirect(url_for("student_dashboard"))
    if session.get("role") == "hod":
        return redirect(url_for("hod_dashboard"))
    db = get_db()
    today = date.today().isoformat()
    total = db.execute("SELECT COUNT(*) c FROM students").fetchone()["c"]
    pending = db.execute("SELECT COUNT(*) c FROM users WHERE role='teacher' AND approved=0").fetchone()["c"] if session.get("role") == "admin" else 0

    subject_filter = "" if session.get("role") == "admin" else " AND a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)"
    base_params = [] if session.get("role") == "admin" else [session["user_id"]]
    today_rows = db.execute(f"""SELECT a.student_id,a.status,a.lecture_no,s.year
        FROM attendance a JOIN students s ON s.id=a.student_id
        WHERE a.attendance_date=?{subject_filter}""", [today] + base_params).fetchall()
    present_marks = sum(1 for r in today_rows if r["status"] == "Present")
    total_marks = len(today_rows)
    absent_marks = total_marks - present_marks
    today_pct = round(100.0 * present_marks / total_marks, 1) if total_marks else 0

    year_stats=[]
    for y in YEAR_OPTIONS:
        yr_students = db.execute("SELECT COUNT(*) c FROM students WHERE year=?", (y,)).fetchone()["c"]
        rows = [r for r in today_rows if r["year"] == y]
        p = sum(1 for r in rows if r["status"] == "Present")
        a = sum(1 for r in rows if r["status"] == "Absent")
        marked_students = {r["student_id"] for r in rows}
        # A student is counted as present if present for every marked lecture; otherwise absent.
        per_student = {}
        for r in rows:
            per_student.setdefault(r["student_id"], []).append(r["status"])
        present_students = sum(1 for sid, sts in per_student.items() if sts and all(x == "Present" for x in sts))
        absent_students = sum(1 for sid, sts in per_student.items() if any(x == "Absent" for x in sts))
        pct = round(100.0 * p / (p+a), 1) if (p+a) else 0
        year_stats.append({"year":y,"total":yr_students,"present":present_students,"absent":absent_students,"pct":pct,"marked":len(marked_students),"present_marks":p,"absent_marks":a})

    threshold_row = db.execute("SELECT value FROM app_settings WHERE key='attendance_threshold'").fetchone()
    try: threshold = float(threshold_row["value"] if threshold_row and threshold_row["value"] else 75)
    except Exception: threshold = 75
    low = db.execute("""SELECT s.prn,s.name,s.course,s.year,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id
        GROUP BY s.id HAVING (CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END) < ?
        ORDER BY pct ASC, s.name LIMIT 10""", (threshold,)).fetchall()
    # Overall attendance uses lecture marks, so multiple lectures are counted correctly.
    avg = db.execute("""SELECT COALESCE(AVG(pct),0) avg_pct FROM (
        SELECT s.id, CASE WHEN COUNT(a.id)=0 THEN 0 ELSE 100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id) END pct
        FROM students s LEFT JOIN attendance a ON a.student_id=s.id GROUP BY s.id)""").fetchone()["avg_pct"]
    if session.get("role") == "admin":
        subject_count = db.execute("SELECT COUNT(*) c FROM subjects").fetchone()["c"]
        teacher_count = db.execute("SELECT COUNT(*) c FROM users WHERE role='teacher' AND approved=1").fetchone()["c"]
    else:
        subject_count = db.execute("SELECT COUNT(*) c FROM subject_teachers WHERE teacher_id=?", (session["user_id"],)).fetchone()["c"]
        teacher_count = 0

    recent_attendance = db.execute("""SELECT a.attendance_date, sub.code, sub.name subject_name, s.year,
        COUNT(a.id) total,
        SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) absent
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id JOIN students s ON s.id=a.student_id
        WHERE 1=1
        """ + (" AND a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?) " if session.get('role')=='teacher' else "") + """GROUP BY a.attendance_date, sub.id, sub.code, sub.name, s.year
        ORDER BY a.attendance_date DESC LIMIT 6""", ([session["user_id"]] if session.get('role')=='teacher' else [])).fetchall()
    db.close()
    return render_template("dashboard.html", total=total, present=present_marks, absent=absent_marks,
        avg=round(avg or 0,1), low=low, pending=pending, year_stats=year_stats, today=today, today_pct=today_pct, today_marks=total_marks, threshold=threshold, subject_count=subject_count, teacher_count=teacher_count, recent_attendance=recent_attendance)


@app.route("/student/dashboard")
@student_required
def student_dashboard():
    db = get_db()
    student = current_student(db)
    if not student:
        db.close()
        session.clear()
        flash("Student account is not linked correctly. Contact the administrator.", "error")
        return redirect(url_for("login"))

    today = date.today().isoformat()
    selected_date = request.args.get("date", "").strip()
    selected_subject = request.args.get("subject", "").strip()
    timetable_days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    selected_timetable_day = request.args.get("timetable_day", "").strip()
    if selected_timetable_day not in timetable_days:
        selected_timetable_day = date.today().strftime('%A')
        if selected_timetable_day not in timetable_days:
            selected_timetable_day = 'Monday'

    today_rows = db.execute("""SELECT sub.code, sub.name subject_name, a.status,
        COALESCE(a.lecture_no,1) lecture_no, COALESCE(a.lecture_time,'') lecture_time
        FROM subjects sub LEFT JOIN attendance a
        ON a.subject_id=sub.id AND a.student_id=? AND a.attendance_date=?
        WHERE sub.department_id=? AND sub.year=?
        ORDER BY sub.code, a.lecture_no""", (student["id"], today, student["department_id"], student["year"])).fetchall()

    subjects = db.execute("SELECT id, code, name FROM subjects WHERE department_id=? AND year=? ORDER BY code", (student["department_id"], student["year"])).fetchall()

    history_sql = """SELECT a.attendance_date, sub.code, sub.name subject_name,
        a.status, COALESCE(a.lecture_no,1) lecture_no, COALESCE(a.lecture_time,'') lecture_time
        FROM attendance a JOIN subjects sub ON sub.id=a.subject_id
        WHERE a.student_id=?"""
    history_params = [student["id"]]
    if selected_date:
        history_sql += " AND a.attendance_date=?"
        history_params.append(selected_date)
    if selected_subject:
        history_sql += " AND a.subject_id=?"
        history_params.append(selected_subject)
    history_sql += " ORDER BY a.attendance_date DESC, sub.code, a.lecture_no"
    history = db.execute(history_sql, tuple(history_params)).fetchall()

    summary = db.execute("""SELECT sub.code, sub.name subject_name,
        COUNT(a.id) total,
        COALESCE(SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END),0) present,
        COALESCE(SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END),0) absent,
        CASE WHEN COUNT(a.id)=0 THEN 0 ELSE ROUND(100.0*SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END)/COUNT(a.id),1) END pct
        FROM subjects sub LEFT JOIN attendance a ON a.subject_id=sub.id AND a.student_id=?
        WHERE sub.department_id=? AND sub.year=? GROUP BY sub.id ORDER BY sub.code""", (student["id"], student["department_id"], student["year"])).fetchall()

    total = db.execute("SELECT COUNT(*) FROM attendance WHERE student_id=?", (student["id"],)).fetchone()[0]
    present = db.execute("SELECT COUNT(*) FROM attendance WHERE student_id=? AND status='Present'", (student["id"],)).fetchone()[0]
    overall = round((present / total * 100), 1) if total else 0

    # Student dashboard timetable is strictly limited to the logged-in
    # student's current year and practical batch. Lectures are shown for the
    # year, while practicals are shown only for the student's batch.
    my_timetable = get_timetable_rows(
        db, role='student', user_id=session.get('user_id'),
        student=student, day_name=selected_timetable_day
    )
    db.close()
    return render_template("student_dashboard.html", student=student, today_rows=today_rows,
                           summary=summary, overall=overall, history=history, subjects=subjects,
                           selected_date=selected_date, selected_subject=selected_subject, today=today,
                           my_timetable=my_timetable, timetable_days=timetable_days,
                           selected_timetable_day=selected_timetable_day)


@app.route("/student/calendar")
@student_required
def student_calendar():
    db=get_db(); student=current_student(db)
    month=request.args.get("month",date.today().strftime("%Y-%m"))
    if not re.match(r"^\d{4}-\d{2}$",month): month=date.today().strftime("%Y-%m")
    start=f"{month}-01"
    y,m=map(int,month.split("-")); end=f"{y+1:04d}-01-01" if m==12 else f"{y:04d}-{m+1:02d}-01"
    rows=db.execute("""SELECT a.attendance_date, SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present, COUNT(a.id) total
        FROM attendance a WHERE a.student_id=? AND a.attendance_date>=? AND a.attendance_date<? GROUP BY a.attendance_date ORDER BY a.attendance_date""",(student["id"],start,end)).fetchall()
    db.close(); return render_template("student_calendar.html",student=student,month=month,days=rows)

@app.route("/students")
@login_required
@admin_required
def students():
    q = request.args.get("q", "").strip()
    selected_year = request.args.get("year", "").strip()
    db = get_db()

    # Four year-wise student cards. Counts come directly from the students table.
    year_counts = {}
    for year in YEAR_OPTIONS:
        year_counts[year] = db.execute(
            "SELECT COUNT(*) FROM students WHERE year=?", (year,)
        ).fetchone()[0]

    conditions = []
    params = []
    if selected_year in YEAR_OPTIONS:
        conditions.append("s.year=?")
        params.append(selected_year)
    if q:
        conditions.append("(s.prn LIKE ? OR s.name LIKE ? OR s.course LIKE ?)")
        params.extend([f"%{q}%", f"%{q}%", f"%{q}%"])

    where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
    rows = db.execute(f"""SELECT s.*, u.username student_username
        FROM students s LEFT JOIN users u ON u.student_id=s.id AND u.role='student'
        {where}
        ORDER BY s.prn, s.name""", tuple(params)).fetchall()

    # Arrange students by complete PRN in ascending numerical order.
    # Numeric PRNs are compared as numbers, so ...3009, ...3010, ...3011
    # comes before ...3020, ...3024, etc. Non-numeric PRNs are kept after
    # numeric PRNs and sorted alphabetically.
    def _prn_sort_key(row):
        value = str(row["prn"] or "").strip()
        if value.isdigit():
            return (0, int(value), "")
        return (1, 0, value.lower())

    rows = sorted(rows, key=_prn_sort_key)
    db.close()
    return render_template("students.html", students=rows, q=q, selected_year=selected_year,
                           year_options=YEAR_OPTIONS, year_counts=year_counts)


@app.route("/students/add", methods=["GET", "POST"])
@login_required
@admin_required
def add_student():
    db = get_db()
    departments = db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall()
    if request.method == "POST":
        data = [request.form.get(k, "").strip() for k in ("prn", "name", "course", "year")]
        batch = request.form.get("batch", "").strip()
        dept_id = departments[0]["id"] if departments else None
        if not data[0] or not data[1] or data[3] not in YEAR_OPTIONS or not dept_id:
            db.close(); flash("PRN, name and year are required.", "error")
            return render_template("student_form.html", student=None, title="Add Student", year_options=YEAR_OPTIONS, batch_options=["Batch A","Batch B","Batch C","Batch D"])
        try:
            db.execute("INSERT INTO students(roll_no,prn,name,course,year,division,department_id,batch) VALUES(?,?,?,?,?,?,?,?)", (data[0], data[0], data[1], data[2] or "B.Pharm", data[3], "A", dept_id, batch))
            db.commit(); flash("Student added successfully.", "success")
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)): raise
            flash("PRN already exists.", "error")
        db.close(); return redirect(url_for("students"))
    db.close(); return render_template("student_form.html", student=None, title="Add Student", year_options=YEAR_OPTIONS, batch_options=["Batch A","Batch B","Batch C","Batch D"])


@app.route("/students/edit/<int:student_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_student(student_id):
    db = get_db(); student = db.execute("SELECT * FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        db.close(); flash("Student not found.", "error"); return redirect(url_for("students"))
    if request.method == "POST":
        data = [request.form.get(k, "").strip() for k in ("prn", "name", "course", "year")]
        batch = request.form.get("batch", "").strip()
        try:
            dept_id = student["department_id"]
            if not dept_id:
                dept_row = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()
                dept_id = dept_row["id"] if dept_row else None
            if not data[0] or not data[1] or data[3] not in YEAR_OPTIONS:
                raise ValueError("PRN, name and valid year are required.")
            db.execute("UPDATE students SET roll_no=?,prn=?,name=?,course=?,year=?,department_id=?,batch=? WHERE id=?", (data[0], data[0], data[1], data[2] or "B.Pharm", data[3], dept_id, batch, student_id))
            db.execute("UPDATE users SET full_name=? WHERE student_id=? AND role='student'", (data[1], student_id))
            db.commit(); flash("Student updated successfully.", "success"); db.close(); return redirect(url_for("students"))
        except ValueError as exc:
            db.close(); flash(str(exc), "error")
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)):
                raise
            flash("PRN already exists.", "error")
    departments=db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall(); db.close(); return render_template("student_form.html", student=student, title="Edit Student", year_options=YEAR_OPTIONS, batch_options=["Batch A","Batch B","Batch C","Batch D"])


@app.route("/students/delete/<int:student_id>", methods=["POST"])
@login_required
@admin_required
def delete_student(student_id):
    db = get_db(); db.execute("DELETE FROM students WHERE id=?", (student_id,)); db.commit(); db.close()
    flash("Student and linked attendance/account were deleted.", "success"); return redirect(url_for("students"))


@app.route("/students/<int:student_id>/toggle-block", methods=["POST"])
@login_required
@admin_required
def toggle_block_student(student_id):
    db = get_db()
    student = db.execute("SELECT id, name, blocked FROM students WHERE id=?", (student_id,)).fetchone()
    if not student:
        db.close()
        flash("Student not found.", "error")
        return redirect(url_for("students"))
    new_status = 0 if int(student["blocked"] or 0) else 1
    db.execute("UPDATE students SET blocked=? WHERE id=?", (new_status, student_id))
    db.commit()
    db.close()
    if new_status:
        flash(f"{student['name']} has been blocked. Student login and registration are disabled.", "success")
    else:
        flash(f"{student['name']} has been unblocked. Student can log in and register again.", "success")
    return redirect(url_for("students"))


@app.route("/teachers")
@login_required
@admin_required
def teachers():
    db = get_db()
    rows = db.execute("SELECT * FROM users WHERE role='teacher' ORDER BY approved ASC, full_name, username").fetchall()
    subjects = db.execute("SELECT * FROM subjects ORDER BY year, code").fetchall()
    assignments = db.execute("SELECT teacher_id, subject_id FROM subject_teachers").fetchall()
    assigned = {(r["teacher_id"], r["subject_id"]) for r in assignments}
    db.close()
    return render_template("teachers.html", teachers=rows, subjects=subjects, assigned=assigned, years=YEAR_OPTIONS)


@app.route("/teachers/<int:teacher_id>/approve", methods=["POST"])
@login_required
@admin_required
def approve_teacher(teacher_id):
    db = get_db(); db.execute("UPDATE users SET approved=1 WHERE id=? AND role='teacher'", (teacher_id,)); db.commit(); db.close()
    flash("Teacher account approved.", "success"); return redirect(url_for("teachers"))


@app.route("/teachers/<int:teacher_id>/toggle", methods=["POST"])
@login_required
@admin_required
def toggle_teacher(teacher_id):
    db = get_db(); db.execute("UPDATE users SET approved=CASE approved WHEN 1 THEN 0 ELSE 1 END WHERE id=? AND role='teacher'", (teacher_id,)); db.commit(); db.close()
    flash("Teacher account status updated.", "success"); return redirect(url_for("teachers"))


@app.route("/teachers/<int:teacher_id>/delete", methods=["POST"])
@login_required
@admin_required
def delete_teacher(teacher_id):
    db = get_db()
    teacher = db.execute("SELECT full_name,username FROM users WHERE id=? AND role='teacher'", (teacher_id,)).fetchone()
    if not teacher:
        db.close(); flash("Teacher not found.", "error"); return redirect(url_for("teachers"))
    db.execute("DELETE FROM users WHERE id=? AND role='teacher'", (teacher_id,))
    db.commit(); db.close()
    flash(f"Teacher {teacher['full_name'] or teacher['username']} removed successfully.", "success")
    return redirect(url_for("teachers"))


@app.route("/teachers/<int:teacher_id>/assign", methods=["POST"])
@login_required
@admin_required
def assign_subjects(teacher_id):
    db = get_db()
    db.execute("DELETE FROM subject_teachers WHERE teacher_id=?", (teacher_id,))
    subject_ids = request.form.getlist("subject_ids")
    valid = db.execute("SELECT id FROM subjects").fetchall()
    valid_ids = {str(r["id"]) for r in valid}
    for sid in subject_ids:
        if sid in valid_ids:
            db.execute("INSERT OR IGNORE INTO subject_teachers(teacher_id,subject_id) VALUES(?,?)", (teacher_id, int(sid)))
    db.commit(); db.close()
    flash("Teacher subject assignments saved year-wise.", "success"); return redirect(url_for("teachers"))


@app.route("/subjects")
@login_required
@admin_required
def subjects():
    db = get_db()
    rows = db.execute("SELECT s.*, d.name department_name FROM subjects s LEFT JOIN departments d ON d.id=s.department_id ORDER BY s.year, s.code").fetchall()
    db.close()
    grouped = {y: [] for y in YEAR_OPTIONS}
    for row in rows:
        grouped.setdefault(row["year"] or "1st Year", []).append(row)
    return render_template("subjects.html", subjects=rows, grouped=grouped, years=YEAR_OPTIONS)


@app.route("/subjects/add", methods=["GET", "POST"])
@login_required
@admin_required
def add_subject():
    db = get_db()
    departments = db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall()
    if request.method == "POST":
        code = request.form.get("code", "").strip().upper(); name = request.form.get("name", "").strip()
        year = request.form.get("year", "").strip()
        # Department is intentionally hidden from the UI. Keep the existing
        # internal department relationship by assigning the default department.
        default_dept_row = db.execute("SELECT id FROM departments ORDER BY id LIMIT 1").fetchone()
        dept_id = default_dept_row["id"] if default_dept_row else None
        if year not in YEAR_OPTIONS: year = ""
        if not code or not name or not year or not dept_id:
            db.close(); flash("Subject code, name and year are required.", "error")
            return render_template("subject_form.html", subject=None, title="Add Subject", departments=departments, years=YEAR_OPTIONS)
        try:
            db.execute("INSERT INTO subjects(code,name,year,department_id) VALUES(?,?,?,?)", (code, name, year, dept_id)); db.commit(); flash("Subject added.", "success")
        except Exception as exc:
            if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)): raise
            flash("Subject code already exists.", "error")
        db.close(); return redirect(url_for("subjects"))
    db.close(); return render_template("subject_form.html", subject=None, title="Add Subject", departments=departments, years=YEAR_OPTIONS)


@app.route("/subjects/edit/<int:subject_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_subject(subject_id):
    db = get_db(); subject = db.execute("SELECT * FROM subjects WHERE id=?", (subject_id,)).fetchone()
    if not subject:
        db.close(); flash("Subject not found.", "error"); return redirect(url_for("subjects"))
    if request.method == "POST":
        code = request.form.get("code", "").strip().upper()
        name = request.form.get("name", "").strip()
        year = request.form.get("year", "").strip()
        # Preserve the existing internal department relationship; users no longer select it.
        dept_id = subject["department_id"]
        if not code or not name or year not in YEAR_OPTIONS:
            flash("Subject code, name and year are required.", "error")
        else:
            try:
                db.execute("UPDATE subjects SET code=?,name=?,year=?,department_id=? WHERE id=?", (code, name, year, dept_id, subject_id)); db.commit(); flash("Subject updated.", "success"); db.close(); return redirect(url_for("subjects"))
            except Exception as exc:
                if not isinstance(exc, sqlite3.IntegrityError) and not (POSTGRES_AVAILABLE and isinstance(exc, psycopg2.IntegrityError)): raise
                flash("Subject code already exists.", "error")
    departments=db.execute("SELECT * FROM departments WHERE active=1 ORDER BY name").fetchall(); db.close(); return render_template("subject_form.html", subject=subject, title="Edit Subject", departments=departments, years=YEAR_OPTIONS)


@app.route("/subjects/delete/<int:subject_id>", methods=["POST"])
@login_required
@admin_required
def delete_subject(subject_id):
    db = get_db(); db.execute("DELETE FROM subjects WHERE id=?", (subject_id,)); db.commit(); db.close()
    flash("Subject deleted.", "success"); return redirect(url_for("subjects"))


def _prn_sort_key(row):
    value = str(row["prn"] or "").strip()
    if value.isdigit():
        return (0, int(value), "")
    return (1, 0, value.lower())


@app.route("/attendance-history")
@login_required
@staff_required
def attendance_history():
    db = get_db()
    years = YEAR_OPTIONS[:]
    year = request.args.get("year", "").strip()
    session_type = request.args.get("session_type", "").strip()
    batch = request.args.get("batch", "").strip()
    subject_id = request.args.get("subject_id", type=int)
    from_date = request.args.get("from_date", "").strip()
    to_date = request.args.get("to_date", "").strip()

    subjects = db.execute("SELECT id,code,name,year FROM subjects ORDER BY year,code").fetchall()
    conditions = ["1=1"]
    params = []
    if year in years:
        conditions.append("s.year=?"); params.append(year)
    if session_type in ("Lecture", "Practical"):
        conditions.append("a.session_type=?"); params.append(session_type)
    if batch:
        conditions.append("a.batch=?"); params.append(batch)
    if subject_id:
        conditions.append("a.subject_id=?"); params.append(subject_id)
    if from_date:
        conditions.append("a.attendance_date>=?"); params.append(from_date)
    if to_date:
        conditions.append("a.attendance_date<=?"); params.append(to_date)
    if session.get("role") == "teacher":
        conditions.append("a.subject_id IN (SELECT subject_id FROM subject_teachers WHERE teacher_id=?)")
        params.append(session["user_id"])

    rows = db.execute(f"""SELECT a.attendance_date,a.session_type,a.batch,a.lecture_no,a.lecture_time,
        sub.code,sub.name subject_name,COUNT(a.id) total,
        SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) absent
        FROM attendance a
        JOIN students s ON s.id=a.student_id
        JOIN subjects sub ON sub.id=a.subject_id
        WHERE {' AND '.join(conditions)}
        GROUP BY a.attendance_date,a.session_type,a.batch,a.lecture_no,a.lecture_time,sub.id,sub.code,sub.name
        ORDER BY a.attendance_date DESC,a.lecture_time DESC,sub.code""", params).fetchall()
    db.close()
    return render_template("attendance_history.html", rows=rows, years=years, year=year,
        session_type=session_type, batch=batch, subject_id=subject_id, from_date=from_date,
        to_date=to_date, subjects=subjects, batch_options=["Batch A","Batch B","Batch C","Batch D"])


@app.route("/qr-attendance", methods=["GET", "POST"])
@login_required
@staff_required
def qr_attendance():
    return take_attendance()


@app.route("/attendance", methods=["GET", "POST"])
@login_required
@staff_required
def take_attendance():
    db=get_db(); years=YEAR_OPTIONS[:]
    selected_year=request.args.get('year') or request.form.get('year') or '1st Year'
    if selected_year not in years:
        selected_year='1st Year'
    session_type=(request.args.get('session_type') or request.form.get('session_type') or 'Lecture').strip()
    if session_type not in ('Lecture','Practical'): session_type='Lecture'
    batch=(request.args.get('batch') or request.form.get('batch') or '').strip()
    subject_id=request.args.get('subject_id',type=int) or request.form.get('subject_id',type=int)
    ad=request.args.get('attendance_date') or request.form.get('attendance_date') or date.today().isoformat()
    try: date.fromisoformat(ad)
    except ValueError: ad=date.today().isoformat()
    tid=request.args.get('timetable_id',type=int) or request.form.get('timetable_id',type=int)
    start_time=request.args.get('start_time') or request.form.get('start_time',''); end_time=request.args.get('end_time') or request.form.get('end_time','')
    teacher_status=request.args.get('teacher_status') or request.form.get('teacher_status','Scheduled teacher present'); conducted_by=request.args.get('conducted_by',type=int) or request.form.get('conducted_by',type=int)
    subjects=db.execute("SELECT id,code,name,year FROM subjects WHERE year=? ORDER BY code",(selected_year,)).fetchall()
    if subject_id and not any(s['id']==subject_id for s in subjects): subject_id=None
    if session_type=='Practical' and batch not in ('Batch A','Batch B','Batch C','Batch D'): batch=''
    day_name=date.fromisoformat(ad).strftime('%A')
    candidates=[]
    if subject_id:
        try: candidates=timetable_candidates(db,selected_year,session_type,batch,subject_id,day_name)
        except Exception: candidates=[]
    teachers=all_teacher_rows(db)
    if request.method=='POST':
        tr=db.execute('SELECT * FROM timetable WHERE id=?',(tid,)).fetchone() if tid else None
        if not subject_id or not any(s['id']==subject_id for s in subjects) or not start_time or not end_time:
            flash('Select valid timetable, subject and start/end time.','error')
        elif session_type=='Practical' and batch not in ('Batch A','Batch B','Batch C','Batch D'):
            flash('Select a valid practical batch.','error')
        else:
            scheduled=tr['teacher_id'] if tr else None
            if teacher_status=='Scheduled teacher present': conducted_by=scheduled or session['user_id']
            elif not conducted_by: conducted_by=session['user_id']
            token=secrets.token_urlsafe(24); exp=datetime.utcnow().replace(microsecond=0)+__import__('datetime').timedelta(minutes=10)
            db.execute("""INSERT INTO qr_sessions(token,subject_id,year,attendance_date,lecture_no,lecture_time,start_time,end_time,session_type,batch,teacher_id,scheduled_teacher_id,conducted_by_teacher_id,teacher_status,expires_at,active)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1)""",(token,subject_id,selected_year,ad,1,f'{start_time}-{end_time}',start_time,end_time,session_type,batch if session_type=='Practical' else '',session['user_id'],scheduled,conducted_by,teacher_status,exp.isoformat()))
            db.commit(); db.close(); return redirect(url_for('qr_attendance',token=token))
    token=request.args.get('token',''); qr=None; qr_session=None
    if token:
        qr_session=db.execute('SELECT q.*,s.code,s.name subject_name FROM qr_sessions q JOIN subjects s ON s.id=q.subject_id WHERE q.token=? AND q.active=1',(token,)).fetchone()
        if qr_session and datetime.fromisoformat(qr_session['expires_at'])<datetime.utcnow(): qr_session=None
        if qr_session and QRCODE_AVAILABLE:
            import base64
            img=qrcode.make(request.url_root.rstrip('/')+url_for('qr_scan',token=token)); bio=io.BytesIO(); img.save(bio,format='PNG'); qr='data:image/png;base64,'+base64.b64encode(bio.getvalue()).decode()
    db.close(); return render_template('qr_attendance.html',subjects=subjects,years=years,today=date.today().isoformat(),qr=qr,qr_session=qr_session,token=token,batch=batch,session_type=session_type,candidates=candidates,teachers=teachers,selected_year=selected_year,day_name=day_name)


@app.route('/qr/scan',methods=['GET','POST'])
@login_required
def qr_scan():
    if session.get('role')!='student': flash('Only students can scan attendance QR codes.','error'); return redirect(url_for('dashboard'))
    token=request.args.get('token') or request.form.get('token',''); db=get_db()
    row=db.execute("""SELECT q.*,s.code,s.name subject_name FROM qr_sessions q JOIN subjects s ON s.id=q.subject_id WHERE q.token=? AND q.active=1""",(token,)).fetchone()
    if not row or datetime.fromisoformat(row['expires_at'])<datetime.utcnow(): db.close(); return render_template('qr_scan.html',valid=False)
    st=current_student(db)
    if st['year']!=row['year'] or (row['session_type']=='Practical' and st['batch']!=row['batch']) or (st['department_id'] and db.execute('SELECT department_id FROM subjects WHERE id=?',(row['subject_id'],)).fetchone()['department_id']!=st['department_id']): db.close(); return render_template('qr_scan.html',valid=False,reason='This QR is not for your class or practical batch.')
    if request.method=='POST':
        db.execute("""INSERT INTO attendance(student_id,subject_id,attendance_date,lecture_no,lecture_time,session_type,batch,status,marked_by) VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(student_id,subject_id,attendance_date,lecture_no) DO UPDATE SET lecture_time=excluded.lecture_time,status=excluded.status,marked_by=excluded.marked_by""",
                   (st['id'],row['subject_id'],row['attendance_date'],row['lecture_no'],row['lecture_time'],row['session_type'],row['batch'],'Present',row['teacher_id']))
        notify_user(db, session['user_id'], "Attendance Marked", f"You were marked Present for {row['subject_name']} (Lecture {row['lecture_no']}).", "/student/dashboard")
        create_low_attendance_notifications(db)
        db.commit(); db.close(); return render_template('qr_scan.html',valid=True,done=True,session_info=row)
    db.close(); return render_template('qr_scan.html',valid=True,session_info=row)


@app.post('/qr/close/<token>')
@login_required
@staff_required
def close_qr(token):
    db=get_db(); db.execute("UPDATE qr_sessions SET active=0 WHERE token=? AND teacher_id=?",(token,session['user_id'])); db.commit(); db.close(); flash('QR attendance session closed.','success'); return redirect(url_for('qr_attendance'))


@app.get('/manifest.json')
def manifest():
    return {"name":COLLEGE_NAME+" Attendance","short_name":"YNP Attendance","start_url":"/","display":"standalone","background_color":"#ffffff","theme_color":"#1f4e79","icons":[{"src":url_for('static',filename=LOGO_FILE),"sizes":"192x192","type":"image/png"},{"src":url_for('static',filename=LOGO_FILE),"sizes":"512x512","type":"image/png"}]}


@app.get('/service-worker.js')
def service_worker():
    js="""self.addEventListener('install',e=>self.skipWaiting());self.addEventListener('activate',e=>self.clients.claim());self.addEventListener('notificationclick',e=>{e.notification.close();e.waitUntil(clients.openWindow('/notifications'))});"""
    return Response(js,mimetype='application/javascript')

# Initialize the database when the application is imported by Gunicorn/Render.
init_db()

if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "0") == "1",
            host=os.environ.get("HOST", "127.0.0.1"),
            port=int(os.environ.get("PORT", "5000")))
