import os
import json
import re
import io
import base64
import secrets
import unicodedata
import string
from datetime import datetime, timedelta, date

from flask import (Flask, render_template, request, jsonify, send_file, redirect, url_for, flash,
                   session, after_this_request, abort)

from flask_login import (LoginManager, login_user, logout_user, login_required,
                          current_user)

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors as rl_colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT

from flask_wtf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from models import (db, User, Subscription, Document, LicenseCode, Payment, Message,
                     PaymentRequest, AdminLog, PlanChangeHistory, FunnelEvent, GenerationError,
                     PLAN_PRICES, PAYMENT_METHODS, PASS_TYPES, Cohort,
                     TrainingSession, SessionMeeting, Enrollment, Attendance, SESSION_STATUSES,
                     create_trial_subscription, next_receipt_number)
from werkzeug.utils import secure_filename
from utils import build_sources_context
from translations import get_translations, TRANSLATIONS
from notifications import (check_and_send_expiry_reminders, send_email, send_activation_nudges,
                           smtp_configured, NUDGE_MAX_AGE_DAYS, NUDGE_STAGES)
from itsdangerous import URLSafeSerializer, BadSignature
from email_templates import nudge_email

SUPPORTED_LANGS = ("ar", "fr", "en")


def current_lang():
    return session.get("lang", "fr")


def tr(key, **kwargs):
    """Translate a key for the current session language, with optional .format() args."""
    text = get_translations(current_lang()).get(key, key)
    return text.format(**kwargs) if kwargs else text


def tr_lang(lang, key, **kwargs):
    """Translate a key for a given language (e.g. another user's), not the session's."""
    text = get_translations(lang if lang in TRANSLATIONS else "fr").get(key, key)
    return text.format(**kwargs) if kwargs else text

# ── APP SETUP ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
# Railway terminates HTTPS at its proxy: trust its X-Forwarded-* headers so
# external links (referral links, emails) are generated as https://.
from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-in-production")

# DATA_DIR points to a persistent volume in production (e.g. Railway mounts one at /data).
# Falls back to the app's own folder for local development.
DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(__file__))
os.makedirs(DATA_DIR, exist_ok=True)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(DATA_DIR, "eduprompt.db")
# The server handles several requests at once (threads): wait for a lock
# instead of failing, and use WAL so readers never block on a writer.
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"connect_args": {"timeout": 30}}
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024  # 15 MB uploads

RECEIPTS_DIR = os.path.join(DATA_DIR, "receipts")
os.makedirs(RECEIPTS_DIR, exist_ok=True)
ALLOWED_RECEIPT_EXT = {"png", "jpg", "jpeg", "webp", "gif", "pdf"}


def save_receipt_file(file_storage, user_id):
    """Save an uploaded payment receipt (screenshot/photo) to the persistent volume.
    Returns the stored filename, or None if no valid file was provided."""
    if not file_storage or not file_storage.filename:
        return None
    ext = file_storage.filename.rsplit(".", 1)[-1].lower() if "." in file_storage.filename else ""
    if ext not in ALLOWED_RECEIPT_EXT:
        return None
    fname = f"u{user_id}_{secrets.token_hex(6)}.{ext}"
    file_storage.save(os.path.join(RECEIPTS_DIR, secure_filename(fname)))
    return secure_filename(fname)

CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", "haithemcomputing@gmail.com")
# Public address used in emails sent by the scheduler (no request to read it from).
APP_URL = os.environ.get("APP_URL", "https://eduprompt-web-production.up.railway.app").rstrip("/")

db.init_app(app)
csrf = CSRFProtect(app)
limiter = Limiter(get_remote_address, app=app, default_limits=[], storage_uri="memory://")

login_manager = LoginManager(app)
login_manager.login_view = "login"


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


@app.before_request
def force_password_change():
    """Accounts created from a training import must choose their own password first."""
    if (current_user.is_authenticated and getattr(current_user, "must_change_password", False)
            and request.endpoint not in ("profile", "logout", "static", "set_lang", None)):
        if request.path.startswith("/api/"):
            return jsonify({"error": {"message": "password_change_required"}}), 403
        flash(tr("flash_must_change_password"), "warning")
        return redirect(url_for("profile"))


@app.before_request
def ensure_lang():
    if "lang" not in session:
        best = request.accept_languages.best_match(SUPPORTED_LANGS)
        session["lang"] = best or "fr"
    # Referral link (?ref=CODE): remember who invited this visitor.
    ref = request.args.get("ref", "").strip().upper()
    if ref and re.fullmatch(r"[A-Z0-9]{4,12}", ref) and session.get("ref") != ref:
        session["ref"] = ref
        session.setdefault("src", "referral")
    login_manager.login_message = tr("login_required_message")


@app.context_processor
def inject_i18n():
    lang = current_lang()
    return dict(lang=lang, t=get_translations(lang), is_rtl=(lang == "ar"), contact_email=CONTACT_EMAIL)


@app.context_processor
def inject_notifications():
    """Feeds the navbar notification bell for both teachers and admins."""
    if not current_user.is_authenticated:
        return dict(notif_unread_msgs=0, notif_expiring=0, notif_reminder=False, notif_payments=0,
                    notif_payments_overdue=0)

    if current_user.role == "admin":
        unread_msgs = Message.query.filter_by(sender="teacher", read_by_admin=False).count()
        now = datetime.utcnow()
        soon = now + timedelta(days=5)
        expiring = Subscription.query.filter(
            Subscription.status == "active",
            Subscription.expires_at.isnot(None),
            Subscription.expires_at >= now,
            Subscription.expires_at <= soon,
        ).count()
        pending_payments = PaymentRequest.query.filter_by(status="pending").count()
        overdue_cutoff = now - timedelta(hours=48)
        overdue_payments = PaymentRequest.query.filter(
            PaymentRequest.status == "pending",
            PaymentRequest.created_at <= overdue_cutoff,
        ).count()
        return dict(notif_unread_msgs=unread_msgs, notif_expiring=expiring, notif_reminder=False,
                    notif_payments=pending_payments, notif_payments_overdue=overdue_payments)
    else:
        unread_msgs = Message.query.filter_by(user_id=current_user.id, sender="admin",
                                               read_by_teacher=False).count()
        sub = current_user.subscription
        days_left = sub.days_until_expiry() if sub else None
        reminder = days_left is not None and 0 <= days_left <= 7
        return dict(notif_unread_msgs=unread_msgs, notif_expiring=0, notif_reminder=reminder,
                    notif_payments=0, notif_payments_overdue=0)


def paginate_list(items, page, per_page=15):
    """Manual pagination for plain Python lists (used where results are assembled
    in Python rather than as a single SQLAlchemy Query)."""
    total = len(items)
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = max(1, min(page, total_pages))
    start = (page - 1) * per_page
    return items[start:start + per_page], page, total_pages, total


@app.route("/set-lang/<lang>")
def set_lang(lang):
    if lang in SUPPORTED_LANGS:
        session["lang"] = lang
        if current_user.is_authenticated:
            current_user.preferred_lang = lang
            db.session.commit()
    return redirect(request.referrer or url_for("home"))


def _migrate_sqlite_schema():
    """Add newly-introduced columns to an existing SQLite db (no full migration
    framework needed for this small app). Safe to run on every startup."""
    import sqlite3
    db_path = os.path.join(DATA_DIR, "eduprompt.db")
    if not os.path.exists(db_path):
        return
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        table_additions = {
            "subscription": {
                "docs_used_today": "INTEGER DEFAULT 0",
                "last_reset_date": "DATE",
                "expiry_reminder_sent": "BOOLEAN DEFAULT 0",
                "last_reminder_stage": "INTEGER",
                "bonus_docs": "INTEGER DEFAULT 0",
                "bonus_expires_at": "DATETIME",
                "bonus_renewed": "BOOLEAN DEFAULT 0",
            },
            "user": {
                "gemini_api_key": "VARCHAR(255)",
                "preferred_lang": "VARCHAR(5) DEFAULT 'fr'",
                "is_active": "BOOLEAN DEFAULT 1",
                "profile": "VARCHAR(10) DEFAULT 'teacher'",
                "signup_source": "VARCHAR(40)",
                "referral_code": "VARCHAR(12)",
                "referred_by_id": "INTEGER",
                "referral_rewarded": "BOOLEAN DEFAULT 0",
                "referral_docs_earned": "INTEGER DEFAULT 0",
                "nudge_stage": "INTEGER DEFAULT 0",
                "email_optout": "BOOLEAN DEFAULT 0",
                "cohort_id": "INTEGER",
                "must_change_password": "BOOLEAN DEFAULT 0",
                "temp_password": "VARCHAR(40)",
                "credentials_pending": "BOOLEAN DEFAULT 0",
                "credentials_sent_at": "DATETIME",
                "discount_pct": "INTEGER",
                "discount_until": "DATETIME",
                "pass_reminder_sent": "BOOLEAN DEFAULT 0",
            },
            "document": {
                "content": "TEXT",
                "payload": "TEXT",
                "rating": "INTEGER",
                "rating_comment": "VARCHAR(500)",
                "rated_at": "DATETIME",
            },
            "training_session": {
                "min_attendance": "INTEGER DEFAULT 80",
                "mode": "VARCHAR(10) DEFAULT 'onsite'",
                "online_url": "VARCHAR(300)",
                "start_time": "VARCHAR(5)",
                "end_time": "VARCHAR(5)",
                "trainer": "VARCHAR(120)",
                "audience": "VARCHAR(200)",
                "programme": "TEXT",
                "prerequisites": "TEXT",
                "title_ar": "VARCHAR(160)",
                "description_ar": "TEXT",
                "location_ar": "VARCHAR(200)",
                "trainer_ar": "VARCHAR(120)",
                "audience_ar": "VARCHAR(200)",
                "programme_ar": "TEXT",
                "prerequisites_ar": "TEXT",
            },
            "payment_request": {
                "method": "VARCHAR(10)",
                "receipt_path": "VARCHAR(255)",
                "receipt_number": "VARCHAR(30)",
                "source": "VARCHAR(20) DEFAULT 'upgrade'",
            },
        }
        for table, additions in table_additions.items():
            cur.execute(f"PRAGMA table_info({table})")
            existing_cols = {row[1] for row in cur.fetchall()}
            if not existing_cols:
                continue  # table doesn't exist yet — db.create_all() will make it fresh
            for col, col_type in additions.items():
                if col not in existing_cols:
                    cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
                    print(f"✅ Migrated: added {table}.{col}")

        cur.execute("PRAGMA table_info(user)")
        if any(r[1] == "referral_code" for r in cur.fetchall()):
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS ix_user_referral_code_u ON user (referral_code)")

        # One-time data migration: the "premium" plan was renamed to "ultimate".
        for table in ("subscription", "payment_request", "license_code"):
            cur.execute(f"PRAGMA table_info({table})")
            if not cur.fetchall():
                continue
            cur.execute(f"UPDATE {table} SET plan='ultimate' WHERE plan='premium'")
            if cur.rowcount:
                print(f"✅ Migrated: {table}.plan premium→ultimate ({cur.rowcount} rows)")
        conn.commit()
    finally:
        conn.close()


with app.app_context():
    from sqlalchemy import event as _sa_event

    @_sa_event.listens_for(db.engine, "connect")
    def _sqlite_pragmas(dbapi_conn, _record):
        cur = dbapi_conn.cursor()
        cur.execute("PRAGMA journal_mode=WAL")
        cur.execute("PRAGMA busy_timeout=30000")
        cur.close()

    db.create_all()
    _migrate_sqlite_schema()

    # ── One-time admin bootstrap (idempotent — safe on every restart) ──
    _admin_email = os.environ.get("ADMIN_EMAIL")
    _admin_password = os.environ.get("ADMIN_PASSWORD")
    if _admin_email and _admin_password and not User.query.filter_by(email=_admin_email).first():
        _admin = User(
            first_name=os.environ.get("ADMIN_FIRST_NAME", "Admin"),
            last_name=os.environ.get("ADMIN_LAST_NAME", "HaithemEduAI"),
            email=_admin_email,
            role="admin",
        )
        _admin.set_password(_admin_password)
        db.session.add(_admin)
        db.session.flush()
        db.session.add(Subscription(
            user_id=_admin.id, plan="ultimate", status="active",
            expires_at=datetime.utcnow() + timedelta(days=3650),
            docs_used=0,
        ))
        db.session.commit()
        print(f"✅ Admin account bootstrapped: {_admin_email}")


def log_payment(user_id, plan, amount=None, note=""):
    """Record a manual payment (bank transfer / CCP) for revenue tracking."""
    db.session.add(Payment(
        user_id=user_id,
        plan=plan,
        amount=amount if amount is not None else PLAN_PRICES.get(plan, 0),
        note=note,
    ))


def log_admin_action(action, target="", details=""):
    """Lightweight audit trail — who did what, when."""
    db.session.add(AdminLog(
        admin_id=current_user.id if current_user.is_authenticated else None,
        action=action,
        target=target,
        details=details,
    ))


# ── FUNNEL TRACKING (campaign measurement, no personal data) ─────────────────
_BOT_RE = re.compile(r"bot|crawl|spider|slurp|facebookexternalhit|preview|monitor|curl|python-requests", re.I)
_SOURCE_HOSTS = (("facebook", "facebook"), ("fb.", "facebook"), ("instagram", "instagram"),
                 ("google", "google"), ("whatsapp", "whatsapp"), ("wa.me", "whatsapp"),
                 ("linkedin", "linkedin"), ("lnkd.in", "linkedin"), ("t.co", "twitter"),
                 ("twitter", "twitter"), ("x.com", "twitter"), ("youtube", "youtube"),
                 ("tiktok", "tiktok"), ("telegram", "telegram"), ("t.me", "telegram"))


def traffic_source():
    """Where this visitor came from, remembered for the whole browser session."""
    if session.get("src"):
        return session["src"]
    utm = re.sub(r"[^a-z0-9_.-]", "", request.args.get("utm_source", "").lower())[:40]
    if utm:
        src = utm
    elif request.args.get("fbclid"):
        src = "facebook"
    elif request.args.get("gclid"):
        src = "google"
    else:
        ref = (request.referrer or "").lower()
        host = ref.split("/")[2] if ref.count("/") >= 2 else ""
        src = "direct"
        if host and request.host.lower() not in host:
            src = next((name for key, name in _SOURCE_HOSTS if key in host), "other")
    session["src"] = src
    return src


def track(event, profile=None, user_id=None):
    """Record one funnel step. Added to the current DB session: the caller commits."""
    try:
        if _BOT_RE.search(request.headers.get("User-Agent", "")):
            return
        if "vid" not in session:
            session["vid"] = secrets.token_hex(8)
        db.session.add(FunnelEvent(event=event, profile=profile, source=traffic_source(),
                                   visitor=session["vid"], user_id=user_id))
    except RuntimeError:
        pass  # no request context (scheduled job): nothing to attribute


# ── REFERRAL PROGRAM ("invite a colleague") ─────────────────────────────────
# Rewards stay small and short-lived so they push towards a paid plan instead of
# replacing it. Every bonus is valid REFERRAL_BONUS_DAYS, extendable once.
REFERRAL_BONUS_INVITEE = 1   # welcome gift for the invited person, at signup
REFERRAL_BONUS_DAYS = 30
# Inviter reward, granted when the invitee generates a first document:
# (documents per invitee, lifetime cap of documents earned this way)
REFERRAL_REWARDS = {"free": (1, 3), "paid": (3, 9)}


def referral_terms(user):
    """(docs per invitee, lifetime cap) for this inviter, from their current plan."""
    sub = user.subscription
    paid = bool(sub and sub.plan in ("pro", "ultimate") and not sub.is_expired())
    return REFERRAL_REWARDS["paid" if paid else "free"]
_REF_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # no 0/O/1/I confusion


def get_referral_code(user):
    """The user's personal invite code, created on first use."""
    if not user.referral_code:
        while True:
            code = "".join(secrets.choice(_REF_ALPHABET) for _ in range(6))
            if not User.query.filter_by(referral_code=code).first():
                break
        user.referral_code = code
        db.session.commit()
    return user.referral_code


def referral_link(user):
    return url_for("home", ref=get_referral_code(user), _external=True)


def pending_referrer():
    """The account whose invite link brought this visitor, if any."""
    code = session.get("ref")
    if not code:
        return None
    return User.query.filter_by(referral_code=code, is_active=True).first()


def reward_referrer_if_due(user):
    """Called on the user's first document: credit whoever invited them (once)."""
    if not user.referred_by_id or user.referral_rewarded:
        return
    referrer = User.query.get(user.referred_by_id)
    user.referral_rewarded = True
    if not referrer or not referrer.subscription:
        return
    per_invitee, cap = referral_terms(referrer)
    n = min(per_invitee, cap - (referrer.referral_docs_earned or 0))
    if n <= 0:
        return  # cap reached: the invitation still counts, but earns nothing more
    referrer.referral_docs_earned = (referrer.referral_docs_earned or 0) + n
    referrer.subscription.add_bonus(n, days=REFERRAL_BONUS_DAYS)
    lang = referrer.preferred_lang or "fr"
    db.session.add(Message(user_id=referrer.id, sender="admin", body=tr_lang(
        lang, "referral_reward_message", name=user.first_name, n=n,
        date=referrer.subscription.bonus_expires_at.strftime("%d/%m/%Y"))))
    db.session.add(FunnelEvent(event="referral_reward", profile=referrer.profile,
                               source=referrer.signup_source, user_id=referrer.id))


# ── TRAINING PASSES (groups imported from an Excel file) ─────────────────────
CREDENTIALS_EMAILS_PER_DAY = int(os.environ.get("CREDENTIALS_EMAILS_PER_DAY", "150"))
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
_HEADER_ALIASES = {
    "last_name":  ("nom", "last name", "lastname", "name", "اللقب", "nom de famille", "surname"),
    "first_name": ("prénom", "prenom", "first name", "firstname", "الاسم", "given name"),
    "email":      ("email", "e-mail", "mail", "adresse email", "adresse e-mail", "البريد", "البريد الإلكتروني", "courriel"),
}


def discounted_price(plan, pct):
    """Annual price after the training discount, rounded to 50 DA."""
    base = PLAN_PRICES.get(plan, 0)
    if not pct:
        return base
    return int(round(base * (100 - pct) / 100 / 50.0) * 50)


def grant_pass(user, pass_type):
    """Give `user` a training pass: fresh quota, pass duration, and the
    end-of-pass discount on the annual plans."""
    cfg = PASS_TYPES[pass_type]
    now = datetime.utcnow()
    sub = user.subscription
    if not sub:
        sub = Subscription(user_id=user.id)
        db.session.add(sub)
        db.session.flush()
    previous = sub.plan
    sub.plan = pass_type
    sub.status = "active"
    sub.started_at = now
    sub.expires_at = now + timedelta(days=cfg["days"])
    sub.docs_limit = cfg["docs"]
    sub.docs_used = 0
    sub.docs_used_today = 0
    sub.last_reset_date = date.today()
    sub.expiry_reminder_sent = False
    sub.last_reminder_stage = None
    user.discount_pct = cfg["discount"]
    user.discount_until = sub.expires_at + timedelta(days=cfg["discount_days"])
    user.pass_reminder_sent = False
    db.session.add(PlanChangeHistory(user_id=user.id, from_plan=previous,
                                      to_plan=pass_type, reason="training_pass"))


def read_participants(file_storage):
    """Rows {first_name, last_name, email} from an .xlsx or .csv file whose first
    row holds the column titles (French, English or Arabic). Returns (rows, error)."""
    name = (file_storage.filename or "").lower()
    data = file_storage.read()
    table = []
    try:
        if name.endswith(".xlsx"):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            for r in wb.active.iter_rows(values_only=True):
                table.append(["" if v is None else str(v).strip() for v in r])
        elif name.endswith(".csv"):
            import csv
            text = data.decode("utf-8-sig", errors="ignore")
            dialect = csv.Sniffer().sniff(text[:2000], delimiters=",;\t") if text.strip() else csv.excel
            table = [[c.strip() for c in r] for r in csv.reader(io.StringIO(text), dialect)]
        else:
            return [], "format"
    except Exception:
        return [], "unreadable"
    table = [r for r in table if any(r)]
    if len(table) < 2:
        return [], "empty"
    header = [h.lower().strip() for h in table[0]]
    cols = {}
    for key, aliases in _HEADER_ALIASES.items():
        for i, h in enumerate(header):
            if h in aliases and i not in cols.values():
                cols[key] = i
                break
    if "email" not in cols:
        return [], "no_email_column"
    rows = []
    for r in table[1:]:
        get = lambda k: (r[cols[k]] if k in cols and cols[k] < len(r) else "").strip()
        rows.append({"first_name": get("first_name"), "last_name": get("last_name"),
                     "email": get("email").lower()})
    return rows, None


def import_participants(rows, pass_type, cohort, send_email):
    """Create the accounts (or give the pass to existing ones). Returns a report list."""
    report = []
    seen = set()
    for row in rows:
        email = row["email"]
        if not _EMAIL_RE.match(email) or email in seen:
            report.append({**row, "status": "invalid" if email not in seen else "duplicate"})
            continue
        seen.add(email)
        user = User.query.filter_by(email=email).first()
        if user:
            sub = user.subscription
            if user.role == "admin" or (sub and sub.plan in PLAN_PRICES and not sub.is_expired()):
                report.append({**row, "status": "already_paid"})  # never downgrade a paying user
                continue
            grant_pass(user, pass_type)
            user.cohort_id = cohort.id
            report.append({**row, "status": "existing"})
            continue
        password = generate_password(10)
        user = User(first_name=row["first_name"] or email.split("@")[0], last_name=row["last_name"] or "-",
                    email=email, signup_source="formation", must_change_password=True,
                    temp_password=password, credentials_pending=bool(send_email), cohort_id=cohort.id)
        user.set_password(password)
        db.session.add(user)
        db.session.flush()
        get_referral_code(user)
        grant_pass(user, pass_type)
        report.append({**row, "status": "created", "password": password})
    return report


def send_credentials_batch():
    """Send queued login emails, at most CREDENTIALS_EMAILS_PER_DAY per day
    (Brevo's free plan allows 300 a day, shared with the other emails)."""
    from notifications import send_email, smtp_configured
    from email_templates import credentials_email
    if not smtp_configured():
        return 0
    with app.app_context():
        day_start = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
        sent_today = User.query.filter(User.credentials_sent_at >= day_start).count()
        budget = max(0, CREDENTIALS_EMAILS_PER_DAY - sent_today)
        if not budget:
            return 0
        queue = User.query.filter(User.credentials_pending.is_(True)).order_by(User.id).limit(budget).all()
        sent = 0
        for u in queue:
            sub = u.subscription
            cfg = PASS_TYPES.get(sub.plan if sub else "", {})
            if not u.temp_password or not u.must_change_password:
                u.credentials_pending = False  # already logged in and changed it
                continue
            subject, html, text = credentials_email(
                u.preferred_lang or "fr", u.first_name, u.email, u.temp_password,
                f"{APP_URL}/login", (sub.plan_label() if sub else ""), cfg.get("days", 0),
                cfg.get("docs", 0), cfg.get("monthly", False))
            if send_email(u.email, subject, text, html):
                u.credentials_pending = False
                u.credentials_sent_at = datetime.utcnow()
                sent += 1
        db.session.commit()
    if sent:
        print(f"✅ Login emails sent: {sent}")
    return sent


def check_passes():
    """J-7 reminder with the discount, and the switch to 'expired' at the end."""
    from notifications import send_email
    from email_templates import pass_reminder_email
    with app.app_context():
        now = datetime.utcnow()
        subs = Subscription.query.filter(Subscription.plan.in_(list(PASS_TYPES))).all()
        for sub in subs:
            user = db.session.get(User, sub.user_id)
            if not user or not sub.expires_at:
                continue
            lang = user.preferred_lang or "fr"
            if sub.expires_at <= now:
                # Pass over: no quota left until they subscribe (discount still valid).
                db.session.add(PlanChangeHistory(user_id=user.id, from_plan=sub.plan,
                                                  to_plan="trial", reason="pass_ended"))
                sub.plan = "trial"
                sub.docs_limit = sub.docs_used or 0
                db.session.add(Message(user_id=user.id, sender="admin", read_by_admin=True,
                                       read_by_teacher=False,
                                       body=tr_lang(lang, "msg_pass_ended", pct=user.discount_pct or 0,
                                                    date=user.discount_until.strftime("%d/%m/%Y") if user.discount_until else "")))
            elif not user.pass_reminder_sent and sub.expires_at - now <= timedelta(days=7):
                subject, html, text = pass_reminder_email(
                    lang, user.first_name, sub.expires_at.strftime("%d/%m/%Y"),
                    user.discount_pct or 0,
                    user.discount_until.strftime("%d/%m/%Y") if user.discount_until else "",
                    f"{APP_URL}/upgrade", optout_link(user))
                if not user.email_optout:
                    send_email(user.email, subject, text, html)
                db.session.add(Message(user_id=user.id, sender="admin", read_by_admin=True,
                                       read_by_teacher=False,
                                       body=tr_lang(lang, "msg_pass_reminder", date=sub.expires_at.strftime("%d/%m/%Y"),
                                                    pct=user.discount_pct or 0)))
                user.pass_reminder_sent = True
        db.session.commit()


def _pass_jobs():
    try:
        send_credentials_batch()
        check_passes()
        send_session_reminders()
    except Exception as e:  # never crash the scheduler
        print(f"⚠️  Training pass jobs failed: {e}")


def apply_payment_approval(teacher, plan, days=365, reason="payment_approved"):
    """Activate or renew a teacher's subscription after a payment has been validated.
    Renewing the same plan extends from the later of (now, current expiry);
    switching plan (or reactivating an expired/trial account) starts a fresh period."""
    sub = teacher.subscription
    if not sub:
        sub = Subscription(user_id=teacher.id)
        db.session.add(sub)
        db.session.flush()

    is_renewal = (sub.plan == plan and sub.status == "active" and not sub.is_expired())
    base = sub.expires_at if (is_renewal and sub.expires_at) else datetime.utcnow()
    previous_plan = sub.plan

    sub.plan = plan
    sub.status = "active"
    sub.expires_at = base + timedelta(days=days)
    sub.docs_used = 0
    sub.docs_used_today = 0
    sub.last_reset_date = None
    sub.expiry_reminder_sent = False
    sub.last_reminder_stage = None

    if previous_plan != plan or not is_renewal:
        db.session.add(PlanChangeHistory(user_id=teacher.id, from_plan=previous_plan,
                                          to_plan=plan, reason=reason))
    # Attributed to the customer's own acquisition source, not the admin's browser.
    db.session.add(FunnelEvent(event="payment_approved", profile=teacher.profile,
                               source=teacher.signup_source, user_id=teacher.id))
    return is_renewal


# ── DATABASE BACKUPS ─────────────────────────────────────────────────────────
# A consistent copy of the SQLite database is taken every night (and at startup
# when the last one is older than a day) into DATA_DIR/backups, keeping the
# most recent BACKUP_KEEP copies. The admin can download them (plus receipts)
# to keep an off-site copy.
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
BACKUP_KEEP = int(os.environ.get("BACKUP_KEEP", "14"))
_BACKUP_NAME_RE = re.compile(r"^eduprompt-\d{8}-\d{6}\.db$")


def backup_database():
    """Write a consistent snapshot of the live database; returns its file name."""
    import sqlite3
    src_path = os.path.join(DATA_DIR, "eduprompt.db")
    if not os.path.exists(src_path):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    name = datetime.utcnow().strftime("eduprompt-%Y%m%d-%H%M%S.db")
    src = sqlite3.connect(src_path)
    dst = sqlite3.connect(os.path.join(BACKUP_DIR, name))
    try:
        src.backup(dst)  # safe while the app is writing
    finally:
        dst.close()
        src.close()
    for old in list_backups()[BACKUP_KEEP:]:
        try:
            os.remove(os.path.join(BACKUP_DIR, old["name"]))
        except OSError:
            pass
    return name


def list_backups():
    """Backups, most recent first: [{name, size, created}]."""
    if not os.path.isdir(BACKUP_DIR):
        return []
    out = []
    for f in os.listdir(BACKUP_DIR):
        if _BACKUP_NAME_RE.match(f):
            p = os.path.join(BACKUP_DIR, f)
            out.append({"name": f, "size": os.path.getsize(p),
                        "created": datetime.strptime(f[10:25], "%Y%m%d-%H%M%S")})
    return sorted(out, key=lambda b: b["name"], reverse=True)


def _backup_job():
    try:
        name = backup_database()
        if name:
            print(f"✅ Database backup written: {name}")
    except Exception as e:  # never let a failed backup crash the scheduler
        print(f"⚠️  Database backup failed: {e}")


# ── ACTIVATION FOLLOW-UP EMAILS ──────────────────────────────────────────────
def _optout_serializer():
    return URLSafeSerializer(app.config["SECRET_KEY"], salt="email-optout")


def optout_link(user):
    return f"{APP_URL}/email/stop/{_optout_serializer().dumps(user.id)}"


def _nudge_tracked(user, stage):
    db.session.add(FunnelEvent(event=f"nudge_{stage}", profile=user.profile,
                               source=user.signup_source, user_id=user.id))


def _nudge_job():
    try:
        send_activation_nudges(app, db, User, Document, APP_URL, optout_link, _nudge_tracked)
    except Exception as e:  # never let it crash the scheduler
        print(f"⚠️  Activation emails failed: {e}")


# ── DAILY REMINDER SCHEDULER (5-day heads-up before subscription expiry) ─────
try:
    from apscheduler.schedulers.background import BackgroundScheduler
    _scheduler = BackgroundScheduler(daemon=True)
    _scheduler.add_job(
        lambda: check_and_send_expiry_reminders(app, db, User, Subscription, CONTACT_EMAIL),
        "interval", hours=24, next_run_time=datetime.utcnow(),
        id="expiry_reminders", replace_existing=True,
    )
    _scheduler.add_job(_backup_job, "cron", hour=2, minute=17,  # 03:17 in Algeria, low traffic
                       id="db_backup", replace_existing=True)
    _last = list_backups()
    if not _last or datetime.utcnow() - _last[0]["created"] > timedelta(days=1):
        # No recent backup (first deploy, or the app was down at 03:17): take one now.
        _scheduler.add_job(_backup_job, "date", run_date=datetime.now() + timedelta(seconds=60),
                           id="db_backup_catchup", replace_existing=True)
    _scheduler.add_job(_nudge_job, "interval", hours=1, id="activation_nudges",
                       replace_existing=True)
    _scheduler.add_job(_pass_jobs, "interval", minutes=30, id="training_passes",
                       replace_existing=True)
    _scheduler.start()
    print("✅ Expiry reminder scheduler started")
except Exception as e:
    print(f"⚠️  Could not start reminder scheduler: {e}")


# ── STATIC REFERENCE DATA ────────────────────────────────────────────────────
DOCUMENT_TYPES = {
    "ar": [
        {"id": "lesson_plan",   "label": "مذكرة درس تفصيلية"},
        {"id": "activity",      "label": "ورقة نشاط / تطبيق"},
        {"id": "quiz",          "label": "اختبار QCM مع التصحيح"},
        {"id": "rubric",        "label": "شبكة تقييم Rubric"},
        {"id": "differentiated","label": "دعم متمايز (3 مستويات)"},
        {"id": "summary",       "label": "ملخص / بطاقة مرجعية"},
        {"id": "course_plan",   "label": "تسلسل وحدة تعليمية"},
        {"id": "homework",      "label": "واجب منزلي مُهيكل"},
    ],
    "fr": [
        {"id": "lesson_plan",   "label": "Fiche de cours détaillée"},
        {"id": "activity",      "label": "Feuille d'activité / exercices"},
        {"id": "quiz",          "label": "QCM avec corrigé"},
        {"id": "rubric",        "label": "Grille d'évaluation Rubric"},
        {"id": "differentiated","label": "Supports différenciés (3 niveaux)"},
        {"id": "summary",       "label": "Résumé / fiche mémo"},
        {"id": "course_plan",   "label": "Progression de séquence"},
        {"id": "homework",      "label": "Devoir maison structuré"},
    ],
    "en": [
        {"id": "lesson_plan",   "label": "Detailed lesson plan"},
        {"id": "activity",      "label": "Activity / worksheet"},
        {"id": "quiz",          "label": "MCQ quiz with answer key"},
        {"id": "rubric",        "label": "Assessment rubric"},
        {"id": "differentiated","label": "Differentiated supports (3 levels)"},
        {"id": "summary",       "label": "Summary / reference card"},
        {"id": "course_plan",   "label": "Unit sequence plan"},
        {"id": "homework",      "label": "Structured homework"},
    ],
}

# Tools offered to parents (profile == "parent") instead of the teacher documents.
PARENT_DOCUMENT_TYPES = {
    "ar": [
        {"id": "revision_sheet",      "label": "بطاقة مراجعة"},
        {"id": "exercises_corrected", "label": "تمارين مع التصحيح"},
        {"id": "simple_explanation",  "label": "شرح مبسّط للدرس"},
        {"id": "exam_prep",           "label": "تحضير للامتحان"},
    ],
    "fr": [
        {"id": "revision_sheet",      "label": "Fiche de révision"},
        {"id": "exercises_corrected", "label": "Exercices + corrigé"},
        {"id": "simple_explanation",  "label": "Explication simplifiée"},
        {"id": "exam_prep",           "label": "Préparation aux examens"},
    ],
    "en": [
        {"id": "revision_sheet",      "label": "Revision sheet"},
        {"id": "exercises_corrected", "label": "Exercises + answer key"},
        {"id": "simple_explanation",  "label": "Simple explanation"},
        {"id": "exam_prep",           "label": "Exam preparation"},
    ],
}

LEVELS = {
    "ar": ["ابتدائي","متوسط","ثانوي","تكوين مهني","جامعي"],
    "fr": ["Primaire","Moyen","Lycée","Formation professionnelle","Université"],
    "en": ["Primary","Middle","High school","Vocational training","University"],
}

SUBJECTS = {
    "ar": ["رياضيات","علوم","تاريخ وجغرافيا","لغة عربية","فيزياء","كيمياء","إعلام آلي","لغة فرنسية","لغة إنجليزية","فلسفة","اقتصاد","أخرى"],
    "fr": ["Mathématiques","Sciences","Histoire-Géographie","Arabe","Physique","Chimie","Informatique","Français","Anglais","Philosophie","Économie","Autre"],
    "en": ["Mathematics","Science","History-Geography","Arabic","Physics","Chemistry","Computer Science","French","English","Philosophy","Economics","Other"],
}


# ── AUTH ROUTES ───────────────────────────────────────────────────────────────
@app.route("/")
def home():
    if not session.get("lv"):
        session["lv"] = 1
        track("landing_view")
        db.session.commit()
    return render_template("landing.html", levels=LEVELS, subjects=SUBJECTS)


@app.route("/about")
def about():
    from cv_data import CV
    return render_template("about.html", cv=CV.get(current_lang(), CV["fr"]))


@app.route("/register", methods=["GET", "POST"])
@limiter.limit("15/hour")
def register():
    if current_user.is_authenticated:
        return redirect(url_for("generator"))

    if request.method == "POST":
        first_name = request.form.get("first_name", "").strip()
        last_name  = request.form.get("last_name", "").strip()
        email      = request.form.get("email", "").strip().lower()
        password   = request.form.get("password", "")
        api_key    = request.form.get("api_key", "").strip()
        wanted_plan = request.form.get("wanted_plan", "trial").strip()
        profile    = "parent" if request.form.get("profile") == "parent" else "teacher"

        if not all([first_name, last_name, email, password]):
            flash(tr("flash_all_fields_required"), "error")
            return render_template("register.html")

        if len(password) < 6:
            flash(tr("flash_password_min_length"), "error")
            return render_template("register.html")

        if User.query.filter_by(email=email).first():
            flash(tr("flash_email_exists"), "error")
            return render_template("register.html")

        # A paid plan chosen at signup requires proof of payment, validated afterwards.
        method = request.form.get("method", "").strip()
        reference = request.form.get("reference", "").strip()
        receipt_file = request.files.get("receipt")
        if wanted_plan in ("pro", "ultimate"):
            if method not in PAYMENT_METHODS or not receipt_file or not receipt_file.filename:
                flash(tr("flash_payment_proof_required"), "error")
                return render_template("register.html")

        user = User(first_name=first_name, last_name=last_name, email=email,
                    gemini_api_key=api_key or None, preferred_lang=current_lang(),
                    profile=profile, signup_source=traffic_source())
        user.set_password(password)
        db.session.add(user)
        db.session.flush()  # get user.id before commit
        track("signup", profile=profile, user_id=user.id)
        create_trial_subscription(user)  # immediate access while any payment is reviewed
        referrer = pending_referrer()
        if referrer and referrer.email != email:
            user.referred_by_id = referrer.id
            user.subscription.add_bonus(REFERRAL_BONUS_INVITEE, days=REFERRAL_BONUS_DAYS)
            session.pop("ref", None)

        if wanted_plan in ("pro", "ultimate"):
            receipt_name = save_receipt_file(receipt_file, user.id)
            db.session.add(PaymentRequest(
                user_id=user.id, plan=wanted_plan, method=method,
                amount_claimed=PLAN_PRICES.get(wanted_plan),
                reference=reference or None, receipt_path=receipt_name,
                source="registration",
            ))

        db.session.commit()

        login_user(user)
        if wanted_plan in ("pro", "ultimate"):
            flash(tr("flash_welcome_trial_pending_payment"), "success")
        elif user.referred_by_id:
            flash(tr("flash_welcome_referral", n=REFERRAL_BONUS_INVITEE), "success")
        else:
            flash(tr("flash_welcome_trial"), "success")
        return redirect(url_for("generator"))

    return render_template("register.html", referrer=pending_referrer(),
                           ref_bonus_invitee=REFERRAL_BONUS_INVITEE)


@app.route("/login", methods=["GET", "POST"])
@limiter.limit("20/hour")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("generator"))

    if request.method == "POST":
        email    = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        user = User.query.filter_by(email=email).first()

        if user and user.check_password(password):
            if not user.is_active:
                flash(tr("flash_account_deactivated"), "error")
                return render_template("login.html")
            login_user(user, remember=True)
            return redirect(url_for("generator"))
        flash(tr("flash_wrong_credentials"), "error")

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# ── DASHBOARD ─────────────────────────────────────────────────────────────────
@app.route("/dashboard")
@login_required
def dashboard():
    sub = current_user.subscription
    recent_docs = current_user.documents[:10]
    recent_messages = Message.query.filter_by(user_id=current_user.id) \
        .order_by(Message.created_at.desc()).limit(3).all()
    unread_admin_msgs = Message.query.filter_by(user_id=current_user.id, sender="admin",
                                                 read_by_teacher=False).count()
    invited = User.query.filter_by(referred_by_id=current_user.id).count()
    invited_active = User.query.filter_by(referred_by_id=current_user.id, referral_rewarded=True).count()
    per_invitee, cap = referral_terms(current_user)
    return render_template("dashboard.html", sub=sub, docs=recent_docs,
                            recent_messages=recent_messages, unread_admin_msgs=unread_admin_msgs,
                            ref_link=referral_link(current_user), ref_invited=invited,
                            ref_active=invited_active, ref_bonus=per_invitee, ref_cap=cap,
                            ref_cap_people=cap // per_invitee,
                            ref_cap_reached=(current_user.referral_docs_earned or 0) >= cap,
                            ref_bonus_invitee=REFERRAL_BONUS_INVITEE, ref_days=REFERRAL_BONUS_DAYS,
                            ref_paid_terms=REFERRAL_REWARDS["paid"])


# ── PROFILE (self-service) ───────────────────────────────────────────────────
@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    if request.method == "POST":
        action = request.form.get("action")

        if action == "update_info":
            first_name = request.form.get("first_name", "").strip()
            last_name  = request.form.get("last_name", "").strip()
            email      = request.form.get("email", "").strip().lower()

            if not all([first_name, last_name, email]):
                flash(tr("flash_all_fields_required"), "error")
            else:
                existing = User.query.filter_by(email=email).first()
                if existing and existing.id != current_user.id:
                    flash(tr("flash_email_exists"), "error")
                else:
                    current_user.first_name = first_name
                    current_user.last_name = last_name
                    current_user.email = email
                    db.session.commit()
                    flash(tr("flash_profile_updated"), "success")

        elif action == "change_password":
            current_pw = request.form.get("current_password", "")
            new_pw = request.form.get("new_password", "")
            if not current_user.check_password(current_pw):
                flash(tr("flash_wrong_current_password"), "error")
            elif len(new_pw) < 6:
                flash(tr("flash_password_min_length"), "error")
            else:
                current_user.set_password(new_pw)
                current_user.must_change_password = False
                current_user.temp_password = None   # no longer kept anywhere
                current_user.credentials_pending = False
                db.session.commit()
                flash(tr("flash_password_changed"), "success")
                return redirect(url_for("dashboard"))

        elif action == "update_apikey":
            api_key = request.form.get("api_key", "").strip()
            # Empty field = go back to the AI service included with the platform.
            current_user.gemini_api_key = api_key or None
            db.session.commit()
            flash(tr("flash_apikey_updated" if api_key else "flash_apikey_removed"), "success")

        return redirect(url_for("profile"))

    return render_template("profile.html")


@app.route("/upgrade", methods=["GET", "POST"])
@login_required
def upgrade():
    if request.method == "POST":
        action = request.form.get("action", "code")

        if action == "declare_payment":
            plan = request.form.get("plan", "pro")
            amount = request.form.get("amount", "").strip()
            method = request.form.get("method", "").strip()
            reference = request.form.get("reference", "").strip()
            note = request.form.get("note", "").strip()
            receipt_file = request.files.get("receipt")

            if plan not in ("pro", "ultimate") or method not in PAYMENT_METHODS:
                flash(tr("flash_payment_fields_required"), "error")
                return redirect(url_for("upgrade"))
            if method == "ccp" and not reference:
                flash(tr("flash_payment_fields_required"), "error")
                return redirect(url_for("upgrade"))
            if not receipt_file or not receipt_file.filename:
                flash(tr("flash_payment_proof_required"), "error")
                return redirect(url_for("upgrade"))

            receipt_name = save_receipt_file(receipt_file, current_user.id)
            pct = current_user.active_discount()
            price = discounted_price(plan, pct)
            if pct:
                note = (f"[Réduction formation -{pct} % : {price} DA au lieu de {PLAN_PRICES[plan]} DA] " + note).strip()
            db.session.add(PaymentRequest(
                user_id=current_user.id, plan=plan, method=method,
                amount_claimed=int(amount) if amount.isdigit() else price,
                reference=reference or None, note=note, receipt_path=receipt_name,
                source="upgrade",
            ))
            track("payment_declared", profile=current_user.profile, user_id=current_user.id)
            db.session.commit()
            flash(tr("flash_payment_declared"), "success")
            return redirect(url_for("upgrade"))

        code_str = request.form.get("code", "").strip().upper()
        lic = LicenseCode.query.filter_by(code=code_str, used=False).first()

        if not lic:
            flash(tr("flash_invalid_code"), "error")
            return redirect(url_for("upgrade"))

        apply_payment_approval(current_user, lic.plan, days=lic.duration_days, reason="license_code")

        lic.used = True
        lic.used_by = current_user.id
        lic.used_at = datetime.utcnow()

        log_payment(current_user.id, lic.plan, note=f"License code {lic.code}")

        db.session.commit()
        flash(tr("flash_upgraded", plan=current_user.subscription.plan_label()), "success")
        return redirect(url_for("dashboard"))

    my_requests = PaymentRequest.query.filter_by(user_id=current_user.id) \
        .order_by(PaymentRequest.created_at.desc()).all()
    pct = current_user.active_discount()
    return render_template("upgrade.html", my_requests=my_requests, discount_pct=pct,
                           discount_until=current_user.discount_until,
                           prices={p: PLAN_PRICES[p] for p in PLAN_PRICES},
                           discounted={p: discounted_price(p, pct) for p in PLAN_PRICES})


# ── ADMIN: generate license codes ────────────────────────────────────────────
@app.route("/admin/licenses", methods=["GET", "POST"])
@login_required
def admin_licenses():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        plan = request.form.get("plan", "pro")
        days = int(request.form.get("days", 365))
        lic = LicenseCode(plan=plan, duration_days=days)
        db.session.add(lic)
        log_admin_action("license_generated", target=f"plan={plan}", details=f"{days} days")
        db.session.commit()
        flash(tr("flash_license_generated", code=lic.code), "success")

    q = request.args.get("q", "").strip()
    status_filter = request.args.get("status", "").strip()
    page = request.args.get("page", 1, type=int)

    query = LicenseCode.query
    if q:
        query = query.filter(LicenseCode.code.ilike(f"%{q}%"))
    if status_filter == "used":
        query = query.filter(LicenseCode.used.is_(True))
    elif status_filter == "available":
        query = query.filter(LicenseCode.used.is_(False))

    query = query.order_by(LicenseCode.created_at.desc())
    pager = query.paginate(page=page, per_page=20, error_out=False)

    return render_template("admin_licenses.html", codes=pager.items, pager=pager,
                            q=q, status_filter=status_filter)


# ── ADMIN: payment validation — activate / renew pro & premium subscriptions ─
@app.route("/admin/payments")
@login_required
def admin_payments():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    status_filter = request.args.get("status", "pending").strip()
    q = request.args.get("q", "").strip()
    page = request.args.get("page", 1, type=int)

    query = PaymentRequest.query.join(User, PaymentRequest.user_id == User.id)
    if status_filter in ("pending", "approved", "rejected"):
        query = query.filter(PaymentRequest.status == status_filter)
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(User.first_name.ilike(like), User.last_name.ilike(like),
                                     User.email.ilike(like), PaymentRequest.reference.ilike(like)))

    query = query.order_by(PaymentRequest.created_at.desc())
    pager = query.paginate(page=page, per_page=20, error_out=False)

    return render_template("admin_payments.html", pager=pager, q=q, status_filter=status_filter)


@app.route("/admin/payments/<int:req_id>/approve", methods=["POST"])
@login_required
def admin_approve_payment(req_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    preq = PaymentRequest.query.get_or_404(req_id)
    if preq.status != "pending":
        flash(tr("flash_payment_already_reviewed"), "error")
        return redirect(url_for("admin_payments"))

    days = request.form.get("days", 365, type=int)
    is_renewal = apply_payment_approval(preq.user, preq.plan, days=days)

    log_payment(preq.user_id, preq.plan, amount=preq.amount_claimed,
                note=f"PaymentRequest #{preq.id} ({preq.reference or preq.method})")

    preq.status = "approved"
    preq.reviewed_at = datetime.utcnow()
    preq.reviewed_by = current_user.id
    preq.receipt_number = next_receipt_number()

    kind = "renewal" if is_renewal else "activation"
    log_admin_action("payment_approved", target=preq.user.email,
                      details=f"plan={preq.plan} {kind} +{days}d")

    notif_body = tr("msg_payment_approved", plan=preq.user.subscription.plan_label(),
                     date=preq.user.subscription.expires_at.strftime("%d/%m/%Y"))
    db.session.add(Message(user_id=preq.user_id, sender="admin", body=notif_body,
                            read_by_admin=True, read_by_teacher=False))

    db.session.commit()
    flash(tr("flash_payment_approved", name=preq.user.full_name), "success")
    return redirect(url_for("admin_payments"))


@app.route("/admin/payments/<int:req_id>/reject", methods=["POST"])
@login_required
def admin_reject_payment(req_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    preq = PaymentRequest.query.get_or_404(req_id)
    if preq.status != "pending":
        flash(tr("flash_payment_already_reviewed"), "error")
        return redirect(url_for("admin_payments"))

    reason = request.form.get("reason", "").strip()
    preq.status = "rejected"
    preq.reviewed_at = datetime.utcnow()
    preq.reviewed_by = current_user.id
    preq.admin_note = reason

    log_admin_action("payment_rejected", target=preq.user.email, details=reason)

    notif_body = tr("msg_payment_rejected", reason=reason or tr("msg_payment_rejected_generic"))
    db.session.add(Message(user_id=preq.user_id, sender="admin", body=notif_body,
                            read_by_admin=True, read_by_teacher=False))

    db.session.commit()
    flash(tr("flash_payment_rejected", name=preq.user.full_name), "success")
    return redirect(url_for("admin_payments"))


@app.route("/admin/payments/<int:req_id>/receipt")
@login_required
def admin_view_receipt(req_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))
    preq = PaymentRequest.query.get_or_404(req_id)
    if not preq.receipt_path:
        flash(tr("flash_no_receipt"), "error")
        return redirect(url_for("admin_payments"))
    path = os.path.join(RECEIPTS_DIR, preq.receipt_path)
    if not os.path.exists(path):
        flash(tr("flash_no_receipt"), "error")
        return redirect(url_for("admin_payments"))
    return send_file(path)


# ── ADMIN: cash register — approved payments by date / method ───────────────
@app.route("/admin/cashbox")
@login_required
def admin_cashbox():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    from sqlalchemy import func

    date_from = request.args.get("from", "").strip()
    date_to = request.args.get("to", "").strip()
    method_filter = request.args.get("method", "").strip()
    page = request.args.get("page", 1, type=int)

    query = PaymentRequest.query.filter_by(status="approved")
    if method_filter in PAYMENT_METHODS:
        query = query.filter(PaymentRequest.method == method_filter)
    if date_from:
        try:
            query = query.filter(PaymentRequest.reviewed_at >= datetime.strptime(date_from, "%Y-%m-%d"))
        except ValueError:
            pass
    if date_to:
        try:
            end = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
            query = query.filter(PaymentRequest.reviewed_at < end)
        except ValueError:
            pass

    query = query.order_by(PaymentRequest.reviewed_at.desc())
    pager = query.paginate(page=page, per_page=25, error_out=False)

    totals_query = query
    total_amount = db.session.query(func.sum(PaymentRequest.amount_claimed)) \
        .filter(PaymentRequest.id.in_([p.id for p in totals_query.all()])).scalar() or 0

    now = datetime.utcnow()
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    year_start = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)

    def sum_since(since):
        return db.session.query(func.sum(PaymentRequest.amount_claimed)).filter(
            PaymentRequest.status == "approved", PaymentRequest.reviewed_at >= since
        ).scalar() or 0

    def sum_by_method(method):
        return db.session.query(func.sum(PaymentRequest.amount_claimed)).filter(
            PaymentRequest.status == "approved", PaymentRequest.method == method
        ).scalar() or 0

    return render_template("admin_cashbox.html", pager=pager, date_from=date_from,
                            date_to=date_to, method_filter=method_filter,
                            filtered_total=total_amount,
                            total_today=sum_since(today_start),
                            total_month=sum_since(month_start),
                            total_year=sum_since(year_start),
                            total_cash=sum_by_method("cash"),
                            total_ccp=sum_by_method("ccp"))


@app.route("/admin/cashbox/export.csv")
@login_required
def admin_cashbox_export():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    import csv, io as _io
    date_from = request.args.get("from", "").strip()
    date_to = request.args.get("to", "").strip()
    method_filter = request.args.get("method", "").strip()

    query = PaymentRequest.query.filter_by(status="approved")
    if method_filter in PAYMENT_METHODS:
        query = query.filter(PaymentRequest.method == method_filter)
    if date_from:
        try:
            query = query.filter(PaymentRequest.reviewed_at >= datetime.strptime(date_from, "%Y-%m-%d"))
        except ValueError:
            pass
    if date_to:
        try:
            end = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
            query = query.filter(PaymentRequest.reviewed_at < end)
        except ValueError:
            pass
    rows = query.order_by(PaymentRequest.reviewed_at.desc()).all()

    buf = _io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["#", "Reçu N°", "Date", "Enseignant", "Email", "Offre", "Montant (DA)",
                      "Mode de paiement", "Référence"])
    for i, p in enumerate(rows, start=1):
        writer.writerow([
            i, p.receipt_number or "", p.reviewed_at.strftime("%d/%m/%Y %H:%M") if p.reviewed_at else "",
            p.user.full_name if p.user else "", p.user.email if p.user else "",
            p.plan, p.amount_claimed or 0, p.method_label(), p.reference or "",
        ])
    log_admin_action("export_csv", target="cashbox")
    db.session.commit()
    return app.response_class(buf.getvalue(), mimetype="text/csv",
                               headers={"Content-Disposition": "attachment; filename=caisse.csv"})


# ── ADMIN: advanced statistics dashboard ─────────────────────────────────────
@app.route("/admin/stats")
@login_required
def admin_stats():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    from sqlalchemy import func

    total_teachers = User.query.filter_by(role="user").count()
    plan_counts = {}
    for p in ("trial", "pro", "ultimate"):
        plan_counts[p] = Subscription.query.filter_by(plan=p).count()

    total_docs = Document.query.count()

    # Documents per day, last 14 days
    since = datetime.utcnow() - timedelta(days=14)
    daily_rows = (
        db.session.query(func.date(Document.created_at), func.count(Document.id))
        .filter(Document.created_at >= since)
        .group_by(func.date(Document.created_at))
        .all()
    )
    daily_map = {str(d): c for d, c in daily_rows}
    chart_labels, chart_values = [], []
    for i in range(13, -1, -1):
        day = (datetime.utcnow() - timedelta(days=i)).date()
        chart_labels.append(day.strftime("%d/%m"))
        chart_values.append(daily_map.get(str(day), 0))

    # Revenue
    total_revenue = db.session.query(func.sum(Payment.amount)).scalar() or 0
    month_start = datetime.utcnow().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    month_revenue = (
        db.session.query(func.sum(Payment.amount))
        .filter(Payment.created_at >= month_start)
        .scalar() or 0
    )
    recent_payments = Payment.query.order_by(Payment.created_at.desc()).limit(20).all()

    # Expiring within 5 days
    now = datetime.utcnow()
    soon = now + timedelta(days=5)
    expiring = (
        Subscription.query.filter(
            Subscription.status == "active",
            Subscription.expires_at.isnot(None),
            Subscription.expires_at >= now,
            Subscription.expires_at <= soon,
        ).all()
    )

    # Top 5 most active teachers by number of documents generated
    top_rows = (
        db.session.query(User, func.count(Document.id).label("cnt"))
        .join(Document, Document.user_id == User.id)
        .group_by(User.id)
        .order_by(func.count(Document.id).desc())
        .limit(5)
        .all()
    )

    pending_payments = PaymentRequest.query.filter_by(status="pending").count()
    errors_7d = GenerationError.query.filter(
        GenerationError.created_at >= datetime.utcnow() - timedelta(days=7)).count()

    # Monthly leaderboard — top active teachers for a selectable month (default: current)
    month_param = request.args.get("month", "").strip()
    try:
        month_ref = datetime.strptime(month_param, "%Y-%m") if month_param else datetime.utcnow()
    except ValueError:
        month_ref = datetime.utcnow()
    lb_start = month_ref.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    lb_end = (lb_start + timedelta(days=32)).replace(day=1)
    leaderboard = (
        db.session.query(User, func.count(Document.id).label("cnt"))
        .join(Document, Document.user_id == User.id)
        .filter(Document.created_at >= lb_start, Document.created_at < lb_end)
        .group_by(User.id)
        .order_by(func.count(Document.id).desc())
        .limit(10)
        .all()
    )

    return render_template("admin_stats.html",
                           total_teachers=total_teachers, plan_counts=plan_counts,
                           total_docs=total_docs, chart_labels=chart_labels, chart_values=chart_values,
                           total_revenue=total_revenue, month_revenue=month_revenue,
                           recent_payments=recent_payments, expiring=expiring,
                           top_teachers=top_rows, pending_payments=pending_payments,
                           leaderboard=leaderboard, leaderboard_month=lb_start.strftime("%Y-%m"),
                           errors_7d=errors_7d)


# ── ADMIN: acquisition funnel (campaign measurement) ────────────────────────
FUNNEL_STEPS = ["landing_view", "demo", "signup", "first_document", "payment_declared", "payment_approved"]


@app.route("/admin/funnel")
@login_required
def admin_funnel():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    from sqlalchemy import func

    days = request.args.get("days", 30, type=int)
    if days not in (1, 7, 30, 90, 0):
        days = 30
    profile = request.args.get("profile", "")
    if profile not in ("teacher", "parent"):
        profile = ""

    base = FunnelEvent.query
    if days:
        since = datetime.utcnow() - timedelta(days=days)
        base = base.filter(FunnelEvent.created_at >= since)

    def step_count(query, event):
        q = query.filter(FunnelEvent.event == event)
        if profile and event != "landing_view":  # visits happen before the profile is known
            q = q.filter(FunnelEvent.profile == profile)
        if event == "landing_view":
            return q.with_entities(func.count(func.distinct(FunnelEvent.visitor))).scalar() or 0
        return q.count()

    counts = {e: step_count(base, e) for e in FUNNEL_STEPS}
    visits = counts["landing_view"]
    steps = []
    prev = None
    for e in FUNNEL_STEPS:
        n = counts[e]
        steps.append({
            "event": e, "count": n,
            "of_visits": round(100 * n / visits, 1) if visits and e != "landing_view" else None,
            "of_prev": round(100 * n / prev, 1) if prev else None,
        })
        prev = n if n else prev

    # By traffic source
    sources = [r[0] or "direct" for r in base.with_entities(FunnelEvent.source).distinct().all()]
    by_source = []
    for src in sorted(set(sources)):
        sq = base.filter(func.coalesce(FunnelEvent.source, "direct") == src)
        row = {"source": src}
        for e in ("landing_view", "demo", "signup", "first_document", "payment_approved"):
            row[e] = step_count(sq, e)
        row["conv"] = round(100 * row["signup"] / row["landing_view"], 1) if row["landing_view"] else None
        by_source.append(row)
    by_source.sort(key=lambda r: (r["signup"], r["landing_view"]), reverse=True)

    # Teacher vs parent split
    split = {}
    for p in ("teacher", "parent"):
        sq = base.filter(FunnelEvent.profile == p)
        split[p] = {e: sq.filter(FunnelEvent.event == e).count()
                    for e in ("demo", "signup", "first_document", "payment_approved")}

    # Day by day (most recent first, at most 30 rows)
    n_days = min(days or 30, 30)
    daily = []
    for i in range(n_days):
        d0 = (datetime.utcnow() - timedelta(days=i)).replace(hour=0, minute=0, second=0, microsecond=0)
        dq = FunnelEvent.query.filter(FunnelEvent.created_at >= d0,
                                      FunnelEvent.created_at < d0 + timedelta(days=1))
        row = {"day": d0.strftime("%d/%m")}
        for e in ("landing_view", "demo", "signup", "first_document"):
            row[e] = step_count(dq, e)
        daily.append(row)

    first_event = db.session.query(func.min(FunnelEvent.created_at)).scalar()
    users_by_profile = dict(db.session.query(func.coalesce(User.profile, "teacher"), func.count(User.id))
                            .filter(User.role == "user").group_by(func.coalesce(User.profile, "teacher")).all())

    return render_template("admin_funnel.html", steps=steps, by_source=by_source, split=split,
                           daily=daily, days=days, profile=profile, first_event=first_event,
                           users_by_profile=users_by_profile)


# ── ADMIN: failed generations ──────────────────────────────────────────────
@app.route("/admin/errors")
@login_required
def admin_errors():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))
    from sqlalchemy import func

    days = request.args.get("days", 7, type=int)
    if days not in (1, 7, 30, 0):
        days = 7
    eq = GenerationError.query
    dq = Document.query
    if days:
        since = datetime.utcnow() - timedelta(days=days)
        eq = eq.filter(GenerationError.created_at >= since)
        dq = dq.filter(Document.created_at >= since)

    total = eq.count()
    by_stage = dict(eq.with_entities(GenerationError.stage, func.count(GenerationError.id))
                    .group_by(GenerationError.stage).all())
    user_failures = by_stage.get("generate", 0) + by_stage.get("download", 0)
    successes = dq.count()
    rate = round(100 * user_failures / (user_failures + successes), 1) if (user_failures + successes) else None

    # Most frequent causes: group messages after removing variable numbers/ids
    causes = {}
    for (msg,) in eq.with_entities(GenerationError.message).all():
        key = re.sub(r"\d+", "#", (msg or "?").strip())[:140]
        causes[key] = causes.get(key, 0) + 1
    top_causes = sorted(causes.items(), key=lambda kv: kv[1], reverse=True)[:8]

    recent = eq.order_by(GenerationError.created_at.desc()).limit(50).all()
    return render_template("admin_errors.html", days=days, total=total, by_stage=by_stage,
                           successes=successes, rate=rate, top_causes=top_causes, recent=recent)


# ── ADMIN: database backups ─────────────────────────────────────────────────
@app.route("/admin/backups", methods=["GET", "POST"])
@login_required
def admin_backups():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        name = backup_database()
        log_admin_action("backup_created", target=name or "")
        db.session.commit()
        flash(tr("flash_backup_created"), "success")
        return redirect(url_for("admin_backups"))
    receipts = len(os.listdir(RECEIPTS_DIR)) if os.path.isdir(RECEIPTS_DIR) else 0
    return render_template("admin_backups.html", backups=list_backups(), keep=BACKUP_KEEP,
                           receipts_count=receipts)


@app.route("/admin/backups/<name>")
@login_required
def admin_backup_download(name):
    if current_user.role != "admin":
        return redirect(url_for("dashboard"))
    if not _BACKUP_NAME_RE.match(name) or not os.path.exists(os.path.join(BACKUP_DIR, name)):
        flash(tr("flash_backup_missing"), "error")
        return redirect(url_for("admin_backups"))
    log_admin_action("backup_downloaded", target=name)
    db.session.commit()
    return send_file(os.path.join(BACKUP_DIR, name), as_attachment=True, download_name=name)


@app.route("/admin/backups/full.zip")
@login_required
def admin_backup_full():
    """Fresh database snapshot + every payment receipt, in one zip (off-site copy)."""
    if current_user.role != "admin":
        return redirect(url_for("dashboard"))
    import tempfile, zipfile
    name = backup_database()
    tmp = tempfile.NamedTemporaryFile(suffix=".zip", delete=False, dir=BACKUP_DIR)
    tmp.close()
    with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as z:
        if name:
            z.write(os.path.join(BACKUP_DIR, name), f"eduprompt.db")
        if os.path.isdir(RECEIPTS_DIR):
            for f in os.listdir(RECEIPTS_DIR):
                fp = os.path.join(RECEIPTS_DIR, f)
                if os.path.isfile(fp):
                    z.write(fp, f"receipts/{f}")
    log_admin_action("backup_full_downloaded", target=name or "")
    db.session.commit()

    @after_this_request
    def _cleanup(response):
        try:
            os.remove(tmp.name)
        except OSError:
            pass
        return response

    return send_file(tmp.name, as_attachment=True,
                     download_name=datetime.utcnow().strftime("haithemeduai-sauvegarde-%Y%m%d.zip"))


# ── ADMIN: audit log (accountability trail of admin actions) ────────────────
@app.route("/admin/audit-log")
@login_required
def admin_audit_log():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    q = request.args.get("q", "").strip()
    page = request.args.get("page", 1, type=int)

    query = AdminLog.query
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(AdminLog.action.ilike(like), AdminLog.target.ilike(like),
                                     AdminLog.details.ilike(like)))
    query = query.order_by(AdminLog.created_at.desc())
    pager = query.paginate(page=page, per_page=30, error_out=False)

    return render_template("admin_audit_log.html", pager=pager, q=q)


def generate_password(length=12):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


# ── ADMIN: create & manage teacher accounts ──────────────────────────────────
@app.route("/admin/teachers", methods=["GET", "POST"])
@login_required
def admin_teachers():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    generated = None  # (email, password) shown once after creation

    if request.method == "POST":
        first_name = request.form.get("first_name", "").strip()
        last_name  = request.form.get("last_name", "").strip()
        email      = request.form.get("email", "").strip().lower()
        plan       = request.form.get("plan", "trial")
        custom_pw  = request.form.get("password", "").strip()
        api_key    = request.form.get("api_key", "").strip()

        if not all([first_name, last_name, email]):
            flash(tr("flash_teacher_fields_required"), "error")
        elif User.query.filter_by(email=email).first():
            flash(tr("flash_email_exists"), "error")
        else:
            password = custom_pw or generate_password()
            user = User(first_name=first_name, last_name=last_name, email=email,
                        gemini_api_key=api_key or None)
            user.set_password(password)
            db.session.add(user)
            db.session.flush()

            if plan == "trial":
                create_trial_subscription(user)
            else:
                db.session.add(Subscription(
                    user_id=user.id, plan=plan, status="active",
                    expires_at=datetime.utcnow() + timedelta(days=365),
                    docs_used=0, docs_used_today=0, last_reset_date=None,
                ))
                log_payment(user.id, plan, note="Created by admin")
                db.session.add(PlanChangeHistory(user_id=user.id, from_plan=None,
                                                  to_plan=plan, reason="admin_manual"))

            log_admin_action("teacher_created", target=email, details=f"plan={plan}")
            db.session.commit()
            generated = (email, password)
            flash(tr("flash_teacher_created", name=f"{first_name} {last_name}"), "success")

    q = request.args.get("q", "").strip()
    plan_filter = request.args.get("plan", "").strip()
    profile_filter = request.args.get("profile", "").strip()
    page = request.args.get("page", 1, type=int)

    query = User.query.filter_by(role="user")
    if profile_filter == "parent":
        query = query.filter(User.profile == "parent")
    elif profile_filter == "teacher":
        query = query.filter(db.or_(User.profile == "teacher", User.profile.is_(None)))
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(User.first_name.ilike(like),
                                     User.last_name.ilike(like),
                                     User.email.ilike(like)))
    if plan_filter:
        query = query.join(Subscription).filter(Subscription.plan == plan_filter)

    query = query.order_by(User.created_at.desc())
    pager = query.paginate(page=page, per_page=15, error_out=False)

    return render_template("admin_teachers.html", teachers=pager.items, pager=pager,
                            q=q, plan_filter=plan_filter, profile_filter=profile_filter,
                            generated=generated)


@app.route("/admin/teachers/export.csv")
@login_required
def admin_teachers_export():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    import csv, io as _io
    q = request.args.get("q", "").strip()
    plan_filter = request.args.get("plan", "").strip()

    query = User.query.filter_by(role="user")
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(User.first_name.ilike(like),
                                     User.last_name.ilike(like),
                                     User.email.ilike(like)))
    if plan_filter:
        query = query.join(Subscription).filter(Subscription.plan == plan_filter)
    teachers = query.order_by(User.created_at.desc()).all()

    buf = _io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["#", "Prénom", "Nom", "Email", "Abonnement", "Statut", "Expire le",
                      "Documents générés", "Créé le", "Actif"])
    for i, t in enumerate(teachers, start=1):
        sub = t.subscription
        writer.writerow([
            i, t.first_name, t.last_name, t.email,
            sub.plan if sub else "", sub.status if sub else "",
            sub.expires_at.strftime("%d/%m/%Y") if sub and sub.expires_at else "",
            len(t.documents), t.created_at.strftime("%d/%m/%Y"),
            "Oui" if t.is_active else "Non",
        ])
    log_admin_action("export_csv", target="teachers")
    db.session.commit()
    return app.response_class(buf.getvalue(), mimetype="text/csv",
                               headers={"Content-Disposition": "attachment; filename=enseignants.csv"})


@app.route("/admin/teachers/<int:user_id>")
@login_required
def admin_teacher_detail(user_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    teacher = User.query.get_or_404(user_id)
    sub = teacher.subscription

    q = request.args.get("q", "").strip()
    page = request.args.get("page", 1, type=int)

    doc_query = Document.query.filter_by(user_id=teacher.id)
    if q:
        like = f"%{q}%"
        doc_query = doc_query.filter(db.or_(Document.title.ilike(like),
                                             Document.subject.ilike(like),
                                             Document.doc_type.ilike(like),
                                             Document.level.ilike(like)))
    doc_query = doc_query.order_by(Document.created_at.desc())
    docs_pager = doc_query.paginate(page=page, per_page=10, error_out=False)

    payments = Payment.query.filter_by(user_id=teacher.id).order_by(Payment.created_at.desc()).all()
    total_docs = Document.query.filter_by(user_id=teacher.id).count()
    plan_history = PlanChangeHistory.query.filter_by(user_id=teacher.id) \
        .order_by(PlanChangeHistory.created_at.desc()).all()

    return render_template("admin_teacher_detail.html", teacher=teacher, sub=sub,
                            docs_pager=docs_pager, payments=payments, q=q,
                            total_docs=total_docs, plan_history=plan_history)


@app.route("/admin/teachers/<int:user_id>/reset-quota", methods=["POST"])
@login_required
def admin_reset_quota(user_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    teacher = User.query.get_or_404(user_id)
    sub = teacher.subscription
    if sub:
        sub.docs_used_today = 0
        sub.docs_used = 0
        sub.last_reset_date = None
        log_admin_action("quota_reset", target=teacher.email)
        db.session.commit()
        flash(tr("flash_quota_reset", name=teacher.full_name), "success")
    return redirect(url_for("admin_teacher_detail", user_id=user_id))


@app.route("/admin/teachers/<int:user_id>/extend", methods=["POST"])
@login_required
def admin_extend_subscription(user_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    teacher = User.query.get_or_404(user_id)
    sub = teacher.subscription
    days = request.form.get("days", 30, type=int)
    if sub:
        base = sub.expires_at if (sub.expires_at and sub.expires_at > datetime.utcnow()) else datetime.utcnow()
        sub.expires_at = base + timedelta(days=days)
        sub.status = "active"
        sub.expiry_reminder_sent = False
        sub.last_reminder_stage = None
        log_admin_action("subscription_extended", target=teacher.email, details=f"+{days}d")
        db.session.commit()
        flash(tr("flash_subscription_extended", name=teacher.full_name, days=days), "success")
    return redirect(url_for("admin_teacher_detail", user_id=user_id))


# ── ADMIN: cross-teacher document search ─────────────────────────────────────
@app.route("/admin/documents")
@login_required
def admin_documents():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    q = request.args.get("q", "").strip()
    doc_type = request.args.get("doc_type", "").strip()
    page = request.args.get("page", 1, type=int)

    query = Document.query.join(User, Document.user_id == User.id)
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(Document.title.ilike(like),
                                     Document.subject.ilike(like),
                                     Document.level.ilike(like),
                                     User.first_name.ilike(like),
                                     User.last_name.ilike(like),
                                     User.email.ilike(like)))
    if doc_type:
        query = query.filter(Document.doc_type == doc_type)

    query = query.order_by(Document.created_at.desc())
    pager = query.paginate(page=page, per_page=20, error_out=False)

    doc_type_options = [row[0] for row in
                         db.session.query(Document.doc_type).distinct().all() if row[0]]

    return render_template("admin_documents.html", pager=pager, q=q,
                            doc_type=doc_type, doc_type_options=doc_type_options)


@app.route("/admin/documents/export.csv")
@login_required
def admin_documents_export():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    import csv, io as _io
    q = request.args.get("q", "").strip()
    doc_type = request.args.get("doc_type", "").strip()

    query = Document.query.join(User, Document.user_id == User.id)
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(Document.title.ilike(like), Document.subject.ilike(like),
                                     Document.level.ilike(like), User.first_name.ilike(like),
                                     User.last_name.ilike(like), User.email.ilike(like)))
    if doc_type:
        query = query.filter(Document.doc_type == doc_type)
    docs = query.order_by(Document.created_at.desc()).all()

    buf = _io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["#", "Enseignant", "Email", "Titre", "Matière", "Niveau", "Type",
                      "Format", "Date"])
    for i, d in enumerate(docs, start=1):
        writer.writerow([i, d.user.full_name if d.user else "", d.user.email if d.user else "",
                          d.title, d.subject, d.level, d.doc_type, d.fmt,
                          d.created_at.strftime("%d/%m/%Y %H:%M")])
    log_admin_action("export_csv", target="documents")
    db.session.commit()
    return app.response_class(buf.getvalue(), mimetype="text/csv",
                               headers={"Content-Disposition": "attachment; filename=documents.csv"})


# ── MESSAGING: teacher <-> admin, free & unlimited ───────────────────────────
@app.route("/messages", methods=["GET", "POST"])
@login_required
def messages():
    if current_user.role == "admin":
        return redirect(url_for("admin_messages"))

    if request.method == "POST":
        body = request.form.get("body", "").strip()
        if body:
            db.session.add(Message(user_id=current_user.id, sender="teacher", body=body,
                                    read_by_teacher=True, read_by_admin=False))
            db.session.commit()
        return redirect(url_for("messages"))

    # Mark admin's replies as read now that the teacher is viewing the thread
    Message.query.filter_by(user_id=current_user.id, sender="admin", read_by_teacher=False) \
        .update({"read_by_teacher": True})
    db.session.commit()

    thread = Message.query.filter_by(user_id=current_user.id).order_by(Message.created_at.asc()).all()
    return render_template("messages.html", thread=thread)


@app.route("/admin/messages")
@login_required
def admin_messages():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    q = request.args.get("q", "").strip()
    page = request.args.get("page", 1, type=int)

    teachers_query = User.query.filter_by(role="user")
    if q:
        like = f"%{q}%"
        teachers_query = teachers_query.filter(db.or_(User.first_name.ilike(like),
                                                        User.last_name.ilike(like),
                                                        User.email.ilike(like)))
    teachers = teachers_query.all()

    threads = []
    for t in teachers:
        last = Message.query.filter_by(user_id=t.id).order_by(Message.created_at.desc()).first()
        unread = Message.query.filter_by(user_id=t.id, sender="teacher", read_by_admin=False).count()
        if last:  # only list teachers who have exchanged at least one message
            threads.append({"teacher": t, "last": last, "unread": unread})

    threads.sort(key=lambda x: (x["unread"] == 0, -x["last"].created_at.timestamp()))

    page_items, page, total_pages, total = paginate_list(threads, page, per_page=15)

    return render_template("admin_messages.html", threads=page_items, q=q,
                            page=page, total_pages=total_pages, total=total)


@app.route("/admin/messages/<int:user_id>", methods=["GET", "POST"])
@login_required
def admin_message_thread(user_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    teacher = User.query.get_or_404(user_id)

    if request.method == "POST":
        body = request.form.get("body", "").strip()
        if body:
            db.session.add(Message(user_id=teacher.id, sender="admin", body=body,
                                    read_by_admin=True, read_by_teacher=False))
            db.session.commit()
        return redirect(url_for("admin_message_thread", user_id=user_id))

    Message.query.filter_by(user_id=teacher.id, sender="teacher", read_by_admin=False) \
        .update({"read_by_admin": True})
    db.session.commit()

    thread = Message.query.filter_by(user_id=teacher.id).order_by(Message.created_at.asc()).all()
    return render_template("admin_message_thread.html", teacher=teacher, thread=thread)


@app.route("/admin/teachers/<int:user_id>/edit", methods=["GET", "POST"])
@login_required
def admin_edit_teacher(user_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    teacher = User.query.get_or_404(user_id)
    sub = teacher.subscription

    if request.method == "POST":
        first_name = request.form.get("first_name", "").strip()
        last_name  = request.form.get("last_name", "").strip()
        email      = request.form.get("email", "").strip().lower()
        plan       = request.form.get("plan", sub.plan if sub else "trial")
        api_key    = request.form.get("api_key", "").strip()

        existing = User.query.filter_by(email=email).first()
        if existing and existing.id != teacher.id:
            flash(tr("flash_email_exists"), "error")
            return render_template("admin_edit_teacher.html", teacher=teacher, sub=sub)

        teacher.first_name = first_name
        teacher.last_name = last_name
        teacher.email = email
        teacher.gemini_api_key = api_key or teacher.gemini_api_key

        if sub and plan != sub.plan:
            old_plan = sub.plan
            sub.plan = plan
            sub.docs_used_today = 0
            sub.last_reset_date = None
            sub.expiry_reminder_sent = False
            sub.last_reminder_stage = None
            if plan == "trial" and old_plan != "trial":
                sub.docs_used = 0
                sub.docs_limit = 5
                sub.expires_at = datetime.utcnow() + timedelta(days=14)
            elif plan != "trial":
                sub.expires_at = datetime.utcnow() + timedelta(days=365)
                log_payment(teacher.id, plan, note="Plan changed by admin")
            db.session.add(PlanChangeHistory(user_id=teacher.id, from_plan=old_plan,
                                              to_plan=plan, reason="admin_manual"))

        log_admin_action("teacher_updated", target=teacher.email)
        db.session.commit()
        flash(tr("flash_teacher_updated", name=teacher.full_name), "success")
        return redirect(url_for("admin_teachers"))

    return render_template("admin_edit_teacher.html", teacher=teacher, sub=sub)


@app.route("/admin/teachers/<int:user_id>/reset-password", methods=["POST"])
@login_required
def admin_reset_password(user_id):
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    user = User.query.get_or_404(user_id)
    new_password = generate_password()
    user.set_password(new_password)
    log_admin_action("password_reset", target=user.email)
    db.session.commit()
    flash(tr("flash_password_reset", email=user.email, password=new_password), "success")
    return redirect(url_for("admin_teachers"))


@app.route("/admin/teachers/<int:user_id>/delete", methods=["POST"])
@login_required
def admin_delete_teacher(user_id):
    """Soft-delete: deactivates the account (blocks login) rather than erasing it,
    so document history, payments and messages are preserved. Toggles back on if
    the account is already deactivated (reactivation)."""
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))

    user = User.query.get_or_404(user_id)
    redirect_to = request.form.get("redirect_to") or url_for("admin_teachers")

    if user.role == "admin":
        flash(tr("flash_admin_delete_protected"), "error")
    elif user.is_active:
        user.is_active = False
        log_admin_action("teacher_deactivated", target=user.email)
        db.session.commit()
        flash(tr("flash_teacher_deactivated", email=user.email), "success")
    else:
        user.is_active = True
        log_admin_action("teacher_reactivated", target=user.email)
        db.session.commit()
        flash(tr("flash_teacher_reactivated", email=user.email), "success")
    return redirect(redirect_to)


# ── MAIN GENERATOR PAGE ───────────────────────────────────────────────────────
@app.route("/generator")
@login_required
def generator():
    is_parent = current_user.is_parent
    return render_template("generator.html", levels=LEVELS, subjects=SUBJECTS,
                           doc_types=PARENT_DOCUMENT_TYPES if is_parent else DOCUMENT_TYPES,
                           is_parent=is_parent)


# ── API: QUOTA CHECK ──────────────────────────────────────────────────────────
@app.route("/api/check-quota")
@login_required
def check_quota():
    sub = current_user.subscription
    allowed = sub.has_quota()
    db.session.commit()  # persist any daily-counter roll-over from has_quota()
    return jsonify({
        "allowed": allowed,
        "remaining": sub.remaining(),
        "plan": sub.plan,
        "plan_label": sub.plan_label(),
        "is_daily": sub.is_daily_plan(),
        "daily_limit": sub.daily_limit() if sub.is_daily_plan() else None,
        "contact_email": CONTACT_EMAIL,
    })


# ── API: LOG ERROR (a generation failed in front of the user) ───────────────
@app.route("/api/log-error", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("60/hour")
def log_error():
    data = request.get_json(silent=True) or {}
    stage = str(data.get("stage", "generate"))[:20]
    db.session.add(GenerationError(
        user_id=current_user.id, stage=stage if stage in ("generate", "download", "sources") else "generate",
        message=str(data.get("message", ""))[:500], doc_type=str(data.get("doc_type", ""))[:40],
        subject=str(data.get("subject", ""))[:80]))
    db.session.commit()
    return jsonify({"ok": True})


# ── API: LOG USAGE (after successful AI generation) ──────────────────────────
@app.route("/api/log-usage", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("60/hour")
def log_usage():
    data = request.get_json() or {}
    sub = current_user.subscription

    if not sub.has_quota():
        return jsonify({"error": "quota_exceeded", "contact_email": CONTACT_EMAIL}), 403

    sub.register_usage()
    if Document.query.filter_by(user_id=current_user.id).count() == 0:
        track("first_document", profile=current_user.profile, user_id=current_user.id)
        reward_referrer_if_due(current_user)

    doc = Document(
        user_id=current_user.id,
        title=data.get("lesson", "")[:255],
        subject=data.get("subject", ""),
        level=data.get("level", ""),
        palier=data.get("palier", ""),
        doc_type=data.get("doc_type", ""),
        fmt=data.get("format", "docx"),
    )
    if isinstance(data.get("content"), str):
        doc.content = data["content"][:200000]
        doc.payload = _document_payload(data)
    db.session.add(doc)
    db.session.commit()

    return jsonify({"ok": True, "remaining": sub.remaining(), "doc_id": doc.id})


DOC_PAYLOAD_MAX = 2_500_000  # bytes; beyond this, AI images are not kept in the history


def _document_payload(data, previous=None):
    """JSON needed to rebuild the Word/PDF file later (header, colours, visuals)."""
    import json as _json
    p = _json.loads(previous) if previous else {}
    for k in ("lang", "teacher_first", "teacher_last", "level", "palier", "subject", "lesson",
              "etablissement", "projet", "activite"):
        if k in data:
            p[k] = str(data.get(k) or "")[:300]
    if isinstance(data.get("colors"), dict):
        p["colors"] = data["colors"]
    if isinstance(data.get("visuals"), list):
        p["visuals"] = data["visuals"]
    out = _json.dumps(p, ensure_ascii=False)
    if len(out) > DOC_PAYLOAD_MAX:  # keep the text, drop the heaviest pictures
        p["visuals"] = [v for v in p.get("visuals", []) if len(str(v.get("data", ""))) < 300_000]
        out = _json.dumps(p, ensure_ascii=False)
        if len(out) > DOC_PAYLOAD_MAX:
            p["visuals"] = []
            out = _json.dumps(p, ensure_ascii=False)
    return out


# ── API: GEMINI PROXY (platform key — users no longer need their own) ────────
# The browser never sees an API key: it posts the Gemini request body here and
# the server forwards it with the teacher's personal key if they set one, or
# the platform key (GEMINI_API_KEY env var) otherwise. The response is passed
# back unchanged (same JSON shape and status), so client retry/fallback logic
# keeps working as before.
_GEMINI_MODEL_RE = re.compile(r"^gemini-[a-z0-9.\-]{1,60}$")


def _ai_rate_key():
    return f"user:{current_user.id}" if current_user.is_authenticated else get_remote_address()


@app.route("/api/ai/generate", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("150/hour", key_func=_ai_rate_key)
def ai_generate():
    data = request.get_json(silent=True) or {}
    resp, err = _gemini_call(data, stream=False)
    if err is not None:
        return err
    return app.response_class(resp.content, status=resp.status_code, mimetype="application/json")


@app.route("/api/ai/stream", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("150/hour", key_func=_ai_rate_key)
def ai_stream():
    """Same as /api/ai/generate, but relays Gemini's text as it is written
    (server-sent events), so the user sees the document appear progressively."""
    data = request.get_json(silent=True) or {}
    # The client sends its list of models; the server walks it itself, so an
    # overloaded or exhausted model costs a fraction of a second instead of a
    # round trip and a visible wait for the teacher.
    models = data.get("models")
    if not isinstance(models, list) or not models:
        models = [data.get("model", "")]
    models = [str(m) for m in models[:8]]
    resp = err = None
    for model in models:
        resp, err = _gemini_call(dict(data, model=model), stream=True)
        if err is not None:
            status = err[1] if isinstance(err, tuple) else 500
            if status == 503 and model != models[-1]:
                continue  # network hiccup towards Google: try the next model
            return err
        if resp.status_code == 200:
            break
        body = resp.content
        resp.close()
        _log_upstream_error({"model": model}, resp.status_code, body)
        if resp.status_code in (404, 429, 500, 503, 504) and model != models[-1]:
            continue  # unavailable, overloaded or out of quota: next model at once
        return app.response_class(body, status=resp.status_code, mimetype="application/json")
    used_model = model

    def relay():
        try:
            for chunk in resp.iter_content(chunk_size=None):
                if chunk:
                    yield chunk
        finally:
            resp.close()

    return app.response_class(relay(), mimetype="text/event-stream",
                              headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no",
                                       "X-Gemini-Model": used_model})


def _log_upstream_error(data, status, body):
    """One log line per Gemini refusal (model, status, Google's message), so the
    real cause of 429/503 waves is visible in the Railway logs."""
    try:
        msg = (json.loads(body).get("error") or {}).get("message", "")
    except Exception:
        msg = body[:200].decode("utf-8", "ignore") if isinstance(body, bytes) else str(body)[:200]
    print(f"⚠️ Gemini {status} on {str(data.get('model',''))[:40]}: {msg[:220]}")


def _gemini_call(data, stream=False):
    """Send one Gemini request with the right key. Returns (response, None), where
    the response may be a Google error to pass through (404, 429…), or
    (None, flask_error_response) when we answer ourselves (bad input, quota…)."""
    import requests as http

    model = str(data.get("model", ""))
    contents = data.get("contents")
    if not _GEMINI_MODEL_RE.match(model) or not isinstance(contents, list):
        return None, (jsonify({"error": {"message": "invalid request"}}), 400)

    personal_key = (current_user.gemini_api_key or "").strip()
    platform_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not personal_key and not platform_key:
        return None, (jsonify({"error": {"message": "AI service not configured"}}), 503)

    body = {"contents": contents}
    if isinstance(data.get("generationConfig"), dict):
        body["generationConfig"] = data["generationConfig"]
    method = "streamGenerateContent?alt=sse" if stream else "generateContent"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:{method}"

    def call(key):
        return http.post(url, headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                         json=body, timeout=(15, 240), stream=stream)

    # 1) Older accounts may still have their own key: try it first (free for us).
    if personal_key:
        try:
            resp = call(personal_key)
        except http.RequestException:
            resp = None
        if resp is not None and not _personal_key_failed(resp):
            return resp, None
        # The personal key is invalid, revoked or out of quota: fall back silently
        # to the platform key so the user never sees an API-key problem again.
        if resp is not None and _personal_key_invalid(resp):
            current_user.gemini_api_key = None  # broken for good: stop retrying it
            db.session.commit()
            app.logger.info("Cleared invalid personal Gemini key for user %s", current_user.id)
        if not platform_key:
            if resp is None:
                return None, (jsonify({"error": {"message": "upstream unavailable"}}), 503)
            return resp, None

    # 2) Platform key: billed to us, so the plan quota applies.
    sub = current_user.subscription
    if not sub or not sub.has_quota():
        db.session.commit()
        return None, (jsonify({"error": {"message": "quota_exceeded"}}), 403)
    try:
        resp = call(platform_key)
    except http.RequestException:
        return None, (jsonify({"error": {"message": "upstream unavailable"}}), 503)

    # Never let the platform key's auth errors look like the user's problem.
    if resp.status_code in (400, 401, 403) and "API key" in resp.text:
        app.logger.error("Platform GEMINI_API_KEY rejected by Google: %s", resp.text[:300])
        db.session.add(GenerationError(user_id=current_user.id, stage="upstream",
                                       message="Platform Gemini key rejected: " + resp.text[:400]))
        db.session.commit()
        return None, (jsonify({"error": {"message": "AI service temporarily unavailable"}}), 503)
    return resp, None


def _personal_key_invalid(resp):
    """Google says the key itself is bad (wrong, deleted, API disabled)."""
    if resp.status_code not in (400, 401, 403):
        return False
    text = resp.text
    return ("API key" in text or "API_KEY" in text or "PERMISSION_DENIED" in text
            or "SERVICE_DISABLED" in text or resp.status_code == 401)


def _personal_key_failed(resp):
    """Any failure that is the personal key's fault: invalid, or its own quota is exhausted."""
    return _personal_key_invalid(resp) or resp.status_code == 429


# ── API: PUBLIC DEMO (show value before signup) ─────────────────────────────
# Anonymous visitors can generate ONE short lesson preview from the landing
# page. The prompt is built server-side from whitelisted fields only, so this
# endpoint cannot be used as a general-purpose Gemini proxy. Cost is bounded
# by a per-IP limit plus a global daily cap.
DEMO_MODELS = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"]
DEMO_DAILY_CAP = int(os.environ.get("DEMO_DAILY_CAP", "300"))
_demo_counter = {"day": None, "count": 0}

_DEMO_PROMPT_PARENT = {
    "fr": ("Tu aides un parent algérien à accompagner son enfant à la maison. Rédige un "
           "APERÇU COURT de fiche de révision (300 mots maximum), en français, pour un enfant "
           "de niveau « {level} », matière « {subject} », leçon « {lesson} », conforme au "
           "programme algérien. Mots simples, ton bienveillant. Structure en Markdown : "
           "## titre, ### L'essentiel (3 puces), ### Un exemple du quotidien, "
           "### Mini-quiz (3 questions puis les réponses). Rien d'autre."),
    "en": ("You help an Algerian parent support their child at home. Write a SHORT PREVIEW "
           "of a revision sheet (300 words max), in English, for a child at level \"{level}\", "
           "subject \"{subject}\", lesson \"{lesson}\", aligned with the Algerian curriculum. "
           "Simple words, warm tone. Markdown structure: ## title, ### Key points (3 bullets), "
           "### An everyday example, ### Mini-quiz (3 questions then the answers). Nothing else."),
    "ar": ("أنت تساعد وليّ أمر جزائري على مرافقة طفله في البيت. اكتب معاينة مختصرة لبطاقة "
           "مراجعة (300 كلمة كحد أقصى) باللغة العربية، لطفل في مستوى «{level}»، المادة "
           "«{subject}»، الدرس «{lesson}»، مطابقة للمنهاج الجزائري. كلمات بسيطة وأسلوب لطيف. "
           "البنية بصيغة Markdown: ## العنوان، ### الأهم (3 نقاط)، ### مثال من الحياة اليومية، "
           "### اختبار قصير (3 أسئلة ثم الأجوبة). لا شيء غير ذلك."),
}

_DEMO_PROMPT = {
    "fr": ("Tu es un expert en ingénierie pédagogique en Algérie. Rédige un APERÇU COURT "
           "de fiche de cours (350 mots maximum), en français, niveau « {level} », matière "
           "« {subject} », leçon « {lesson} ». Structure en Markdown : ## titre, ### Objectifs "
           "(3 puces), ### Déroulement (3 étapes courtes), ### Activité (1 exercice). "
           "Pas d'introduction ni de conclusion hors structure."),
    "en": ("You are an instructional design expert in Algeria. Write a SHORT PREVIEW of a "
           "lesson plan (350 words max), in English, level \"{level}\", subject \"{subject}\", "
           "lesson \"{lesson}\". Markdown structure: ## title, ### Objectives (3 bullets), "
           "### Lesson flow (3 short steps), ### Activity (1 exercise). Nothing outside it."),
    "ar": ("أنت خبير في الهندسة البيداغوجية في الجزائر. اكتب معاينة مختصرة لمذكرة درس "
           "(350 كلمة كحد أقصى) باللغة العربية، المستوى «{level}»، المادة «{subject}»، "
           "الدرس «{lesson}». البنية بصيغة Markdown: ## العنوان، ### الأهداف (3 نقاط)، "
           "### سير الدرس (3 مراحل قصيرة)، ### نشاط (تمرين واحد). لا شيء خارج هذه البنية."),
}


@app.route("/api/demo", methods=["POST"])
@csrf.exempt
@limiter.limit("3/day;2/hour", deduct_when=lambda response: response.status_code == 200)
def api_demo():
    import requests as http

    data = request.get_json(silent=True) or {}
    lang = data.get("lang") if data.get("lang") in LEVELS else current_lang()
    level = str(data.get("level", "")).strip()
    subject = str(data.get("subject", "")).strip()
    lesson = " ".join(str(data.get("lesson", "")).split())[:120]
    if level not in LEVELS[lang] or subject not in SUBJECTS[lang] or len(lesson) < 3:
        return jsonify({"error": "invalid"}), 400

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return jsonify({"error": "unavailable"}), 503

    today = datetime.utcnow().date()
    if _demo_counter["day"] != today:
        _demo_counter.update(day=today, count=0)
    if _demo_counter["count"] >= DEMO_DAILY_CAP:
        return jsonify({"error": "busy"}), 429
    _demo_counter["count"] += 1

    templates = _DEMO_PROMPT_PARENT if data.get("profile") == "parent" else _DEMO_PROMPT
    prompt = templates[lang].format(level=level, subject=subject, lesson=lesson)
    body = {"contents": [{"role": "user", "parts": [{"text": prompt}]}],
            "generationConfig": {"maxOutputTokens": 1200, "temperature": 0.7}}
    for model in DEMO_MODELS:
        try:
            resp = http.post(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
                json=body, timeout=60)
        except http.RequestException:
            continue
        if resp.status_code != 200:
            continue
        try:
            text = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
        except (ValueError, KeyError, IndexError, TypeError):
            continue
        track("demo", profile="parent" if data.get("profile") == "parent" else "teacher")
        db.session.commit()
        return jsonify({"text": text})
    db.session.add(GenerationError(stage="demo", message="All demo models failed",
                                   subject=subject[:80]))
    db.session.commit()
    return jsonify({"error": "unavailable"}), 503


# ── API: EXTRACT SOURCES (NotebookLM-style: files + links) ──────────────────
@app.route("/api/extract-sources", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("30/hour")
def extract_sources():
    files = request.files.getlist("files")
    urls_raw = request.form.get("urls", "")
    urls = [u.strip() for u in urls_raw.split("\n") if u.strip()]

    context, warnings, attachments = build_sources_context(files, urls)
    return jsonify({"context": context, "warnings": warnings,
                    "attachments": [{"mimeType": a["mime"], "data": a["data"], "name": a["name"]}
                                    for a in attachments]})


# ── DOCUMENT GENERATION HELPERS ───────────────────────────────────────────────
def hex_to_rgb(hex_color, default="102a53"):
    h = (hex_color or default).lstrip("#")
    if len(h) != 6:
        h = default
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def set_cell_background(cell, hex_color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color.lstrip("#"))
    cell._tc.get_or_add_tcPr().append(shd)


def try_parse_md_table(lines, i):
    """If lines[i] starts a Markdown pipe-table, parse it and return (rows, next_index).
    Otherwise return None. Handles the |---|---| separator row Gemini emits."""
    def is_row(s):
        s = s.strip()
        return s.startswith("|") and s.endswith("|") and s.count("|") >= 2

    def is_sep(s):
        s = s.strip().strip("|")
        return bool(s) and all(c in "-:| " for c in s) and "-" in s

    if i + 1 >= len(lines) or not is_row(lines[i]) or not is_sep(lines[i + 1]):
        return None

    def split_row(s):
        s = s.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        # Gemini often puts <br> inside cells to break lines: keep the line break, drop the tag.
        return [re.sub(r"<br\s*/?>", "\n", c, flags=re.I).replace("**", "").strip() for c in s.split("|")]

    rows = [split_row(lines[i])]
    j = i + 2
    while j < len(lines) and is_row(lines[j]):
        rows.append(split_row(lines[j]))
        j += 1
    return rows, j


HEADER_LABELS = {
    "fr": {"teacher": "Enseignant(e)", "etablissement": "Établissement", "level": "Niveau",
           "palier": "Palier", "subject": "Matière", "projet": "Projet / Unité", "activite": "Activité"},
    "en": {"teacher": "Teacher", "etablissement": "Institution", "level": "Level",
           "palier": "Grade", "subject": "Subject", "projet": "Project / Unit", "activite": "Activity"},
    "ar": {"teacher": "الأستاذ(ة)", "etablissement": "المؤسسة", "level": "المستوى",
           "palier": "الطور", "subject": "المادة", "projet": "المشروع / الوحدة", "activite": "النشاط"},
}


def _visuals_by_id(visuals):
    """Build a lookup dict {id: base64_png_bytes} from the client-supplied visuals list.
    Tolerant of missing/malformed entries — a bad entry is simply skipped."""
    out = {}
    for v in (visuals or []):
        try:
            vid = int(v.get("id"))
            data_url = v.get("data", "")
            b64 = data_url.split(",", 1)[1] if "," in data_url else data_url
            out[vid] = base64.b64decode(b64)
        except Exception:
            continue
    return out


VISUAL_ID_RE = re.compile(r'\[\[VISUAL_ID:(\d+)\]\]')


_HR_RE = re.compile(r"^(?:-{3,}|\*{3,}|_{3,})$")
_BULLET_RE = re.compile(r"^(\s*)[-*•+]\s+(.*)$")


def _list_item(raw_line):
    """'- x', '* x', '• x' (with indentation for sub-levels) -> (level, text), else None."""
    m = _BULLET_RE.match(raw_line.replace("\t", "    "))
    if not m or not m.group(2).strip():
        return None
    indent = len(m.group(1))
    level = 0 if indent < 2 else (1 if indent < 6 else 2)  # 2 or 4 spaces = sub-level
    return level, m.group(2).strip()


def make_docx(content, meta, colors_cfg, lang="fr", visuals=None):
    """
    meta: dict with teacher_first, teacher_last, level, palier, subject, lesson
    colors_cfg: dict with c1 (header bg), c2 (H1), c3 (accent/H2), text (body text)
    """
    c1 = colors_cfg.get("c1", "#102a53")
    c2 = colors_cfg.get("c2", "#f97316")
    c3 = colors_cfg.get("c3", "#173f6c")
    ctext = colors_cfg.get("text", "#1a1a1a")

    doc = DocxDocument()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # ── Brand title ──
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("HaithemEduAI")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(*hex_to_rgb(c1))
    title_p.paragraph_format.space_after = Pt(14)

    # ── Header table: vertical label:value rows (scales to any number of fields) ──
    labels = HEADER_LABELS.get(lang, HEADER_LABELS["fr"])
    rows_data = [
        (labels["teacher"], f"{meta.get('teacher_first','')} {meta.get('teacher_last','')}".strip() or "—"),
        (labels["etablissement"], meta.get("etablissement") or "—"),
        (labels["level"], meta.get("level") or "—"),
        (labels["palier"], meta.get("palier") or "—"),
        (labels["subject"], meta.get("subject") or "—"),
        (labels["projet"], meta.get("projet") or "—"),
        (labels["activite"], meta.get("activite") or "—"),
    ]

    table = doc.add_table(rows=len(rows_data), cols=2)
    table.style = "Table Grid"
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(4.2)

    for i, (h, v) in enumerate(rows_data):
        hc = table.cell(i, 0)
        hc.text = ""
        hp = hc.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT if lang != "ar" else WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run(h)
        hr.bold = True
        hr.font.size = Pt(10)
        hr.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(hc, c1)
        hc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        vc = table.cell(i, 1)
        vc.text = ""
        vp = vc.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.LEFT if lang != "ar" else WD_ALIGN_PARAGRAPH.RIGHT
        vr = vp.add_run(v)
        vr.font.size = Pt(11)
        vr.font.color.rgb = RGBColor(*hex_to_rgb(ctext, "1a1a1a"))
        vc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ── Lesson topic banner ──
    lesson_p = doc.add_paragraph()
    lesson_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lr = lesson_p.add_run(meta.get("lesson", ""))
    lr.bold = True
    lr.font.size = Pt(14)
    lr.font.color.rgb = RGBColor(*hex_to_rgb(c3))
    lesson_p.paragraph_format.space_after = Pt(16)

    # ── Body content (markdown-ish parsing, with real tables) ──
    visuals_map = _visuals_by_id(visuals)
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if not line or _HR_RE.match(line):  # Markdown "---" separators: just space
            doc.add_paragraph("")
            i += 1
            continue

        visual_match = VISUAL_ID_RE.fullmatch(line)
        if visual_match:
            img_bytes = visuals_map.get(int(visual_match.group(1)))
            if img_bytes:
                try:
                    img_p = doc.add_paragraph()
                    img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    img_p.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(5.2))
                except Exception:
                    pass
            i += 1
            continue

        table_result = try_parse_md_table(lines, i)
        if table_result:
            rows, i = table_result
            n_cols = max(len(r) for r in rows)
            tbl = doc.add_table(rows=len(rows), cols=n_cols)
            tbl.style = "Table Grid"
            for ri, row in enumerate(rows):
                for ci in range(n_cols):
                    cell = tbl.cell(ri, ci)
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "ar" else WD_ALIGN_PARAGRAPH.LEFT
                    r = p.add_run(row[ci] if ci < len(row) else "")
                    r.font.size = Pt(10)
                    if ri == 0:
                        r.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
                        set_cell_background(cell, c3)
                    else:
                        r.font.color.rgb = RGBColor(*hex_to_rgb(ctext, "1a1a1a"))
            doc.add_paragraph().paragraph_format.space_after = Pt(10)
            continue

        if line.startswith("### "):
            h = doc.add_heading("", level=3)
            r = h.add_run(line[4:])
            r.font.color.rgb = RGBColor(*hex_to_rgb(c3))
        elif line.startswith("## "):
            h = doc.add_heading("", level=2)
            r = h.add_run(line[3:])
            r.font.color.rgb = RGBColor(*hex_to_rgb(c3))
        elif line.startswith("# "):
            h = doc.add_heading("", level=1)
            r = h.add_run(line[2:])
            r.font.color.rgb = RGBColor(*hex_to_rgb(c2))
        elif _list_item(lines[i]):
            level, item = _list_item(lines[i])
            style = ["List Bullet", "List Bullet 2", "List Bullet 3"][level]
            try:
                p = doc.add_paragraph(style=style)
            except KeyError:
                p = doc.add_paragraph(style="List Bullet")
            for pi, part in enumerate(item.split("**")):
                r = p.add_run(part)
                r.font.color.rgb = RGBColor(*hex_to_rgb(ctext, "1a1a1a"))
                if pi % 2 == 1:
                    r.bold = True
        else:
            p = doc.add_paragraph()
            parts = line.split("**")
            for pi, part in enumerate(parts):
                r = p.add_run(part)
                r.font.color.rgb = RGBColor(*hex_to_rgb(ctext, "1a1a1a"))
                if pi % 2 == 1:
                    r.bold = True

        i += 1

    # ── Footer ──
    doc.add_paragraph().paragraph_format.space_before = Pt(20)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer_p.add_run("Généré avec HaithemEduAI")
    fr.font.size = Pt(8)
    fr.font.color.rgb = RGBColor(150, 150, 150)
    fr.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def make_pdf(content, meta, colors_cfg, lang="fr", visuals=None):
    c1 = colors_cfg.get("c1", "#102a53")
    c2 = colors_cfg.get("c2", "#f97316")
    c3 = colors_cfg.get("c3", "#173f6c")
    ctext = colors_cfg.get("text", "#1a1a1a")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=1.6*cm, bottomMargin=1.6*cm)
    styles = getSampleStyleSheet()
    align = TA_RIGHT if lang == "ar" else TA_LEFT

    title_s = ParagraphStyle("T", parent=styles["Title"], textColor=rl_colors.HexColor(c1),
                              fontSize=20, spaceAfter=10, alignment=TA_CENTER)
    lesson_s = ParagraphStyle("L", parent=styles["Normal"], textColor=rl_colors.HexColor(c3),
                               fontSize=14, spaceAfter=14, alignment=TA_CENTER, fontName="Helvetica-Bold")
    h1_s = ParagraphStyle("H1", parent=styles["Heading1"], textColor=rl_colors.HexColor(c2),
                           fontSize=14, spaceBefore=14, spaceAfter=4, alignment=align)
    h2_s = ParagraphStyle("H2", parent=styles["Heading2"], textColor=rl_colors.HexColor(c3),
                           fontSize=12, spaceBefore=10, spaceAfter=3, alignment=align)
    body_s = ParagraphStyle("B", parent=styles["Normal"], textColor=rl_colors.HexColor(ctext),
                             fontSize=10, leading=16, spaceAfter=5, alignment=align)
    blt_s = ParagraphStyle("Bl", parent=styles["Normal"], textColor=rl_colors.HexColor(ctext),
                            fontSize=10, leading=16, leftIndent=20, spaceAfter=3, alignment=align)

    story = [Paragraph("HaithemEduAI", title_s)]

    teacher = f"{meta.get('teacher_first','')} {meta.get('teacher_last','')}".strip() or "—"
    labels = HEADER_LABELS.get(lang, HEADER_LABELS["fr"])
    table_data = [
        [labels["teacher"], teacher],
        [labels["etablissement"], meta.get("etablissement") or "—"],
        [labels["level"], meta.get("level") or "—"],
        [labels["palier"], meta.get("palier") or "—"],
        [labels["subject"], meta.get("subject") or "—"],
        [labels["projet"], meta.get("projet") or "—"],
        [labels["activite"], meta.get("activite") or "—"],
    ]
    if lang == "ar":
        table_data = [[row[1], row[0]] for row in table_data]  # value first visually for RTL

    tbl = Table(table_data, colWidths=[5*cm, 9*cm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1) if lang != "ar" else (1,-1), rl_colors.HexColor(c1)),
        ("TEXTCOLOR", (0,0), (0,-1) if lang != "ar" else (1,-1), rl_colors.white),
        ("TEXTCOLOR", (1,0), (1,-1) if lang != "ar" else (0,-1), rl_colors.HexColor(ctext)),
        ("FONTNAME", (0,0), (0,-1) if lang != "ar" else (1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (0,0), (-1,-1), "RIGHT" if lang == "ar" else "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.75, rl_colors.HexColor("#dce6f0")),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(meta.get("lesson",""), lesson_s))

    cell_s = ParagraphStyle("Cell", parent=styles["Normal"], textColor=rl_colors.HexColor(ctext),
                             fontSize=9, leading=13, alignment=align)
    cell_head_s = ParagraphStyle("CellH", parent=styles["Normal"], textColor=rl_colors.white,
                                  fontSize=9, leading=13, alignment=align, fontName="Helvetica-Bold")

    visuals_map = _visuals_by_id(visuals)
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if not line or _HR_RE.match(line):
            story.append(Spacer(1, 0.2*cm))
            i += 1
            continue

        visual_match = VISUAL_ID_RE.fullmatch(line)
        if visual_match:
            img_bytes = visuals_map.get(int(visual_match.group(1)))
            if img_bytes:
                try:
                    from reportlab.platypus import Image as RLImage
                    rl_img = RLImage(io.BytesIO(img_bytes), width=13*cm, height=9*cm, kind="proportional")
                    rl_img.hAlign = "CENTER"
                    story.append(rl_img)
                    story.append(Spacer(1, 0.3*cm))
                except Exception:
                    pass
            i += 1
            continue

        table_result = try_parse_md_table(lines, i)
        if table_result:
            rows, i = table_result
            n_cols = max(len(r) for r in rows)
            wrapped = []
            for ri, row in enumerate(rows):
                style = cell_head_s if ri == 0 else cell_s
                wrapped.append([Paragraph(row[ci] if ci < len(row) else "", style) for ci in range(n_cols)])
            col_w = (16*cm) / n_cols
            body_tbl = Table(wrapped, colWidths=[col_w]*n_cols)
            body_tbl.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), rl_colors.HexColor(c3)),
                ("GRID", (0,0), (-1,-1), 0.6, rl_colors.HexColor("#dce6f0")),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 5),
                ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ]))
            story.append(body_tbl)
            story.append(Spacer(1, 0.3*cm))
            continue

        if line.startswith("### ") or line.startswith("## ") or line.startswith("# "):
            style = h1_s if line.startswith("# ") else h2_s
            story.append(Paragraph(line.lstrip("#").strip(), style))
        elif _list_item(lines[i]):
            level, item = _list_item(lines[i])
            safe = item.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            safe = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', safe)
            mark = ["•", "◦", "▪"][level]
            pad = "&nbsp;" * (6 * level)
            story.append(Paragraph(f"{pad}{mark} {safe}", blt_s))
        else:
            safe = line.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            safe = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', safe)
            story.append(Paragraph(safe, body_s))

        i += 1

    doc.build(story)
    buf.seek(0)
    return buf


# ── API: DOWNLOAD ──────────────────────────────────────────────────────────────
@app.route("/api/download", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("60/hour")
def download():
    data = request.get_json()
    content = data.get("content", "")
    lang    = data.get("lang", "fr")
    fmt     = data.get("format", "docx")

    meta = {
        "teacher_first": data.get("teacher_first", current_user.first_name),
        "teacher_last":  data.get("teacher_last", current_user.last_name),
        "level":         data.get("level", ""),
        "palier":        data.get("palier", ""),
        "subject":       data.get("subject", ""),
        "lesson":        data.get("lesson", ""),
        "etablissement": data.get("etablissement", ""),
        "projet":        data.get("projet", ""),
        "activite":      data.get("activite", ""),
    }
    colors_cfg = data.get("colors", {})
    visuals = data.get("visuals", [])

    # Keep the history copy in sync with what the user actually downloaded
    doc_id = data.get("doc_id")
    if isinstance(doc_id, int):
        doc = Document.query.filter_by(id=doc_id, user_id=current_user.id).first()
        if doc:
            doc.content = content[:200000]
            doc.fmt = fmt if fmt in ("docx", "pdf") else doc.fmt
            doc.payload = _document_payload(data, doc.payload)
            db.session.commit()

    safe_name = re.sub(r'[^\w\-]', '_', meta["lesson"][:40] or "document")

    if fmt == "pdf":
        buf = make_pdf(content, meta, colors_cfg, lang, visuals=visuals)
        return send_file(buf, as_attachment=True, download_name=f"{safe_name}.pdf",
                         mimetype="application/pdf")
    else:
        buf = make_docx(content, meta, colors_cfg, lang, visuals=visuals)
        return send_file(buf, as_attachment=True, download_name=f"{safe_name}.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ── MY DOCUMENTS (history, re-download without quota, rating) ───────────────
@app.route("/mes-documents")
@login_required
def my_documents():
    q = request.args.get("q", "").strip()
    page = request.args.get("page", 1, type=int)
    query = Document.query.filter_by(user_id=current_user.id)
    if q:
        like = f"%{q}%"
        query = query.filter(db.or_(Document.title.ilike(like), Document.subject.ilike(like)))
    pager = query.order_by(Document.created_at.desc()).paginate(page=page, per_page=15, error_out=False)
    labels = {d["id"]: d["label"] for lst in (DOCUMENT_TYPES, PARENT_DOCUMENT_TYPES)
              for d in lst.get(current_lang(), lst["fr"])}
    return render_template("my_documents.html", pager=pager, q=q, type_labels=labels)


@app.route("/mes-documents/<int:doc_id>/<fmt>")
@login_required
def my_document_download(doc_id, fmt):
    import json as _json
    doc = Document.query.filter_by(id=doc_id, user_id=current_user.id).first_or_404()
    if not doc.content or fmt not in ("docx", "pdf"):
        flash(tr("history_not_available"), "error")
        return redirect(url_for("my_documents"))
    p = _json.loads(doc.payload) if doc.payload else {}
    meta = {k: p.get(k, "") for k in ("level", "palier", "subject", "lesson",
                                      "etablissement", "projet", "activite")}
    meta["teacher_first"] = p.get("teacher_first") or current_user.first_name
    meta["teacher_last"] = p.get("teacher_last") or current_user.last_name
    meta["lesson"] = meta["lesson"] or doc.title or ""
    meta["subject"] = meta["subject"] or doc.subject or ""
    meta["level"] = meta["level"] or doc.level or ""
    lang = p.get("lang") or current_lang()
    colors_cfg = p.get("colors") or {}
    visuals = p.get("visuals") or []
    safe_name = re.sub(r'[^\w\-]', '_', (meta["lesson"] or "document")[:40])
    if fmt == "pdf":
        buf = make_pdf(doc.content, meta, colors_cfg, lang, visuals=visuals)
        return send_file(buf, as_attachment=True, download_name=f"{safe_name}.pdf",
                         mimetype="application/pdf")
    buf = make_docx(doc.content, meta, colors_cfg, lang, visuals=visuals)
    return send_file(buf, as_attachment=True, download_name=f"{safe_name}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/api/documents/<int:doc_id>/rate", methods=["POST"])
@login_required
@csrf.exempt
@limiter.limit("120/hour")
def rate_document(doc_id):
    doc = Document.query.filter_by(id=doc_id, user_id=current_user.id).first_or_404()
    data = request.get_json(silent=True) or request.form
    try:
        value = int(data.get("rating", 0))
    except (TypeError, ValueError):
        value = 0
    if value not in (1, -1):
        return jsonify({"error": "invalid"}), 400
    doc.rating = value
    comment = str(data.get("comment", "") or "").strip()[:500]
    if comment or value == 1:
        doc.rating_comment = comment or None
    doc.rated_at = datetime.utcnow()
    db.session.commit()
    return jsonify({"ok": True})


# ── Follow-up emails: unsubscribe link + admin page ──────────────────────────
@app.route("/email/stop/<token>")
def email_optout(token):
    try:
        uid = _optout_serializer().loads(token)
    except BadSignature:
        abort(404)
    user = db.session.get(User, uid)
    if user:
        user.email_optout = True
        db.session.commit()
    flash(tr("flash_email_optout"), "success")
    return redirect(url_for("home"))


@app.route("/admin/emails", methods=["GET", "POST"])
@login_required
def admin_emails():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        action = request.form.get("action")
        if action == "test":
            # The test sends the real first follow-up, so the admin sees exactly what teachers get.
            subject, html, body = nudge_email(1, session.get("lang") or current_user.preferred_lang or "fr",
                                              current_user.first_name, f"{APP_URL}/generator",
                                              optout_link(current_user))
            ok = send_email(current_user.email, "[Test] " + subject, body, html)
            flash(tr("flash_email_test_ok" if ok else "flash_email_test_fail").format(email=current_user.email),
                  "success" if ok else "error")
        elif action == "run":
            n = send_activation_nudges(app, db, User, Document, APP_URL, optout_link, _nudge_tracked)
            flash(tr("flash_nudges_sent").format(n=n), "success")
        return redirect(url_for("admin_emails"))
    now = datetime.utcnow()
    base = User.query.filter(User.role == "user")
    stats = []
    for stage, hours in NUDGE_STAGES:
        sent_q = base.filter(User.nudge_stage >= stage)
        sent = sent_q.count()
        activated = sent_q.filter(User.documents.any()).count()
        stats.append({"stage": stage, "hours": hours, "sent": sent, "activated": activated})
    pending = base.filter(User.is_active.is_(True),
                          db.or_(User.email_optout.is_(False), User.email_optout.is_(None)),
                          User.created_at >= now - timedelta(days=NUDGE_MAX_AGE_DAYS),
                          User.created_at <= now - timedelta(hours=NUDGE_STAGES[0][1]),
                          ~User.documents.any(),
                          db.or_(User.nudge_stage.is_(None), User.nudge_stage < len(NUDGE_STAGES))).count()
    missing = [k for k in ("SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASSWORD") if not os.environ.get(k)]
    optouts = base.filter(User.email_optout.is_(True)).count()
    return render_template("admin_emails.html", configured=smtp_configured(), missing=missing,
                           stats=stats, pending=pending, optouts=optouts)


# ── ADMIN: training groups (import an Excel file of participants) ────────────
def _xlsx_response(rows, filename, widths=None):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="102A53")
        c.alignment = Alignment(horizontal="center")
    for i, w in enumerate(widths or [], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/cohorts", methods=["GET", "POST"])
@login_required
def admin_cohorts():
    if current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        name = request.form.get("name", "").strip()[:120]
        pass_type = request.form.get("pass_type", "")
        send_mail = bool(request.form.get("send_email"))
        f = request.files.get("file")
        if not name or pass_type not in PASS_TYPES or not f or not f.filename:
            flash(tr("flash_cohort_fields"), "error")
            return redirect(url_for("admin_cohorts"))
        rows, err = read_participants(f)
        if err:
            flash(tr("flash_cohort_file_" + err), "error")
            return redirect(url_for("admin_cohorts"))
        cohort = Cohort(name=name, pass_type=pass_type)
        db.session.add(cohort)
        db.session.flush()
        report = import_participants(rows, pass_type, cohort, send_mail)
        created = sum(1 for r in report if r["status"] == "created")
        existing = sum(1 for r in report if r["status"] == "existing")
        log_admin_action("cohort_imported", target=name,
                         details=f"{pass_type}: {created} created, {existing} existing, {len(report)} rows")
        db.session.commit()
        session["cohort_report"] = [{k: v for k, v in r.items() if k != "password"} for r in report]
        flash(tr("flash_cohort_imported", created=created, existing=existing), "success")
        return redirect(url_for("admin_cohort_detail", cohort_id=cohort.id))

    cohorts = Cohort.query.order_by(Cohort.created_at.desc()).all()
    stats = {c.id: _cohort_stats(c) for c in cohorts}
    queued = User.query.filter(User.credentials_pending.is_(True)).count()
    return render_template("admin_cohorts.html", cohorts=cohorts, stats=stats, passes=PASS_TYPES,
                           queued=queued, per_day=CREDENTIALS_EMAILS_PER_DAY)


def _cohort_stats(cohort):
    members = cohort.members
    total = members.count()
    logged = members.filter(User.must_change_password.is_(False)).count()
    active = members.filter(User.documents.any()).count()
    paid = members.join(Subscription, Subscription.user_id == User.id) \
        .filter(Subscription.plan.in_(list(PLAN_PRICES))).count()
    return {"total": total, "logged": logged, "active": active, "paid": paid,
            "rate": round(paid * 100 / total) if total else 0}


@app.route("/admin/cohorts/<int:cohort_id>")
@login_required
def admin_cohort_detail(cohort_id):
    if current_user.role != "admin":
        return redirect(url_for("dashboard"))
    cohort = Cohort.query.get_or_404(cohort_id)
    members = cohort.members.order_by(User.last_name, User.first_name).all()
    report = session.pop("cohort_report", None)
    return render_template("admin_cohort_detail.html", cohort=cohort, members=members,
                           stats=_cohort_stats(cohort), report=report, passes=PASS_TYPES)


@app.route("/admin/cohorts/<int:cohort_id>/credentials.xlsx")
@login_required
def admin_cohort_credentials(cohort_id):
    if current_user.role != "admin":
        return redirect(url_for("dashboard"))
    cohort = Cohort.query.get_or_404(cohort_id)
    rows = [[tr("cohort_col_last"), tr("cohort_col_first"), tr("cohort_col_email"),
             tr("cohort_col_password"), tr("cohort_col_link")]]
    for u in cohort.members.order_by(User.last_name, User.first_name):
        pw = u.temp_password if (u.must_change_password and u.temp_password) else tr("cohort_pw_changed")
        rows.append([u.last_name, u.first_name, u.email, pw, f"{APP_URL}/login"])
    log_admin_action("cohort_credentials_downloaded", target=cohort.name)
    db.session.commit()
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", cohort.name)[:40] or "groupe"
    return _xlsx_response(rows, f"identifiants_{safe}.xlsx", [18, 18, 32, 22, 46])


@app.route("/admin/cohorts/template.xlsx")
@login_required
def admin_cohort_template():
    if current_user.role != "admin":
        return redirect(url_for("dashboard"))
    return _xlsx_response([["Nom", "Prénom", "Email"],
                           ["Benali", "Amina", "amina.benali@exemple.dz"],
                           ["Saidi", "Karim", "karim.saidi@exemple.dz"]],
                          "modele_import_formation.xlsx", [20, 20, 34])


# ── TRAINING SESSIONS: public sign-up form, attendance, J-1 reminder, gift pass ──
def _admin_only():
    if not current_user.is_authenticated or current_user.role != "admin":
        flash(tr("flash_admin_only"), "error")
        return redirect(url_for("dashboard"))
    return None


def _parse_date(value):
    try:
        return datetime.strptime((value or "").strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def _unique_slug(title, exclude_id=None):
    base = re.sub(r"[^a-z0-9]+", "-", unicodedata.normalize("NFKD", title).encode("ascii", "ignore").decode().lower()).strip("-")[:50] or "formation"
    slug, n = base, 1
    while True:
        q = TrainingSession.query.filter_by(slug=slug)
        if exclude_id:
            q = q.filter(TrainingSession.id != exclude_id)
        if not q.first():
            return slug
        n += 1
        slug = f"{base}-{n}"


def _session_public_url(s):
    return f"{APP_URL}/formation/{s.slug}"


def _session_when(s):
    first = s.meetings.first()
    d = first.day if first else s.starts_on
    if not d:
        return ""
    return d.strftime("%d/%m/%Y") + (f" {first.time_label}" if first and first.time_label else "")


def _send_enrollment_email(enr, kind):
    """kind: confirmed | pending | waitlist | promoted. Never raises."""
    from notifications import send_email
    from email_templates import enrollment_email
    try:
        s = enr.session
        link = s.online_url if (kind in ("confirmed", "promoted") and s.mode != "onsite") else ""
        subject, html, text = enrollment_email(enr.lang or "fr", enr.first_name, kind, s.title,
                                               _session_when(s), s.location or "", _session_public_url(s), link,
                                               ar={"title": s.loc("title", "ar"), "place": s.loc("location", "ar")})
        send_email(enr.email, subject, text, html)
    except Exception as e:
        print(f"⚠️  Enrollment email failed: {e}")


def promote_waitlist(s):
    """Fill freed seats with the oldest people on the waiting list."""
    promoted = []
    while True:
        left = s.seats_left()
        if left is not None and left <= 0:
            break
        nxt = s.enrollments.filter_by(status="waitlist").order_by(Enrollment.created_at).first()
        if not nxt:
            break
        nxt.status = "confirmed"
        db.session.flush()
        promoted.append(nxt)
    db.session.commit()
    for enr in promoted:
        _send_enrollment_email(enr, "promoted")
    return len(promoted)


def grant_session_gift(enr):
    """Create the account (or extend an existing one) with the session's gift pass.
    Called the first time the person is marked present: the gift rewards showing up."""
    s = enr.session
    if not s.gift_pass or s.gift_pass not in PASS_TYPES or enr.gift_granted:
        return None
    if not s.cohort_id:
        cohort = Cohort(name=f"{s.title}"[:120], pass_type=s.gift_pass)
        db.session.add(cohort)
        db.session.flush()
        s.cohort_id = cohort.id
    cohort = db.session.get(Cohort, s.cohort_id)
    report = import_participants([{"first_name": enr.first_name, "last_name": enr.last_name,
                                   "email": enr.email.lower()}], s.gift_pass, cohort, True)
    status = report[0]["status"] if report else "invalid"
    if status in ("created", "existing", "already_paid"):
        enr.gift_granted = True
        user = User.query.filter_by(email=enr.email.lower()).first()
        if user:
            enr.user_id = user.id
            if not user.preferred_lang:
                user.preferred_lang = enr.lang
    return status


def set_attendance(meeting, enr, status):
    att = Attendance.query.filter_by(meeting_id=meeting.id, enrollment_id=enr.id).first()
    if not att:
        att = Attendance(meeting_id=meeting.id, enrollment_id=enr.id)
        db.session.add(att)
    att.status = status or None
    if status in ("present", "late"):
        grant_session_gift(enr)
    return att


def send_session_reminders():
    """J-1: ask each confirmed participant to confirm (or release the seat)."""
    from notifications import send_email
    from email_templates import session_reminder_email
    with app.app_context():
        tomorrow = (datetime.utcnow() + timedelta(hours=1)).date() + timedelta(days=1)
        meetings = SessionMeeting.query.filter(SessionMeeting.day == tomorrow,
                                               SessionMeeting.reminder_sent.is_(False)).all()
        sent = 0
        for m in meetings:
            s = m.session
            if s.status not in ("open", "running"):
                continue
            when = m.day.strftime("%d/%m/%Y") + (f" {m.time_label}" if m.time_label else "")
            for enr in s.enrollments.filter_by(status="confirmed"):
                base = f"{APP_URL}/formation/rsvp/{enr.token}/{m.id}"
                subject, html, text = session_reminder_email(enr.lang or "fr", enr.first_name, s.title, when,
                                                             s.location or "", base + "?a=yes", base + "?a=no",
                                                             s.online_url if s.mode != "onsite" else "",
                                                             ar={"title": s.loc("title", "ar"), "place": s.loc("location", "ar")})
                if send_email(enr.email, subject, text, html):
                    sent += 1
            m.reminder_sent = True
        db.session.commit()
    if sent:
        print(f"✅ Session reminders sent: {sent}")
    return sent


_SESSION_FIELDS_BOOL = ("is_free", "auto_confirm")


def _apply_session_form(s, form):
    s.title = form.get("title", "").strip()[:160] or s.title
    s.description = form.get("description", "").strip()[:4000] or None
    s.is_free = bool(form.get("is_free"))
    s.auto_confirm = bool(form.get("auto_confirm"))
    s.location = form.get("location", "").strip()[:200] or None
    s.starts_on = _parse_date(form.get("starts_on"))
    s.ends_on = _parse_date(form.get("ends_on"))
    for field, hi in (("price_total", 10_000_000), ("installments", 36), ("hours_total", 5000), ("seats", 5000),
                      ("min_attendance", 100)):
        try:
            v = int(form.get(field) or 0)
        except ValueError:
            v = 0
        setattr(s, field, min(max(v, 0), hi))
    if s.is_free:
        s.price_total, s.installments = 0, 1
    s.installments = max(1, s.installments or 1)
    if not form.get("min_attendance", "").strip():
        s.min_attendance = 80
    gp = form.get("gift_pass", "")
    s.gift_pass = gp if gp in PASS_TYPES else None
    mode = form.get("mode", "onsite")
    s.mode = mode if mode in ("onsite", "online", "hybrid") else "onsite"
    url = form.get("online_url", "").strip()[:300]
    s.online_url = url if url.lower().startswith(("http://", "https://")) and s.mode != "onsite" else None
    for field in ("start_time", "end_time"):
        v = form.get(field, "").strip()
        setattr(s, field, v if re.match(r"^([01]\d|2[0-3]):[0-5]\d$", v) else None)
    s.trainer = form.get("trainer", "").strip()[:120] or None
    s.audience = form.get("audience", "").strip()[:200] or None
    s.programme = form.get("programme", "").strip()[:6000] or None
    s.prerequisites = form.get("prerequisites", "").strip()[:2000] or None
    for field, limit in (("title", 160), ("description", 4000), ("location", 200), ("trainer", 120),
                         ("audience", 200), ("programme", 6000), ("prerequisites", 2000)):
        setattr(s, field + "_ar", form.get(field + "_ar", "").strip()[:limit] or None)


def attendance_stats(s):
    """Attendance per confirmed participant over the meetings already held (date reached or marked).
    Present and late count as attended; excused absences are left out of the denominator."""
    today = (datetime.utcnow() + timedelta(hours=1)).date()
    meetings = s.meetings.all()
    marked_ids = {a.meeting_id for a in Attendance.query.filter(
        Attendance.meeting_id.in_([m.id for m in meetings] or [0]), Attendance.status.isnot(None))}
    held = [m for m in meetings if m.day <= today or m.id in marked_ids]
    held_ids = {m.id for m in held}
    by_enr = {}
    for a in Attendance.query.filter(Attendance.meeting_id.in_(held_ids or {0})):
        by_enr.setdefault(a.enrollment_id, {})[a.meeting_id] = a
    threshold = s.min_attendance if s.min_attendance is not None else 80
    rows = []
    for e in s.enrollments.filter_by(status="confirmed").order_by(Enrollment.last_name, Enrollment.first_name):
        marks = by_enr.get(e.id, {})
        attended = sum(1 for a in marks.values() if a.status in ("present", "late"))
        excused = sum(1 for a in marks.values() if a.status == "excused")
        denom = len(held) - excused
        rate = round(attended * 100 / denom) if denom > 0 else 0
        rows.append({"enr": e, "attended": attended, "excused": excused, "absent": max(0, len(held) - attended - excused),
                     "rate": rate, "eligible": bool(held) and rate >= threshold, "marks": marks})
    avg = round(sum(r["rate"] for r in rows) / len(rows)) if rows else 0
    return {"held": held, "rows": rows, "avg": avg, "threshold": threshold,
            "eligible": sum(1 for r in rows if r["eligible"]), "all_meetings": meetings}


@app.route("/admin/sessions", methods=["GET", "POST"])
@login_required
def admin_sessions():
    denied = _admin_only()
    if denied:
        return denied
    if request.method == "POST":
        title = request.form.get("title", "").strip()[:160]
        if not title:
            flash(tr("flash_session_title"), "error")
            return redirect(url_for("admin_sessions"))
        s = TrainingSession(title=title, slug=_unique_slug(title))
        _apply_session_form(s, request.form)
        db.session.add(s)
        db.session.flush()
        log_admin_action("session_created", target=s.title, details=f"free={s.is_free} seats={s.seats}")
        db.session.commit()
        flash(tr("flash_session_created"), "success")
        return redirect(url_for("admin_session_detail", session_id=s.id))
    sessions = TrainingSession.query.order_by(TrainingSession.created_at.desc()).all()
    counts = {s.id: {"confirmed": s.enrollments.filter_by(status="confirmed").count(),
                     "waitlist": s.enrollments.filter_by(status="waitlist").count(),
                     "pending": s.enrollments.filter_by(status="pending").count()} for s in sessions}
    return render_template("admin_sessions.html", sessions=sessions, counts=counts, passes=PASS_TYPES)


@app.route("/admin/sessions/<int:session_id>")
@login_required
def admin_session_detail(session_id):
    denied = _admin_only()
    if denied:
        return denied
    s = TrainingSession.query.get_or_404(session_id)
    meetings = s.meetings.all()
    enrollments = s.enrollments.all()
    # attendance rate per enrollment (over meetings already held or marked)
    stats = attendance_stats(s)
    by_enr = {r["enr"].id: r for r in stats["rows"]}
    return render_template("admin_session_detail.html", s=s, meetings=meetings, enrollments=enrollments,
                           passes=PASS_TYPES, public_url=_session_public_url(s), stats=stats, by_enr=by_enr,
                           statuses=SESSION_STATUSES)


@app.route("/admin/sessions/<int:session_id>/action", methods=["POST"])
@login_required
def admin_session_action(session_id):
    denied = _admin_only()
    if denied:
        return denied
    s = TrainingSession.query.get_or_404(session_id)
    f = request.form
    action = f.get("action", "")
    back = redirect(url_for("admin_session_detail", session_id=s.id))
    if action == "edit":
        _apply_session_form(s, f)
        db.session.commit()
        flash(tr("flash_session_saved"), "success")
    elif action == "status":
        new = f.get("status", "")
        if new in SESSION_STATUSES:
            s.status = new
            log_admin_action("session_status", target=s.title, details=new)
            db.session.commit()
            flash(tr("flash_session_saved"), "success")
    elif action == "add_meeting":
        d = _parse_date(f.get("day"))
        if d:
            db.session.add(SessionMeeting(session_id=s.id, day=d, time_label=f.get("time_label", "").strip()[:30] or s.default_time_label() or None,
                                          topic=f.get("topic", "").strip()[:200] or None))
            db.session.commit()
        else:
            flash(tr("flash_session_date"), "error")
    elif action == "gen_meetings":
        d = _parse_date(f.get("day"))
        try:
            n = min(max(int(f.get("count") or 0), 1), 60)
            every = min(max(int(f.get("every_days") or 7), 1), 31)
        except ValueError:
            n, every = 0, 7
        if d and n:
            for i in range(n):
                db.session.add(SessionMeeting(session_id=s.id, day=d + timedelta(days=every * i),
                                              time_label=f.get("time_label", "").strip()[:30] or s.default_time_label() or None))
            db.session.commit()
            flash(tr("flash_meetings_generated", n=n), "success")
        else:
            flash(tr("flash_session_date"), "error")
    elif action == "del_meeting":
        m = SessionMeeting.query.filter_by(id=f.get("meeting_id", type=int), session_id=s.id).first()
        if m:
            db.session.delete(m)
            db.session.commit()
    elif action in ("enr_confirm", "enr_cancel", "enr_delete"):
        enr = Enrollment.query.filter_by(id=f.get("enrollment_id", type=int), session_id=s.id).first()
        if enr:
            if action == "enr_confirm":
                was = enr.status
                enr.status = "confirmed"
                db.session.commit()
                if was != "confirmed":
                    _send_enrollment_email(enr, "confirmed")
            elif action == "enr_cancel":
                enr.status = "cancelled"
                db.session.commit()
                promote_waitlist(s)
            else:
                db.session.delete(enr)
                db.session.commit()
                promote_waitlist(s)
    elif action == "enr_add":
        email = f.get("email", "").strip().lower()
        if _EMAIL_RE.match(email) and f.get("last_name", "").strip() and not s.enrollments.filter_by(email=email).first():
            enr = Enrollment(session_id=s.id, first_name=f.get("first_name", "").strip()[:80] or email.split("@")[0],
                             last_name=f.get("last_name", "").strip()[:80], email=email,
                             phone=f.get("phone", "").strip()[:30] or None, lang="fr", status="confirmed")
            db.session.add(enr)
            db.session.commit()
            _send_enrollment_email(enr, "confirmed")
        else:
            flash(tr("flash_session_enr_invalid"), "error")
    return back


@app.route("/admin/sessions/<int:session_id>/meetings/<int:meeting_id>", methods=["GET", "POST"])
@login_required
def admin_session_attendance(session_id, meeting_id):
    denied = _admin_only()
    if denied:
        return denied
    s = TrainingSession.query.get_or_404(session_id)
    m = SessionMeeting.query.filter_by(id=meeting_id, session_id=s.id).first_or_404()
    enrollments = s.enrollments.filter_by(status="confirmed").order_by(Enrollment.last_name, Enrollment.first_name).all()
    if request.method == "POST":
        gifts = 0
        for enr in enrollments:
            st = request.form.get(f"st_{enr.id}", "")
            if st in ("present", "absent", "late", "excused", ""):
                had_gift = enr.gift_granted
                set_attendance(m, enr, st)
                gifts += 1 if (enr.gift_granted and not had_gift) else 0
        log_admin_action("attendance_saved", target=s.title, details=f"{m.day} gifts={gifts}")
        db.session.commit()
        flash(tr("flash_attendance_saved", n=gifts), "success")
        return redirect(url_for("admin_session_attendance", session_id=s.id, meeting_id=m.id))
    att = {a.enrollment_id: a for a in m.attendances}
    return render_template("admin_attendance.html", s=s, m=m, enrollments=enrollments, att=att)


@app.route("/admin/sessions/<int:session_id>/enrollments.xlsx")
@login_required
def admin_session_export(session_id):
    denied = _admin_only()
    if denied:
        return denied
    s = TrainingSession.query.get_or_404(session_id)
    rows = [[tr("cohort_col_last"), tr("cohort_col_first"), tr("cohort_col_email"), tr("enr_phone"),
             tr("enr_wilaya"), tr("enr_school"), tr("enr_subject"), tr("enr_status")]]
    for e in s.enrollments:
        rows.append([e.last_name, e.first_name, e.email, e.phone or "", e.wilaya or "", e.school or "",
                     e.subject or "", tr("enr_status_" + e.status)])
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", s.slug)[:40] or "formation"
    return _xlsx_response(rows, f"inscrits_{safe}.xlsx", [18, 18, 32, 16, 18, 28, 20, 14])


@app.route("/admin/sessions/<int:session_id>/attendance.xlsx")
@login_required
def admin_session_attendance_export(session_id):
    denied = _admin_only()
    if denied:
        return denied
    s = TrainingSession.query.get_or_404(session_id)
    st = attendance_stats(s)
    sym = {"present": tr("att_present"), "late": tr("att_late"), "excused": tr("att_excused"), "absent": tr("att_absent")}
    head = [tr("cohort_col_last"), tr("cohort_col_first"), tr("cohort_col_email")] + \
           [m.day.strftime("%d/%m") for m in st["held"]] + [tr("sess_rate") + " %", tr("sess_eligible")]
    rows = [head]
    for r in st["rows"]:
        e = r["enr"]
        cells = [sym.get(r["marks"][m.id].status, "") if m.id in r["marks"] and r["marks"][m.id].status else ""
                 for m in st["held"]]
        rows.append([e.last_name, e.first_name, e.email] + cells + [r["rate"], tr("sess_yes") if r["eligible"] else tr("sess_no")])
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", s.slug)[:40] or "formation"
    return _xlsx_response(rows, f"presences_{safe}.xlsx", [18, 18, 30] + [11] * len(st["held"]) + [10, 12])


@app.route("/formation/<slug>", methods=["GET", "POST"])
@limiter.limit("30/hour", methods=["POST"])
def formation_public(slug):
    s = TrainingSession.query.filter_by(slug=slug).first_or_404()
    if s.status == "draft":
        abort(404)
    is_open = s.status == "open"
    done = None
    if request.method == "POST" and is_open:
        f = request.form
        if f.get("website"):  # honeypot
            return redirect(url_for("formation_public", slug=slug))
        first, last = f.get("first_name", "").strip()[:80], f.get("last_name", "").strip()[:80]
        email, phone = f.get("email", "").strip().lower(), f.get("phone", "").strip()[:30]
        if not first or not last or not _EMAIL_RE.match(email) or len(re.sub(r"\D", "", phone)) < 8:
            flash(tr("enr_invalid"), "error")
            return render_template("formation_public.html", s=s, is_open=is_open, done=None, form=f)
        existing = s.enrollments.filter_by(email=email).first()
        if existing and existing.status != "cancelled":
            return render_template("formation_public.html", s=s, is_open=is_open, done="duplicate", form={})
        left = s.seats_left()
        if left is not None and left <= 0:
            status = "waitlist"
        else:
            status = "confirmed" if s.auto_confirm else "pending"
        if existing:  # re-registration after a cancellation
            enr = existing
            enr.status = status
        else:
            enr = Enrollment(session_id=s.id, email=email)
            db.session.add(enr)
        enr.status = status
        enr.first_name, enr.last_name, enr.phone = first, last, phone
        enr.wilaya = f.get("wilaya", "").strip()[:60] or None
        enr.school = f.get("school", "").strip()[:150] or None
        enr.subject = f.get("subject", "").strip()[:100] or None
        enr.lang = current_lang()
        db.session.commit()
        _send_enrollment_email(enr, status)
        return render_template("formation_public.html", s=s, is_open=is_open, done=status, form={})
    meetings = s.meetings.all()
    return render_template("formation_public.html", s=s, is_open=is_open, done=done, form={}, meetings=meetings)


@app.route("/formation/rsvp/<token>/<int:meeting_id>", methods=["GET", "POST"])
def formation_rsvp(token, meeting_id):
    enr = Enrollment.query.filter_by(token=token).first_or_404()
    m = SessionMeeting.query.filter_by(id=meeting_id, session_id=enr.session_id).first_or_404()
    s = enr.session
    answer = (request.values.get("a") or "").lower()
    if answer not in ("yes", "no"):
        abort(404)
    result = None
    if request.method == "POST":
        att = Attendance.query.filter_by(meeting_id=m.id, enrollment_id=enr.id).first()
        if not att:
            att = Attendance(meeting_id=m.id, enrollment_id=enr.id)
            db.session.add(att)
        att.rsvp = answer
        released = False
        if answer == "no":
            first_meeting = s.meetings.first()
            attended_before = Attendance.query.filter(Attendance.enrollment_id == enr.id,
                                                      Attendance.status.in_(("present", "late"))).count() > 0
            if first_meeting and first_meeting.id == m.id and not attended_before and enr.status == "confirmed":
                enr.status = "cancelled"  # frees the seat for the waiting list
                released = True
        db.session.commit()
        if released:
            promote_waitlist(s)
        result = "released" if released else answer
    return render_template("formation_rsvp.html", s=s, m=m, enr=enr, answer=answer, result=result)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
