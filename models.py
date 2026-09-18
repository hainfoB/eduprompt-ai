from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, date
import secrets

db = SQLAlchemy()

TRIAL_DAYS  = 14
TRIAL_DOCS  = 5

# Daily lesson quotas per paid plan (used instead of a total document cap)
PLAN_DAILY_LIMITS = {
    "pro": 2,
    "premium": 5,
}

# Reference prices (DA / year) — used to log payments automatically
PLAN_PRICES = {
    "pro": 1500,
    "premium": 1800,
}


class User(UserMixin, db.Model):
    id            = db.Column(db.Integer, primary_key=True)
    first_name    = db.Column(db.String(80), nullable=False)
    last_name     = db.Column(db.String(80), nullable=False)
    email         = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role          = db.Column(db.String(20), default="user")  # user | admin
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)
    gemini_api_key = db.Column(db.String(255), nullable=True)
    preferred_lang = db.Column(db.String(5), default="fr")
    is_active      = db.Column(db.Boolean, default=True)  # soft-delete: False = deactivated account

    subscription  = db.relationship("Subscription", backref="user", uselist=False,
                                     cascade="all, delete-orphan")
    documents     = db.relationship("Document", backref="user",
                                     cascade="all, delete-orphan",
                                     order_by="Document.created_at.desc()")

    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)

    @property
    def full_name(self):
        return f"{self.first_name} {self.last_name}"


class Subscription(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), unique=True, nullable=False)
    plan       = db.Column(db.String(20), default="trial")     # trial | pro | premium
    status     = db.Column(db.String(20), default="active")    # active | expired | cancelled
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    expires_at = db.Column(db.DateTime)
    docs_used  = db.Column(db.Integer, default=0)               # trial: total docs used
    docs_limit = db.Column(db.Integer, default=TRIAL_DOCS)      # trial: total docs cap

    # Daily quota tracking (used for pro / premium plans)
    docs_used_today  = db.Column(db.Integer, default=0)
    last_reset_date  = db.Column(db.Date, nullable=True)

    # Expiry reminder (5-day heads-up email), reset whenever the plan is renewed
    expiry_reminder_sent = db.Column(db.Boolean, default=False)

    def is_expired(self):
        return bool(self.expires_at and datetime.utcnow() > self.expires_at)

    def is_daily_plan(self):
        return self.plan in PLAN_DAILY_LIMITS

    def daily_limit(self):
        return PLAN_DAILY_LIMITS.get(self.plan)

    def _roll_daily_counter_if_needed(self):
        today = date.today()
        if self.last_reset_date != today:
            self.docs_used_today = 0
            self.last_reset_date = today

    def has_quota(self):
        if self.is_expired():
            return False
        if self.is_daily_plan():
            self._roll_daily_counter_if_needed()
            return self.docs_used_today < self.daily_limit()
        # trial (or any other non-daily plan): total cap
        return self.docs_used < self.docs_limit

    def remaining(self):
        if self.is_daily_plan():
            self._roll_daily_counter_if_needed()
            return max(0, self.daily_limit() - self.docs_used_today)
        return max(0, self.docs_limit - self.docs_used)

    def register_usage(self):
        """Call after a successful generation to decrement the right counter."""
        if self.is_daily_plan():
            self._roll_daily_counter_if_needed()
            self.docs_used_today += 1
        else:
            self.docs_used += 1

    def days_until_expiry(self):
        if not self.expires_at:
            return None
        delta = self.expires_at - datetime.utcnow()
        return delta.days

    def plan_label(self):
        """Localised plan name for the currently active session language."""
        try:
            from flask import session
            lang = session.get("lang", "fr")
        except Exception:
            lang = "fr"
        labels = {
            "fr": {"trial": "Essai gratuit", "pro": "Pro", "premium": "Premium"},
            "en": {"trial": "Free trial", "pro": "Pro", "premium": "Premium"},
            "ar": {"trial": "تجربة مجانية", "pro": "برو", "premium": "بريميوم"},
        }
        return labels.get(lang, labels["fr"]).get(self.plan, self.plan)


class Payment(db.Model):
    """Manual payment log (bank transfer / CCP) — used for revenue tracking."""
    id         = db.Column(db.Integer, primary_key=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    plan       = db.Column(db.String(20))
    amount     = db.Column(db.Integer)          # in DA
    currency   = db.Column(db.String(10), default="DA")
    note       = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship("User")


class PaymentRequest(db.Model):
    """A teacher's self-declared payment (bank transfer / CCP / WhatsApp), awaiting
    admin validation before the corresponding plan is activated or renewed."""
    id             = db.Column(db.Integer, primary_key=True)
    user_id        = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    plan           = db.Column(db.String(20), nullable=False)   # pro | premium
    amount_claimed = db.Column(db.Integer, nullable=True)
    reference      = db.Column(db.String(120), nullable=True)   # transfer/CCP/WhatsApp reference
    note           = db.Column(db.Text, nullable=True)
    status         = db.Column(db.String(20), default="pending")  # pending | approved | rejected
    created_at     = db.Column(db.DateTime, default=datetime.utcnow)
    reviewed_at    = db.Column(db.DateTime, nullable=True)
    reviewed_by    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    admin_note     = db.Column(db.Text, nullable=True)

    user     = db.relationship("User", foreign_keys=[user_id])
    reviewer = db.relationship("User", foreign_keys=[reviewed_by])

    def status_label(self):
        labels = {
            "fr": {"pending": "En attente", "approved": "Validé", "rejected": "Refusé"},
            "en": {"pending": "Pending", "approved": "Approved", "rejected": "Rejected"},
            "ar": {"pending": "قيد الانتظار", "approved": "تم القبول", "rejected": "مرفوض"},
        }
        try:
            from flask import session
            lang = session.get("lang", "fr")
        except Exception:
            lang = "fr"
        return labels.get(lang, labels["fr"]).get(self.status, self.status)


class AdminLog(db.Model):
    """Lightweight audit trail of admin actions, for accountability."""
    id         = db.Column(db.Integer, primary_key=True)
    admin_id   = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    action     = db.Column(db.String(80), nullable=False)
    target     = db.Column(db.String(255), nullable=True)   # human-readable target description
    details    = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    admin = db.relationship("User")


class Message(db.Model):
    """Free, unlimited direct-message thread between a teacher and the admin.
    One thread per teacher (user_id); `sender` says who wrote this particular message."""
    id              = db.Column(db.Integer, primary_key=True)
    user_id         = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    sender          = db.Column(db.String(10), default="teacher")   # teacher | admin
    body            = db.Column(db.Text, nullable=False)
    created_at      = db.Column(db.DateTime, default=datetime.utcnow)
    read_by_admin   = db.Column(db.Boolean, default=False)
    read_by_teacher = db.Column(db.Boolean, default=False)

    user = db.relationship("User")


class Document(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    title      = db.Column(db.String(255))
    subject    = db.Column(db.String(100))
    level      = db.Column(db.String(100))
    palier     = db.Column(db.String(100))
    doc_type   = db.Column(db.String(50))
    fmt        = db.Column(db.String(10))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class LicenseCode(db.Model):
    """Manually-issued activation codes (sold via bank transfer / WhatsApp / etc.)."""
    id            = db.Column(db.Integer, primary_key=True)
    code          = db.Column(db.String(40), unique=True,
                               default=lambda: secrets.token_hex(6).upper())
    plan          = db.Column(db.String(20), default="pro")
    duration_days = db.Column(db.Integer, default=365)
    used          = db.Column(db.Boolean, default=False)
    used_by       = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    used_at       = db.Column(db.DateTime, nullable=True)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)

    used_by_user  = db.relationship("User", foreign_keys=[used_by])


def create_trial_subscription(user):
    sub = Subscription(
        user_id=user.id,
        plan="trial",
        status="active",
        started_at=datetime.utcnow(),
        expires_at=datetime.utcnow() + timedelta(days=TRIAL_DAYS),
        docs_used=0,
        docs_limit=TRIAL_DOCS,
    )
    db.session.add(sub)
    return sub
