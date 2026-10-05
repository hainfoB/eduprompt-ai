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
    "ultimate": 5,
}

# Reference prices (DA / year) — used to log payments automatically
PLAN_PRICES = {
    "pro": 1500,
    "ultimate": 1800,
}

PAYMENT_METHODS = ("cash", "ccp")

# Training passes, granted by the admin to the participants of a training
# session (imported from an Excel file). "monthly" passes reset their document
# counter every 30 days; at the end, a discount on the annual plan is offered.
PASS_TYPES = {
    "pass_decouverte": {"days": 30, "docs": 15, "monthly": False, "discount": 30, "discount_days": 15},
    "pass_plus":       {"days": 90, "docs": 10, "monthly": True,  "discount": 20, "discount_days": 30},
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
    profile = db.Column(db.String(10), default="teacher")  # "teacher" | "parent"
    signup_source = db.Column(db.String(40), nullable=True)  # facebook, google, direct…
    # Referral program: personal invite code, who invited this user, and whether
    # the inviter has already been rewarded for this user's first document.
    referral_code = db.Column(db.String(12), unique=True, nullable=True, index=True)
    referred_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    referral_rewarded = db.Column(db.Boolean, default=False)
    referral_docs_earned = db.Column(db.Integer, default=0)  # lifetime, as inviter (capped)
    # Training passes: the group the account was imported with, a forced password
    # change at first login, the temporary password kept only until the user
    # changes it (so the admin can print it and the email queue can send it),
    # and the end-of-pass discount on the annual plan.
    cohort_id = db.Column(db.Integer, db.ForeignKey("cohort.id"), nullable=True, index=True)
    must_change_password = db.Column(db.Boolean, default=False)
    temp_password = db.Column(db.String(40), nullable=True)
    credentials_pending = db.Column(db.Boolean, default=False)
    credentials_sent_at = db.Column(db.DateTime, nullable=True)
    discount_pct = db.Column(db.Integer, nullable=True)
    discount_until = db.Column(db.DateTime, nullable=True)
    pass_reminder_sent = db.Column(db.Boolean, default=False)

    def active_discount(self):
        """Percentage off the annual plans still valid for this user, else 0."""
        if self.discount_pct and self.discount_until and datetime.utcnow() < self.discount_until:
            return self.discount_pct
        return 0
    # Activation follow-up for sign-ups who never generated a document:
    # 0 = nothing sent, 1 = 24 h email sent, 2 = 72 h email sent.
    nudge_stage = db.Column(db.Integer, default=0)
    email_optout = db.Column(db.Boolean, default=False)  # unsubscribed from follow-up emails

    @property
    def is_parent(self):
        return self.profile == "parent"
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
    plan       = db.Column(db.String(20), default="trial")     # trial | pro | ultimate
    status     = db.Column(db.String(20), default="active")    # active | expired | cancelled
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    expires_at = db.Column(db.DateTime)
    docs_used  = db.Column(db.Integer, default=0)               # trial: total docs used
    docs_limit = db.Column(db.Integer, default=TRIAL_DOCS)      # trial: total docs cap

    # Extra documents earned (referral program). Used only once the normal
    # quota is exhausted. They are valid for one month; the validity can be
    # extended once (by a new reward earned while they are still valid).
    bonus_docs = db.Column(db.Integer, default=0)
    bonus_expires_at = db.Column(db.DateTime, nullable=True)
    bonus_renewed = db.Column(db.Boolean, default=False)

    # Daily quota tracking (used for pro / premium plans)
    docs_used_today  = db.Column(db.Integer, default=0)
    last_reset_date  = db.Column(db.Date, nullable=True)

    # Expiry reminder (5-day heads-up email), reset whenever the plan is renewed
    expiry_reminder_sent = db.Column(db.Boolean, default=False)
    # Multi-stage reminder tracking: last threshold (7, 3 or 0 days) already notified for.
    # Reset to None whenever the plan is renewed/changed so reminders fire again next cycle.
    last_reminder_stage = db.Column(db.Integer, nullable=True)

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

    def _roll_monthly_if_needed(self):
        """Monthly passes: the document counter starts again every 30 days."""
        cfg = PASS_TYPES.get(self.plan)
        if not cfg or not cfg["monthly"]:
            return
        today = date.today()
        if self.last_reset_date is None:
            self.last_reset_date = today
        elif (today - self.last_reset_date).days >= 30:
            self.docs_used = 0
            self.last_reset_date = today

    def is_pass(self):
        return self.plan in PASS_TYPES

    def _plan_has_quota(self):
        self._roll_monthly_if_needed()
        if self.is_expired():
            return False
        if self.is_daily_plan():
            self._roll_daily_counter_if_needed()
            return self.docs_used_today < self.daily_limit()
        # trial (or any other non-daily plan): total cap
        return (self.docs_used or 0) < (self.docs_limit or 0)

    def valid_bonus(self):
        """Bonus documents still usable (0 once their validity has passed)."""
        if not self.bonus_docs or not self.bonus_expires_at:
            return 0
        return self.bonus_docs if datetime.utcnow() < self.bonus_expires_at else 0

    def add_bonus(self, n, days=30):
        """Credit n bonus documents. A fresh grant is valid `days`; a grant made
        while bonuses are still valid extends their validity once, then never again."""
        now = datetime.utcnow()
        if not self.valid_bonus():
            self.bonus_docs = n                 # expired leftovers are dropped
            self.bonus_expires_at = now + timedelta(days=days)
            self.bonus_renewed = False
        else:
            self.bonus_docs += n
            if not self.bonus_renewed:
                self.bonus_expires_at = now + timedelta(days=days)
                self.bonus_renewed = True

    def has_quota(self):
        return self._plan_has_quota() or self.valid_bonus() > 0

    def plan_remaining(self):
        self._roll_monthly_if_needed()
        if self.is_expired():
            return 0
        if self.is_daily_plan():
            self._roll_daily_counter_if_needed()
            return max(0, self.daily_limit() - self.docs_used_today)
        return max(0, (self.docs_limit or 0) - (self.docs_used or 0))

    def remaining(self):
        return self.plan_remaining() + self.valid_bonus()

    def register_usage(self):
        """Call after a successful generation: plan quota first, then bonus documents."""
        if self._plan_has_quota():
            if self.is_daily_plan():
                self.docs_used_today += 1
            else:
                self.docs_used = (self.docs_used or 0) + 1
        elif self.valid_bonus() > 0:
            self.bonus_docs -= 1

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
            "fr": {"trial": "Essai gratuit", "pro": "Pro", "ultimate": "Ultimate",
                   "pass_decouverte": "Pass Découverte", "pass_plus": "Pass Formation+"},
            "en": {"trial": "Free trial", "pro": "Pro", "ultimate": "Ultimate",
                   "pass_decouverte": "Discovery Pass", "pass_plus": "Training Pass+"},
            "ar": {"trial": "تجربة مجانية", "pro": "برو", "ultimate": "ألتيميت",
                   "pass_decouverte": "باس الاكتشاف", "pass_plus": "باس التكوين+"},
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
    """A teacher's self-declared payment (cash / CCP transfer), awaiting
    admin validation before the corresponding plan is activated or renewed."""
    id             = db.Column(db.Integer, primary_key=True)
    user_id        = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    plan           = db.Column(db.String(20), nullable=False)   # pro | ultimate
    amount_claimed = db.Column(db.Integer, nullable=True)
    method         = db.Column(db.String(10), nullable=True)    # cash | ccp
    reference      = db.Column(db.String(120), nullable=True)   # CCP transfer reference
    receipt_path   = db.Column(db.String(255), nullable=True)   # uploaded screenshot/photo of the receipt
    note           = db.Column(db.Text, nullable=True)
    status         = db.Column(db.String(20), default="pending")  # pending | approved | rejected
    created_at     = db.Column(db.DateTime, default=datetime.utcnow)
    reviewed_at    = db.Column(db.DateTime, nullable=True)
    reviewed_by    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    admin_note     = db.Column(db.Text, nullable=True)
    receipt_number = db.Column(db.String(30), nullable=True)    # assigned on approval, e.g. REC-2026-00042
    source         = db.Column(db.String(20), default="upgrade")  # upgrade | registration

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

    def method_label(self):
        labels = {
            "fr": {"cash": "Espèces", "ccp": "Virement CCP"},
            "en": {"cash": "Cash", "ccp": "CCP transfer"},
            "ar": {"cash": "نقدًا", "ccp": "تحويل CCP"},
        }
        try:
            from flask import session
            lang = session.get("lang", "fr")
        except Exception:
            lang = "fr"
        return labels.get(lang, labels["fr"]).get(self.method, self.method or "—")

    def is_pending_overdue(self, hours=48):
        if self.status != "pending":
            return False
        return (datetime.utcnow() - self.created_at) > timedelta(hours=hours)


class PlanChangeHistory(db.Model):
    """Audit trail of every plan transition for a teacher (upgrade, renewal,
    manual admin change, or automatic downgrade at expiry)."""
    id         = db.Column(db.Integer, primary_key=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    from_plan  = db.Column(db.String(20), nullable=True)
    to_plan    = db.Column(db.String(20), nullable=False)
    reason     = db.Column(db.String(40), nullable=False)  # payment_approved | admin_manual | auto_downgrade | license_code
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship("User")


class Cohort(db.Model):
    """A group of participants imported together after a training session."""
    id         = db.Column(db.Integer, primary_key=True)
    name       = db.Column(db.String(120), nullable=False)
    pass_type  = db.Column(db.String(20), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    members    = db.relationship("User", backref="cohort", lazy="dynamic",
                                 foreign_keys="User.cohort_id")


class FunnelEvent(db.Model):
    """One step of the acquisition funnel, used to measure campaigns.
    event: landing_view | demo | signup | first_document | payment_declared | payment_approved
    visitor: random per-browser id (no personal data), to count unique visitors."""
    id         = db.Column(db.Integer, primary_key=True)
    event      = db.Column(db.String(30), nullable=False, index=True)
    profile    = db.Column(db.String(10), nullable=True)
    source     = db.Column(db.String(40), nullable=True, index=True)
    visitor    = db.Column(db.String(32), nullable=True, index=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)


class GenerationError(db.Model):
    """A document generation that failed in front of the user (or a failed demo),
    so the admin can see how often it happens and why."""
    id         = db.Column(db.Integer, primary_key=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    stage      = db.Column(db.String(20), default="generate")  # generate | demo | download | upstream
    message    = db.Column(db.String(500), nullable=True)
    doc_type   = db.Column(db.String(40), nullable=True)
    subject    = db.Column(db.String(80), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)

    user = db.relationship("User")


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
    # Saved so the user can download it again later (history), without quota:
    content    = db.Column(db.Text, nullable=True)   # generated text, with [[VISUAL_ID:n]] tags
    payload    = db.Column(db.Text, nullable=True)   # JSON: lang, header meta, colors, visuals
    # User feedback on quality
    rating     = db.Column(db.Integer, nullable=True)  # 1 = useful, -1 = not useful
    rating_comment = db.Column(db.String(500), nullable=True)
    rated_at   = db.Column(db.DateTime, nullable=True)


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


def next_receipt_number():
    """Sequential, year-scoped receipt number: REC-2026-00001, REC-2026-00002, ..."""
    year = datetime.utcnow().year
    prefix = f"REC-{year}-"
    last = (PaymentRequest.query
            .filter(PaymentRequest.receipt_number.like(f"{prefix}%"))
            .order_by(PaymentRequest.receipt_number.desc())
            .first())
    if last and last.receipt_number:
        try:
            n = int(last.receipt_number.rsplit("-", 1)[-1]) + 1
        except ValueError:
            n = 1
    else:
        n = 1
    return f"{prefix}{n:05d}"


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

# ── TRAINING SESSIONS (registration form, attendance, later payments/certificates) ──
SESSION_STATUSES = ("draft", "open", "running", "done")
ENROLL_STATUSES = ("pending", "confirmed", "waitlist", "cancelled")


class TrainingSession(db.Model):
    """A training session the admin runs (free or paid), with its public sign-up form."""
    id            = db.Column(db.Integer, primary_key=True)
    title         = db.Column(db.String(160), nullable=False)
    slug          = db.Column(db.String(80), unique=True, nullable=False, index=True)
    description   = db.Column(db.Text, nullable=True)
    is_free       = db.Column(db.Boolean, default=True)
    price_total   = db.Column(db.Integer, default=0)      # DA, paid sessions
    installments  = db.Column(db.Integer, default=1)      # number of payments (patch 3)
    starts_on     = db.Column(db.Date, nullable=True)
    ends_on       = db.Column(db.Date, nullable=True)
    location      = db.Column(db.String(200), nullable=True)
    hours_total   = db.Column(db.Integer, default=0)
    seats         = db.Column(db.Integer, default=0)      # 0 = unlimited
    min_attendance = db.Column(db.Integer, default=80)    # % of held meetings needed for the certificate
    auto_confirm  = db.Column(db.Boolean, default=True)
    gift_pass     = db.Column(db.String(20), nullable=True)   # PASS_TYPES key, given after 1st attended meeting
    status        = db.Column(db.String(10), default="draft")
    cohort_id     = db.Column(db.Integer, db.ForeignKey("cohort.id"), nullable=True)
    mode          = db.Column(db.String(10), default="onsite")   # onsite | online | hybrid
    online_url    = db.Column(db.String(300), nullable=True)     # sent to confirmed people only
    group_url     = db.Column(db.String(300), nullable=True)     # WhatsApp / Telegram group, sent to confirmed people only
    start_time    = db.Column(db.String(5), nullable=True)       # "HH:MM", default for every meeting
    end_time      = db.Column(db.String(5), nullable=True)
    trainer       = db.Column(db.String(120), nullable=True)
    audience      = db.Column(db.String(200), nullable=True)
    programme     = db.Column(db.Text, nullable=True)            # one line per point
    prerequisites = db.Column(db.Text, nullable=True)
    # Arabic version of the texts (shown to Arabic visitors; the French ones are the default)
    title_ar        = db.Column(db.String(160), nullable=True)
    description_ar  = db.Column(db.Text, nullable=True)
    location_ar     = db.Column(db.String(200), nullable=True)
    trainer_ar      = db.Column(db.String(120), nullable=True)
    audience_ar     = db.Column(db.String(200), nullable=True)
    programme_ar    = db.Column(db.Text, nullable=True)
    prerequisites_ar= db.Column(db.Text, nullable=True)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)
    meetings      = db.relationship("SessionMeeting", backref="session", lazy="dynamic",
                                    order_by="SessionMeeting.day", cascade="all, delete-orphan")
    enrollments   = db.relationship("Enrollment", backref="session", lazy="dynamic",
                                    order_by="Enrollment.created_at", cascade="all, delete-orphan")

    def taken(self):
        """Seats occupied (confirmed + waiting for admin validation)."""
        return self.enrollments.filter(Enrollment.status.in_(("confirmed", "pending"))).count()

    def seats_left(self):
        return None if not self.seats else max(0, self.seats - self.taken())

    def default_time_label(self):
        if self.start_time and self.end_time:
            return f"{self.start_time} - {self.end_time}"
        return self.start_time or ""

    def scheduled_minutes(self):
        """Total length of the planned meetings (a meeting without a readable time uses the default one)."""
        return sum(mt.duration_minutes(self) for mt in self.meetings)

    def loc(self, field, lang):
        """Text of `field` in `lang` (Arabic if provided, else French), falling back to the other one."""
        base, ar = getattr(self, field) or "", getattr(self, field + "_ar") or ""
        return (ar or base) if lang == "ar" else (base or ar)

    @staticmethod
    def _items(text):
        return [l.strip(" -•\t") for l in (text or "").splitlines() if l.strip(" -•\t")]

    def programme_items(self, lang="fr"):
        return self._items(self.loc("programme", lang))

    def prerequisite_items(self, lang="fr"):
        return self._items(self.loc("prerequisites", lang))


class SessionMeeting(db.Model):
    """One séance of a session (one date, one time)."""
    id         = db.Column(db.Integer, primary_key=True)
    session_id = db.Column(db.Integer, db.ForeignKey("training_session.id"), nullable=False, index=True)
    day        = db.Column(db.Date, nullable=False)
    time_label = db.Column(db.String(30), nullable=True)   # e.g. "14:00 - 16:00"
    topic      = db.Column(db.String(200), nullable=True)
    reminder_sent = db.Column(db.Boolean, default=False)

    def duration_minutes(self, session=None):
        """Minutes between the two times of "HH:MM - HH:MM" (0 if unreadable)."""
        import re as _re
        for label in (self.time_label, session.default_time_label() if session else ""):
            mm = _re.match(r"^\s*(\d{1,2}):(\d{2})\s*[-–]\s*(\d{1,2}):(\d{2})\s*$", label or "")
            if mm:
                a, b = int(mm[1]) * 60 + int(mm[2]), int(mm[3]) * 60 + int(mm[4])
                if b > a:
                    return b - a
        return 0

    attendances = db.relationship("Attendance", backref="meeting", lazy="dynamic",
                                  cascade="all, delete-orphan")


class Enrollment(db.Model):
    """A person registered to a session through the public form (or added by the admin)."""
    id         = db.Column(db.Integer, primary_key=True)
    session_id = db.Column(db.Integer, db.ForeignKey("training_session.id"), nullable=False, index=True)
    first_name = db.Column(db.String(80), nullable=False)
    last_name  = db.Column(db.String(80), nullable=False)
    email      = db.Column(db.String(150), nullable=False, index=True)
    phone      = db.Column(db.String(30), nullable=True)
    wilaya     = db.Column(db.String(60), nullable=True)
    school     = db.Column(db.String(150), nullable=True)
    subject    = db.Column(db.String(100), nullable=True)
    lang       = db.Column(db.String(2), default="fr")
    status     = db.Column(db.String(10), default="pending", index=True)
    token      = db.Column(db.String(40), unique=True, nullable=False, default=lambda: secrets.token_urlsafe(24))
    gift_granted = db.Column(db.Boolean, default=False)
    source     = db.Column(db.String(40), nullable=True, index=True)   # ?src=… of the link they came from
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    attendances = db.relationship("Attendance", backref="enrollment", lazy="dynamic",
                                  cascade="all, delete-orphan")

    @property
    def full_name(self):
        return f"{self.first_name} {self.last_name}".strip()


class Attendance(db.Model):
    """Attendance of one enrollment at one meeting. status: present | absent | late | excused.
    rsvp is the answer to the J-1 reminder (yes | no)."""
    id            = db.Column(db.Integer, primary_key=True)
    meeting_id    = db.Column(db.Integer, db.ForeignKey("session_meeting.id"), nullable=False, index=True)
    enrollment_id = db.Column(db.Integer, db.ForeignKey("enrollment.id"), nullable=False, index=True)
    status        = db.Column(db.String(10), nullable=True)
    rsvp          = db.Column(db.String(3), nullable=True)
    __table_args__ = (db.UniqueConstraint("meeting_id", "enrollment_id", name="uq_attendance"),)


class SessionVisit(db.Model):
    """One visitor (per browser) opening a session's public page, to measure which post brings people."""
    id         = db.Column(db.Integer, primary_key=True)
    session_id = db.Column(db.Integer, db.ForeignKey("training_session.id"), nullable=False, index=True)
    source     = db.Column(db.String(40), nullable=False, default="direct", index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
