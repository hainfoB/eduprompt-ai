import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta

from translations import get_translations

REMINDER_STAGES = (7, 3, 0)  # days-before-expiry checkpoints, in order


def send_email(to_email, subject, body):
    """Send a plain-text email via SMTP. Returns True on success.
    Requires SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD env vars.
    Degrades gracefully (logs + returns False) if not configured, so the
    app never crashes for lack of email credentials."""
    host = os.environ.get("SMTP_HOST")
    port = os.environ.get("SMTP_PORT")
    user = os.environ.get("SMTP_USER")
    password = os.environ.get("SMTP_PASSWORD")
    sender = os.environ.get("SMTP_FROM", user)

    if not all([host, port, user, password]):
        print(f"⚠️  SMTP not configured — skipping email to {to_email}: {subject}")
        return False

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = to_email

        port = int(port)
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=15) as server:
                server.login(user, password)
                server.sendmail(sender, [to_email], msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=15) as server:
                server.starttls()
                server.login(user, password)
                server.sendmail(sender, [to_email], msg.as_string())
        print(f"✅ Reminder email sent to {to_email}")
        return True
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")
        return False


def check_and_send_expiry_reminders(app, db, User, Subscription, contact_email):
    """Multi-stage reminders at J-7 / J-3 / J-0 before a paid subscription expires
    (email + an in-app message so the notification bell also lights up), plus an
    automatic downgrade to a fresh trial once a paid plan has actually expired
    without being renewed — so the teacher is never left completely locked out."""
    from models import Message, PlanChangeHistory, TRIAL_DAYS, TRIAL_DOCS

    with app.app_context():
        now = datetime.utcnow()

        # ── Stage reminders (J-7 / J-3 / J-0) ──
        subs = Subscription.query.filter(
            Subscription.status == "active",
            Subscription.plan != "trial",
            Subscription.expires_at.isnot(None),
            Subscription.expires_at >= now,
        ).all()

        for sub in subs:
            days = (sub.expires_at - now).days
            # Smallest threshold still >= days-left: the most urgent bucket we're currently in.
            stage = next((s for s in sorted(REMINDER_STAGES) if days <= s), None)
            if stage is None:
                continue
            already_sent_this_or_more_urgent = (
                sub.last_reminder_stage is not None and sub.last_reminder_stage <= stage
            )
            if already_sent_this_or_more_urgent:
                continue

            user = User.query.get(sub.user_id)
            if not user:
                continue
            lang = user.preferred_lang or "fr"
            t = get_translations(lang)
            subject = t["email_reminder_subject"]
            body = t["email_reminder_body"].format(
                name=user.full_name, plan=sub.plan_label(), days=days,
                date=sub.expires_at.strftime("%d/%m/%Y"), email=contact_email,
            )
            send_email(user.email, subject, body)

            in_app_body = t["msg_expiry_reminder"].format(plan=sub.plan_label(), days=days,
                                                           date=sub.expires_at.strftime("%d/%m/%Y"))
            db.session.add(Message(user_id=user.id, sender="admin", body=in_app_body,
                                    read_by_admin=True, read_by_teacher=False))

            sub.last_reminder_stage = stage
            sub.expiry_reminder_sent = True

        # ── Auto-downgrade to a fresh trial once a paid plan has expired ──
        expired_subs = Subscription.query.filter(
            Subscription.status == "active",
            Subscription.plan != "trial",
            Subscription.expires_at.isnot(None),
            Subscription.expires_at < now,
        ).all()

        for sub in expired_subs:
            user = User.query.get(sub.user_id)
            if not user:
                continue
            previous_plan = sub.plan
            sub.plan = "trial"
            sub.status = "active"
            sub.docs_used = 0
            sub.docs_limit = TRIAL_DOCS
            sub.docs_used_today = 0
            sub.last_reset_date = None
            sub.expires_at = now + timedelta(days=TRIAL_DAYS)
            sub.expiry_reminder_sent = False
            sub.last_reminder_stage = None

            db.session.add(PlanChangeHistory(user_id=user.id, from_plan=previous_plan,
                                              to_plan="trial", reason="auto_downgrade"))

            lang = user.preferred_lang or "fr"
            t = get_translations(lang)
            db.session.add(Message(user_id=user.id, sender="admin",
                                    body=t["msg_auto_downgrade"].format(plan=previous_plan),
                                    read_by_admin=True, read_by_teacher=False))
            send_email(user.email, t["email_downgrade_subject"],
                       t["email_downgrade_body"].format(name=user.full_name, plan=previous_plan,
                                                         email=contact_email))

        db.session.commit()
