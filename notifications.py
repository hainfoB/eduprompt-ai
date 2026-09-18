import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

from translations import get_translations


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
    """Scan for subscriptions expiring within 5 days and send a one-time
    reminder email + flag them so the in-app banner can also show it."""
    with app.app_context():
        now = datetime.utcnow()
        subs = Subscription.query.filter(
            Subscription.status == "active",
            Subscription.expires_at.isnot(None),
            Subscription.expiry_reminder_sent.is_(False),
        ).all()

        for sub in subs:
            days = (sub.expires_at - now).days
            if 0 <= days <= 5:
                user = User.query.get(sub.user_id)
                if not user:
                    continue
                lang = user.preferred_lang or "fr"
                t = get_translations(lang)
                subject = t["email_reminder_subject"]
                body = t["email_reminder_body"].format(
                    name=user.full_name,
                    plan=sub.plan_label(),
                    days=days,
                    date=sub.expires_at.strftime("%d/%m/%Y"),
                    email=contact_email,
                )
                send_email(user.email, subject, body)
                sub.expiry_reminder_sent = True

        db.session.commit()
