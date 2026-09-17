import os
import re
import io
from datetime import datetime, timedelta

from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, flash
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

from models import db, User, Subscription, Document, LicenseCode, create_trial_subscription
from utils import build_sources_context

# ── APP SETUP ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-in-production")

# DATA_DIR points to a persistent volume in production (e.g. Railway mounts one at /data).
# Falls back to the app's own folder for local development.
DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(__file__))
os.makedirs(DATA_DIR, exist_ok=True)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(DATA_DIR, "eduprompt.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024  # 15 MB uploads

db.init_app(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Veuillez vous connecter pour accéder à cette page."


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


with app.app_context():
    db.create_all()

    # ── One-time admin bootstrap (idempotent — safe on every restart) ──
    _admin_email = os.environ.get("ADMIN_EMAIL")
    _admin_password = os.environ.get("ADMIN_PASSWORD")
    if _admin_email and _admin_password and not User.query.filter_by(email=_admin_email).first():
        _admin = User(
            first_name=os.environ.get("ADMIN_FIRST_NAME", "Admin"),
            last_name=os.environ.get("ADMIN_LAST_NAME", "EduPrompt"),
            email=_admin_email,
            role="admin",
        )
        _admin.set_password(_admin_password)
        db.session.add(_admin)
        db.session.flush()
        db.session.add(Subscription(
            user_id=_admin.id, plan="premium", status="active",
            expires_at=datetime.utcnow() + timedelta(days=3650),
            docs_used=0, docs_limit=999999,
        ))
        db.session.commit()
        print(f"✅ Admin account bootstrapped: {_admin_email}")


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
    if current_user.is_authenticated:
        return redirect(url_for("generator"))
    return redirect(url_for("login"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("generator"))

    if request.method == "POST":
        first_name = request.form.get("first_name", "").strip()
        last_name  = request.form.get("last_name", "").strip()
        email      = request.form.get("email", "").strip().lower()
        password   = request.form.get("password", "")

        if not all([first_name, last_name, email, password]):
            flash("Tous les champs sont obligatoires.", "error")
            return render_template("register.html")

        if len(password) < 6:
            flash("Le mot de passe doit contenir au moins 6 caractères.", "error")
            return render_template("register.html")

        if User.query.filter_by(email=email).first():
            flash("Un compte existe déjà avec cet email.", "error")
            return render_template("register.html")

        user = User(first_name=first_name, last_name=last_name, email=email)
        user.set_password(password)
        db.session.add(user)
        db.session.flush()  # get user.id before commit
        create_trial_subscription(user)
        db.session.commit()

        login_user(user)
        flash("Bienvenue ! Votre essai gratuit de 14 jours a commencé.", "success")
        return redirect(url_for("generator"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("generator"))

    if request.method == "POST":
        email    = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        user = User.query.filter_by(email=email).first()

        if user and user.check_password(password):
            login_user(user, remember=True)
            return redirect(url_for("generator"))
        flash("Email ou mot de passe incorrect.", "error")

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
    return render_template("dashboard.html", sub=sub, docs=recent_docs)


@app.route("/upgrade", methods=["GET", "POST"])
@login_required
def upgrade():
    if request.method == "POST":
        code_str = request.form.get("code", "").strip().upper()
        lic = LicenseCode.query.filter_by(code=code_str, used=False).first()

        if not lic:
            flash("Code invalide ou déjà utilisé.", "error")
            return render_template("upgrade.html")

        sub = current_user.subscription
        sub.plan = lic.plan
        sub.status = "active"
        sub.started_at = datetime.utcnow()
        sub.expires_at = datetime.utcnow() + timedelta(days=lic.duration_days)
        sub.docs_used = 0

        lic.used = True
        lic.used_by = current_user.id
        lic.used_at = datetime.utcnow()

        db.session.commit()
        flash(f"🎉 Compte mis à niveau vers {sub.plan_label()} !", "success")
        return redirect(url_for("dashboard"))

    return render_template("upgrade.html")


# ── ADMIN: generate license codes ────────────────────────────────────────────
@app.route("/admin/licenses", methods=["GET", "POST"])
@login_required
def admin_licenses():
    if current_user.role != "admin":
        flash("Accès réservé aux administrateurs.", "error")
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        plan = request.form.get("plan", "pro")
        days = int(request.form.get("days", 365))
        lic = LicenseCode(plan=plan, duration_days=days)
        db.session.add(lic)
        db.session.commit()
        flash(f"Code généré : {lic.code}", "success")

    codes = LicenseCode.query.order_by(LicenseCode.created_at.desc()).limit(50).all()
    return render_template("admin_licenses.html", codes=codes)


# ── MAIN GENERATOR PAGE ───────────────────────────────────────────────────────
@app.route("/generator")
@login_required
def generator():
    return render_template("generator.html",
                           levels=LEVELS, subjects=SUBJECTS, doc_types=DOCUMENT_TYPES)


# ── API: QUOTA CHECK ──────────────────────────────────────────────────────────
@app.route("/api/check-quota")
@login_required
def check_quota():
    sub = current_user.subscription
    return jsonify({
        "allowed": sub.has_quota(),
        "remaining": sub.remaining(),
        "plan": sub.plan,
        "plan_label": sub.plan_label(),
    })


# ── API: LOG USAGE (after successful AI generation) ──────────────────────────
@app.route("/api/log-usage", methods=["POST"])
@login_required
def log_usage():
    data = request.get_json() or {}
    sub = current_user.subscription

    if not sub.has_quota():
        return jsonify({"error": "quota_exceeded"}), 403

    if not sub.is_unlimited():
        sub.docs_used += 1

    doc = Document(
        user_id=current_user.id,
        title=data.get("lesson", "")[:255],
        subject=data.get("subject", ""),
        level=data.get("level", ""),
        palier=data.get("palier", ""),
        doc_type=data.get("doc_type", ""),
        fmt=data.get("format", "docx"),
    )
    db.session.add(doc)
    db.session.commit()

    return jsonify({"ok": True, "remaining": sub.remaining()})


# ── API: EXTRACT SOURCES (NotebookLM-style: files + links) ──────────────────
@app.route("/api/extract-sources", methods=["POST"])
@login_required
def extract_sources():
    files = request.files.getlist("files")
    urls_raw = request.form.get("urls", "")
    urls = [u.strip() for u in urls_raw.split("\n") if u.strip()]

    context, warnings = build_sources_context(files, urls)
    return jsonify({"context": context, "warnings": warnings})


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


def make_docx(content, meta, colors_cfg):
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
    run = title_p.add_run("EduPrompt AI")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(*hex_to_rgb(c1))
    title_p.paragraph_format.space_after = Pt(14)

    # ── Header table: Teacher / Level / Palier / Subject ──
    table = doc.add_table(rows=2, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    headers = ["Enseignant(e)", "Niveau", "Palier", "Matière"]
    values  = [
        f"{meta.get('teacher_first','')} {meta.get('teacher_last','')}".strip() or "—",
        meta.get("level", "—"),
        meta.get("palier", "—"),
        meta.get("subject", "—"),
    ]

    for i, (h, v) in enumerate(zip(headers, values)):
        hc = table.cell(0, i)
        hc.text = ""
        hp = hc.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = hp.add_run(h)
        hr.bold = True
        hr.font.size = Pt(10)
        hr.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(hc, c1)
        hc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        vc = table.cell(1, i)
        vc.text = ""
        vp = vc.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
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

    # ── Body content (markdown-ish parsing) ──
    for line in content.split("\n"):
        line = line.strip()
        if not line:
            doc.add_paragraph("")
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
        elif line.startswith("- ") or line.startswith("• "):
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(line[2:])
            r.font.color.rgb = RGBColor(*hex_to_rgb(ctext, "1a1a1a"))
        else:
            p = doc.add_paragraph()
            parts = line.split("**")
            for i, part in enumerate(parts):
                r = p.add_run(part)
                r.font.color.rgb = RGBColor(*hex_to_rgb(ctext, "1a1a1a"))
                if i % 2 == 1:
                    r.bold = True

    # ── Footer ──
    doc.add_paragraph().paragraph_format.space_before = Pt(20)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer_p.add_run("Généré avec EduPrompt AI")
    fr.font.size = Pt(8)
    fr.font.color.rgb = RGBColor(150, 150, 150)
    fr.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def make_pdf(content, meta, colors_cfg, lang="fr"):
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

    story = [Paragraph("EduPrompt AI", title_s)]

    teacher = f"{meta.get('teacher_first','')} {meta.get('teacher_last','')}".strip() or "—"
    table_data = [
        ["Enseignant(e)", "Niveau", "Palier", "Matière"],
        [teacher, meta.get("level","—"), meta.get("palier","—"), meta.get("subject","—")],
    ]
    tbl = Table(table_data, colWidths=[4.2*cm]*4)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), rl_colors.HexColor(c1)),
        ("TEXTCOLOR", (0,0), (-1,0), rl_colors.white),
        ("TEXTCOLOR", (0,1), (-1,1), rl_colors.HexColor(ctext)),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.75, rl_colors.HexColor("#dce6f0")),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(meta.get("lesson",""), lesson_s))

    for line in content.split("\n"):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 0.2*cm))
        elif line.startswith("### ") or line.startswith("## ") or line.startswith("# "):
            style = h1_s if line.startswith("# ") else h2_s
            story.append(Paragraph(line.lstrip("#").strip(), style))
        elif line.startswith("- ") or line.startswith("• "):
            story.append(Paragraph("• " + line[2:], blt_s))
        else:
            safe = line.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            safe = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', safe)
            story.append(Paragraph(safe, body_s))

    doc.build(story)
    buf.seek(0)
    return buf


# ── API: DOWNLOAD ──────────────────────────────────────────────────────────────
@app.route("/api/download", methods=["POST"])
@login_required
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
    }
    colors_cfg = data.get("colors", {})

    safe_name = re.sub(r'[^\w\-]', '_', meta["lesson"][:40] or "document")

    if fmt == "pdf":
        buf = make_pdf(content, meta, colors_cfg, lang)
        return send_file(buf, as_attachment=True, download_name=f"{safe_name}.pdf",
                         mimetype="application/pdf")
    else:
        buf = make_docx(content, meta, colors_cfg)
        return send_file(buf, as_attachment=True, download_name=f"{safe_name}.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
