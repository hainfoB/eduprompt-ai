"""Branded, trilingual HTML emails (site colours: navy #102a53, orange #f97316, white).
The recipient's language comes first; the two others follow below."""
from html import escape

NAVY, ORANGE, SKY, INK, MUTED, LINE = "#102a53", "#f97316", "#eaf2fb", "#142a44", "#60758d", "#dce6f0"

NUDGE = {
    1: {
        "fr": {
            "subject": "Votre premier document est à 2 minutes",
            "title": "Bonjour {name}, votre premier document vous attend",
            "intro": "Vous avez créé votre compte HaithemEduAI, mais vous n’avez pas encore généré de document. Il suffit de quelques clics :",
            "steps": ["Choisissez le niveau et la matière", "Indiquez la leçon que vous préparez", "Téléchargez votre document en Word ou en PDF"],
            "benefits": ["Fiches de préparation conformes au programme", "Évaluations et devoirs avec corrigé", "Modifiable dans Word, prêt à imprimer"],
            "cta": "Générer mon premier document",
            "ps": "Astuce : commencez par la leçon de cette semaine, vous gagnez du temps dès aujourd’hui.",
        },
        "ar": {
            "subject": "وثيقتك الأولى على بعد دقيقتين",
            "title": "مرحباً {name}، وثيقتك الأولى في انتظارك",
            "intro": "أنشأت حسابك على HaithemEduAI لكنك لم تولّد أي وثيقة بعد. يكفي بضع نقرات:",
            "steps": ["اختر المستوى والمادة", "حدّد الدرس الذي تحضّره", "حمّل وثيقتك بصيغة Word أو PDF"],
            "benefits": ["مذكرات مطابقة للمنهاج", "فروض واختبارات مع الحل", "قابلة للتعديل في Word وجاهزة للطباعة"],
            "cta": "ولّد وثيقتي الأولى",
            "ps": "نصيحة: ابدأ بدرس هذا الأسبوع، وستربح الوقت من اليوم.",
        },
        "en": {
            "subject": "Your first document is 2 minutes away",
            "title": "Hello {name}, your first document is waiting",
            "intro": "You created your HaithemEduAI account but haven’t generated a document yet. It only takes a few clicks:",
            "steps": ["Pick the level and subject", "Enter the lesson you are preparing", "Download your document in Word or PDF"],
            "benefits": ["Curriculum-aligned lesson plans", "Tests and homework with answer keys", "Editable in Word, ready to print"],
            "cta": "Generate my first document",
            "ps": "Tip: start with this week’s lesson and save time today.",
        },
    },
    2: {
        "fr": {
            "subject": "Besoin d’aide pour votre premier document ?",
            "title": "{name}, et si vous essayiez maintenant ?",
            "intro": "Votre essai HaithemEduAI est actif, mais aucun document n’a encore été généré. En quelques minutes, le générateur vous prépare un document complet, prêt à adapter.",
            "steps": ["Ouvrez le générateur", "Choisissez une leçon de votre programme", "Relisez, ajustez dans Word, imprimez"],
            "benefits": ["Gain de temps chaque semaine", "Documents en arabe, français ou anglais", "Retrouvez tous vos documents dans « Mes documents »"],
            "cta": "Essayer maintenant",
            "ps": "Quelque chose vous bloque ? Répondez simplement à cet e-mail, nous vous aidons.",
        },
        "ar": {
            "subject": "هل تحتاج مساعدة في وثيقتك الأولى؟",
            "title": "{name}، ما رأيك أن تجرّب الآن؟",
            "intro": "فترتك التجريبية على HaithemEduAI مفعّلة، لكن لم تُولَّد أي وثيقة بعد. في دقائق يحضّر لك المولّد وثيقة كاملة جاهزة للتكييف.",
            "steps": ["افتح المولّد", "اختر درساً من منهاجك", "راجع وعدّل في Word ثم اطبع"],
            "benefits": ["ربح الوقت كل أسبوع", "وثائق بالعربية أو الفرنسية أو الإنجليزية", "كل وثائقك محفوظة في «وثائقي»"],
            "cta": "جرّب الآن",
            "ps": "هل تواجه صعوبة؟ رد على هذا البريد وسنساعدك.",
        },
        "en": {
            "subject": "Need help with your first document?",
            "title": "{name}, why not try it now?",
            "intro": "Your HaithemEduAI trial is active, but no document has been generated yet. In a few minutes, the generator prepares a complete document, ready to adapt.",
            "steps": ["Open the generator", "Choose a lesson from your curriculum", "Review, adjust in Word, print"],
            "benefits": ["Save time every week", "Documents in Arabic, French or English", "All your documents kept in “My documents”"],
            "cta": "Try it now",
            "ps": "Something blocking you? Just reply to this email and we’ll help.",
        },
    },
}

FOOTER = {
    "fr": "Ne plus recevoir ces rappels",
    "ar": "إيقاف هذه التذكيرات",
    "en": "Stop these reminders",
}
LANG_LABEL = {"fr": "Français", "ar": "العربية", "en": "English"}


def _order(lang):
    lang = lang if lang in ("fr", "ar", "en") else "fr"
    return [lang] + [l for l in ("ar", "fr", "en") if l != lang]


def _block(c, lang, name, link, first):
    rtl = lang == "ar"
    d, align = ("rtl", "right") if rtl else ("ltr", "left")
    font = "'Segoe UI',Tahoma,Arial,sans-serif"
    steps = "".join(
        f'<tr><td style="padding:6px 0;vertical-align:top;width:34px">'
        f'<span style="display:inline-block;width:26px;height:26px;line-height:26px;border-radius:13px;'
        f'background:{ORANGE};color:#fff;font-weight:800;font-size:13px;text-align:center">{i}</span></td>'
        f'<td style="padding:6px 0;font-size:15px;color:{INK}">{escape(s)}</td></tr>'
        for i, s in enumerate(c["steps"], 1))
    benefits = "".join(
        f'<tr><td style="padding:4px 0;font-size:14px;color:{INK}">'
        f'<span style="color:{ORANGE};font-weight:800">✓</span>&nbsp; {escape(b)}</td></tr>'
        for b in c["benefits"])
    title_size = "22px" if first else "18px"
    sep = "" if first else f'<tr><td style="border-top:1px solid {LINE};padding-top:26px"></td></tr>'
    return f'''{sep}<tr><td dir="{d}" style="text-align:{align};font-family:{font};padding:0 0 26px">
  <div style="font-size:11px;font-weight:700;letter-spacing:.6px;color:{MUTED};text-transform:uppercase;margin-bottom:8px">{LANG_LABEL[lang]}</div>
  <h1 style="margin:0 0 12px;font-size:{title_size};line-height:1.35;color:{NAVY}">{escape(c["title"].format(name=name))}</h1>
  <p style="margin:0 0 16px;font-size:15px;line-height:1.65;color:{INK}">{escape(c["intro"])}</p>
  <table role="presentation" dir="{d}" cellpadding="0" cellspacing="0" style="margin:0 0 16px">{steps}</table>
  <table role="presentation" dir="{d}" width="100%" cellpadding="0" cellspacing="0" style="background:{SKY};border-radius:12px;margin:0 0 20px"><tr><td style="padding:14px 18px"><table role="presentation" dir="{d}" cellpadding="0" cellspacing="0">{benefits}</table></td></tr></table>
  <a href="{escape(link)}" style="display:inline-block;background:{ORANGE};color:#ffffff;text-decoration:none;font-weight:800;font-size:16px;padding:14px 28px;border-radius:10px">{escape(c["cta"])} {"←" if rtl else "→"}</a>
  <p style="margin:16px 0 0;font-size:13px;line-height:1.6;color:{MUTED}">{escape(c["ps"])}</p>
</td></tr>'''


def nudge_email(stage, lang, name, link, optout):
    """Return (subject, html, text) for activation follow-up `stage` (1 or 2)."""
    order = _order(lang)
    content = NUDGE[stage]
    subject = content[order[0]]["subject"]
    blocks = "".join(_block(content[l], l, name, link, i == 0) for i, l in enumerate(order))
    footer_links = " · ".join(
        f'<a href="{escape(optout)}" style="color:{MUTED}">{FOOTER[l]}</a>' for l in order)
    html = f'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape(subject)}</title></head>
<body style="margin:0;padding:0;background:#f6f8fc">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f6f8fc"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(16,42,83,.10)">
  <tr><td style="background:{NAVY};padding:22px 28px;font-family:'Segoe UI',Arial,sans-serif">
    <span style="font-size:22px;font-weight:800;color:#ffffff">Haithem<span style="color:{ORANGE}">Edu</span>AI</span>
    <div style="font-size:12px;color:#c7d5e8;margin-top:4px">منصة الأستاذ الذكية · La plateforme de l’enseignant · The teacher’s platform</div>
  </td></tr>
  <tr><td style="height:4px;background:{ORANGE};line-height:4px;font-size:0">&nbsp;</td></tr>
  <tr><td style="padding:28px 28px 4px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0">{blocks}</table></td></tr>
  <tr><td style="background:{SKY};padding:18px 28px;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;line-height:1.7;color:{MUTED};text-align:center">
    HaithemEduAI · Bordj Bou Arréridj, Algérie<br>{footer_links}
  </td></tr>
</table></td></tr></table></body></html>'''
    text_parts = []
    for l in order:
        c = content[l]
        text_parts.append("\n".join([
            c["title"].format(name=name), "", c["intro"], "",
            *[f"{i}. {s}" for i, s in enumerate(c["steps"], 1)], "",
            *[f"✓ {b}" for b in c["benefits"]], "",
            f'{c["cta"]} : {link}', "", c["ps"]]))
    text = "\n\n────────────\n\n".join(text_parts) + f"\n\n—\n{' / '.join(FOOTER[l] for l in order)} : {optout}"
    return subject, html, text


# ── Generic branded trilingual email (used by the training-pass emails) ──────
def _simple_block(lang, title, paras, rows, cta, link, first):
    rtl = lang == "ar"
    d, align = ("rtl", "right") if rtl else ("ltr", "left")
    font = "'Segoe UI',Tahoma,Arial,sans-serif"
    ps = "".join(f'<p style="margin:0 0 12px;font-size:15px;line-height:1.65;color:{INK}">{p}</p>' for p in paras)
    tbl = ""
    if rows:
        tbl = (f'<table role="presentation" dir="{d}" width="100%" cellpadding="0" cellspacing="0" '
               f'style="background:{SKY};border-radius:12px;margin:4px 0 18px">'
               + "".join(f'<tr><td style="padding:10px 18px;font-size:13px;color:{MUTED};width:40%">{escape(k)}</td>'
                         f'<td style="padding:10px 18px;font-size:15px;font-weight:800;color:{NAVY};direction:ltr;text-align:{align}">{escape(v)}</td></tr>'
                         for k, v in rows) + "</table>")
    sep = "" if first else f'<tr><td style="border-top:1px solid {LINE};padding-top:26px"></td></tr>'
    return f'''{sep}<tr><td dir="{d}" style="text-align:{align};font-family:{font};padding:0 0 26px">
  <div style="font-size:11px;font-weight:700;letter-spacing:.6px;color:{MUTED};text-transform:uppercase;margin-bottom:8px">{LANG_LABEL[lang]}</div>
  <h1 style="margin:0 0 12px;font-size:{"22px" if first else "18px"};line-height:1.35;color:{NAVY}">{escape(title)}</h1>
  {ps}{tbl}
  <a href="{escape(link)}" style="display:inline-block;background:{ORANGE};color:#ffffff;text-decoration:none;font-weight:800;font-size:16px;padding:14px 28px;border-radius:10px">{escape(cta)} {"←" if rtl else "→"}</a>
</td></tr>'''


def _wrap(subject, blocks, footer=""):
    return f'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape(subject)}</title></head>
<body style="margin:0;padding:0;background:#f6f8fc">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f6f8fc"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#ffffff;border-radius:16px;overflow:hidden">
  <tr><td style="background:{NAVY};padding:22px 28px;font-family:'Segoe UI',Arial,sans-serif">
    <span style="font-size:22px;font-weight:800;color:#ffffff">Haithem<span style="color:{ORANGE}">Edu</span>AI</span></td></tr>
  <tr><td style="height:4px;background:{ORANGE};line-height:4px;font-size:0">&nbsp;</td></tr>
  <tr><td style="padding:28px 28px 4px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0">{blocks}</table></td></tr>
  <tr><td style="background:{SKY};padding:18px 28px;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;line-height:1.7;color:{MUTED};text-align:center">HaithemEduAI · Bordj Bou Arréridj, Algérie{footer}</td></tr>
</table></td></tr></table></body></html>'''


CRED = {
    "fr": {"subject": "Vos accès HaithemEduAI (offerts par votre formation)",
           "title": "Bienvenue {name} !",
           "p1": "Suite à votre formation, un accès gratuit à HaithemEduAI vous est offert : <b>{plan}</b>, {quota}, pendant <b>{days} jours</b>.",
           "p2": "Connectez-vous avec les identifiants ci-dessous. Vous choisirez votre propre mot de passe à la première connexion.",
           "email": "E-mail", "password": "Mot de passe provisoire", "cta": "Me connecter",
           "q_month": "{n} documents par mois", "q_total": "{n} documents"},
    "ar": {"subject": "حسابك على HaithemEduAI (هدية من التكوين)",
           "title": "مرحباً {name}!",
           "p1": "بعد مشاركتك في التكوين، تحصّلت على وصول مجاني إلى HaithemEduAI: <b>{plan}</b>، {quota}، لمدة <b>{days} يوماً</b>.",
           "p2": "ادخل بالمعلومات أدناه، وستختار كلمة السر الخاصة بك عند أول دخول.",
           "email": "البريد الإلكتروني", "password": "كلمة السر المؤقتة", "cta": "الدخول إلى حسابي",
           "q_month": "{n} وثائق كل شهر", "q_total": "{n} وثيقة"},
    "en": {"subject": "Your HaithemEduAI access (a gift from your training)",
           "title": "Welcome {name}!",
           "p1": "Following your training, you get free access to HaithemEduAI: <b>{plan}</b>, {quota}, for <b>{days} days</b>.",
           "p2": "Sign in with the details below. You will choose your own password at first login.",
           "email": "Email", "password": "Temporary password", "cta": "Sign in",
           "q_month": "{n} documents per month", "q_total": "{n} documents"},
}


def credentials_email(lang, name, email, password, link, plan_label, days, docs, monthly):
    order = _order(lang)
    subject = CRED[order[0]]["subject"]
    blocks = ""
    for i, l in enumerate(order):
        c = CRED[l]
        quota = (c["q_month"] if monthly else c["q_total"]).format(n=docs)
        blocks += _simple_block(l, c["title"].format(name=name),
                                [c["p1"].format(plan=escape(plan_label), quota=quota, days=days), c["p2"]],
                                [(c["email"], email), (c["password"], password)], c["cta"], link, i == 0)
    text = "\n\n".join(f'{CRED[l]["title"].format(name=name)}\n{CRED[l]["email"]}: {email}\n{CRED[l]["password"]}: {password}\n{link}'
                       for l in order)
    return subject, _wrap(subject, blocks), text


REM = {
    "fr": {"subject": "Votre pass formation se termine le {date} : -{pct} % pour continuer",
           "title": "{name}, votre pass se termine bientôt",
           "p1": "Votre accès offert à HaithemEduAI prend fin le <b>{date}</b>.",
           "p2": "Pour continuer à préparer vos cours en quelques minutes, profitez de <b>-{pct} % sur l’abonnement annuel</b>, valable jusqu’au <b>{until}</b>. Vos documents restent disponibles dans « Mes documents ».",
           "cta": "Profiter de -{pct} %"},
    "ar": {"subject": "ينتهي باس التكوين يوم {date}: تخفيض {pct}% للمواصلة",
           "title": "{name}، اقترب موعد انتهاء الباس",
           "p1": "ينتهي وصولك المجاني إلى HaithemEduAI يوم <b>{date}</b>.",
           "p2": "لتواصل تحضير دروسك في دقائق، استفد من <b>تخفيض {pct}% على الاشتراك السنوي</b>، صالح إلى غاية <b>{until}</b>. وثائقك تبقى محفوظة في «وثائقي».",
           "cta": "أستفيد من تخفيض {pct}%"},
    "en": {"subject": "Your training pass ends on {date}: {pct}% off to keep going",
           "title": "{name}, your pass ends soon",
           "p1": "Your free HaithemEduAI access ends on <b>{date}</b>.",
           "p2": "To keep preparing your lessons in minutes, get <b>{pct}% off the annual plan</b>, valid until <b>{until}</b>. Your documents stay in “My documents”.",
           "cta": "Get {pct}% off"},
}


def pass_reminder_email(lang, name, date, pct, until, link, optout):
    order = _order(lang)
    subject = REM[order[0]]["subject"].format(date=date, pct=pct)
    blocks = "".join(_simple_block(l, REM[l]["title"].format(name=name),
                                   [REM[l]["p1"].format(date=date), REM[l]["p2"].format(pct=pct, until=until)],
                                   [], REM[l]["cta"].format(pct=pct), link, i == 0) for i, l in enumerate(order))
    footer = "<br>" + " · ".join(f'<a href="{escape(optout)}" style="color:{MUTED}">{FOOTER[l]}</a>' for l in order)
    text = "\n\n".join(REM[l]["p1"].format(date=date).replace("<b>", "").replace("</b>", "") + "\n" +
                       REM[l]["p2"].format(pct=pct, until=until).replace("<b>", "").replace("</b>", "") + f"\n{link}" for l in order)
    return subject, _wrap(subject, blocks, footer), text


# ── Training-session emails (registration confirmation, J-1 reminder) ────────
ENR = {
    "fr": {"label_session": "Formation", "label_when": "Date", "label_place": "Lieu", "label_link": "Lien de connexion", "label_group": "Groupe WhatsApp", "cal": "Ajouter à mon agenda", "ticket": "Mon ticket de présence (QR code)", "cta": "Voir la formation",
           "confirmed": ("Inscription confirmée : {title}", "{name}, votre place est réservée",
                         "Votre inscription à la formation <b>{title}</b> est confirmée. Nous vous enverrons un rappel la veille de la première séance."),
           "pending":   ("Inscription reçue : {title}", "{name}, nous avons bien reçu votre inscription",
                         "Votre inscription à <b>{title}</b> est en cours de validation. Vous recevrez un e-mail dès qu’elle sera confirmée."),
           "waitlist":  ("Liste d’attente : {title}", "{name}, vous êtes sur liste d’attente",
                         "La formation <b>{title}</b> est complète pour le moment. Vous êtes sur liste d’attente : si une place se libère, vous serez prévenu(e) automatiquement par e-mail."),
           "promoted":  ("Une place s’est libérée : {title}", "{name}, bonne nouvelle : une place s’est libérée !",
                         "Une place vient de se libérer pour <b>{title}</b>. Votre inscription est maintenant confirmée.")},
    "ar": {"label_session": "التكوين", "label_when": "التاريخ", "label_place": "المكان", "label_link": "رابط الحضور", "label_group": "مجموعة واتساب", "cal": "أضِف إلى جدولي", "ticket": "تذكرة حضوري (رمز QR)", "cta": "تفاصيل التكوين",
           "confirmed": ("تم تأكيد تسجيلك: {title}", "{name}، مقعدك محجوز",
                         "تم تأكيد تسجيلك في التكوين <b>{title}</b>. سنرسل لك تذكيراً عشية الحصة الأولى."),
           "pending":   ("استلمنا تسجيلك: {title}", "{name}، استلمنا طلب تسجيلك",
                         "تسجيلك في <b>{title}</b> قيد المراجعة. ستصلك رسالة بمجرد تأكيده."),
           "waitlist":  ("قائمة الانتظار: {title}", "{name}، أنت في قائمة الانتظار",
                         "التكوين <b>{title}</b> مكتمل حالياً. أنت في قائمة الانتظار، وإذا شغر مقعد سنُعلمك تلقائياً عبر البريد."),
           "promoted":  ("شغر مقعد لك: {title}", "{name}، خبر سار: شغر مقعد!",
                         "شغر مقعد في <b>{title}</b> وتم الآن تأكيد تسجيلك.")},
    "en": {"label_session": "Training", "label_when": "Date", "label_place": "Place", "label_link": "Join link", "label_group": "WhatsApp group", "cal": "Add to my calendar", "ticket": "My attendance ticket (QR code)", "cta": "View the training",
           "confirmed": ("Registration confirmed: {title}", "{name}, your seat is reserved",
                         "Your registration to <b>{title}</b> is confirmed. We will send you a reminder the day before the first session."),
           "pending":   ("Registration received: {title}", "{name}, we received your registration",
                         "Your registration to <b>{title}</b> is being reviewed. You will get an email as soon as it is confirmed."),
           "waitlist":  ("Waiting list: {title}", "{name}, you are on the waiting list",
                         "<b>{title}</b> is full for now. You are on the waiting list: if a seat frees up, you will be notified automatically by email."),
           "promoted":  ("A seat opened up: {title}", "{name}, good news: a seat opened up!",
                         "A seat just opened up for <b>{title}</b>. Your registration is now confirmed.")},
}


def enrollment_email(lang, name, status, title, when, place, link, online_url="", ar=None, extra=None):
    ar = ar or {}
    extra = extra or {}
    order = _order(lang)
    subject = ENR[order[0]][status][0].format(title=(ar.get("title") or title) if order[0] == "ar" else title)
    blocks, text_parts = "", []
    for i, l in enumerate(order):
        c = ENR[l]
        _, h, p = c[status]
        ttl = (ar.get("title") or title) if l == "ar" else title
        plc = (ar.get("place") or place) if l == "ar" else place
        rows = [(c["label_session"], ttl)]
        if when:
            rows.append((c["label_when"], when))
        if plc:
            rows.append((c["label_place"], plc))
        if online_url:
            rows.append((c["label_link"], online_url))
        if extra.get("group"):
            rows.append((c["label_group"], extra["group"]))
        paras = [p.format(title=escape(ttl))]
        if extra.get("gcal"):
            paras.append(f'📅 {c["cal"]} : <a href="{escape(extra["gcal"])}" style="color:#f97316;font-weight:700">Google</a> · '
                         f'<a href="{escape(extra["ics"])}" style="color:#f97316;font-weight:700">Outlook / Apple (.ics)</a>')
        if extra.get("ticket"):
            paras.append(f'🎫 <a href="{escape(extra["ticket"])}" style="color:#f97316;font-weight:700">{c["ticket"]}</a>')
        blocks += _simple_block(l, h.format(name=name), paras, rows, c["cta"], link, i == 0)
        text_parts.append(h.format(name=name) + "\n" + p.format(title=ttl).replace("<b>", "").replace("</b>", "") +
                          "".join(f"\n{k}: {v}" for k, v in rows[1:]) +
                          (f"\n{c['cal']}: {extra['gcal']}\n.ics: {extra['ics']}" if extra.get("gcal") else "") +
                          (f"\n{c['ticket']}: {extra['ticket']}" if extra.get("ticket") else "") + f"\n{link}")
    return subject, _wrap(subject, blocks), "\n\n".join(text_parts)


SRM = {
    "fr": {"subject": "Demain : {title}. Confirmez votre venue",
           "title": "{name}, on vous attend demain",
           "p1": "Rappel : la séance de <b>{title}</b> a lieu demain, <b>{when}</b>{place}.",
           "p2": "Pour nous aider à organiser la salle, confirmez-nous votre venue en un clic. Si vous ne pouvez pas venir, <a href=\"{no}\" style=\"color:#f97316;font-weight:700\">prévenez-nous ici</a> : votre place sera libérée pour une personne en liste d’attente.",
           "cta": "Je confirme ma venue"},
    "ar": {"subject": "غداً: {title}. أكّد حضورك",
           "title": "{name}، ننتظرك غداً",
           "p1": "تذكير: حصة <b>{title}</b> غداً، <b>{when}</b>{place}.",
           "p2": "لمساعدتنا في تنظيم القاعة، أكّد حضورك بنقرة واحدة. وإذا تعذّر عليك الحضور، <a href=\"{no}\" style=\"color:#f97316;font-weight:700\">أعلمنا من هنا</a> وسيُحرَّر مقعدك لصالح شخص في قائمة الانتظار.",
           "cta": "أؤكد حضوري"},
    "en": {"subject": "Tomorrow: {title}. Please confirm you are coming",
           "title": "{name}, we expect you tomorrow",
           "p1": "Reminder: the <b>{title}</b> session takes place tomorrow, <b>{when}</b>{place}.",
           "p2": "To help us organise the room, confirm you are coming in one click. If you cannot make it, <a href=\"{no}\" style=\"color:#f97316;font-weight:700\">let us know here</a>: your seat will be released to someone on the waiting list.",
           "cta": "I confirm I’m coming"},
}
_AT = {"fr": " à ", "ar": " · ", "en": " at "}


def session_reminder_email(lang, name, title, when, place, yes_link, no_link, online_url="", ar=None, group_url="", ticket_url=""):
    ar = ar or {}
    order = _order(lang)
    subject = SRM[order[0]]["subject"].format(title=(ar.get("title") or title) if order[0] == "ar" else title)
    blocks, text_parts = "", []
    for i, l in enumerate(order):
        c = SRM[l]
        ttl = (ar.get("title") or title) if l == "ar" else title
        plc = (ar.get("place") or place) if l == "ar" else place
        pl = f" ({escape(plc)})" if plc else ""
        blocks += _simple_block(l, c["title"].format(name=name),
                                [c["p1"].format(title=escape(ttl), when=escape(when), place=pl),
                                 c["p2"].format(no=escape(no_link))] +
                                ([f'🎫 <a href="{escape(ticket_url)}" style="color:#f97316;font-weight:700">{ENR[l]["ticket"]}</a>'] if ticket_url else []),
                                ([(ENR[l]["label_link"], online_url)] if online_url else []) +
                                ([(ENR[l]["label_group"], group_url)] if group_url else []), c["cta"], yes_link, i == 0)
        text_parts.append(c["title"].format(name=name) + "\n" +
                          c["p1"].format(title=ttl, when=when, place=f" ({plc})" if plc else "").replace("<b>", "").replace("</b>", "") +
                          (f"\n{ENR[l]['label_link']}: {online_url}" if online_url else "") +
                          (f"\n{ENR[l]['ticket']}: {ticket_url}" if ticket_url else "") +
                          f"\n{c['cta']}: {yes_link}\n→ {no_link}")
    return subject, _wrap(subject, blocks), "\n\n".join(text_parts)
