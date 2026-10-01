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
