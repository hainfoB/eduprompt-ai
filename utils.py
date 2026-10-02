import io
import re
import base64
import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader
from docx import Document as DocxDocument

MAX_SOURCE_CHARS   = 6000   # per source
MAX_TOTAL_CHARS    = 18000  # combined context sent to the AI
# Files the server cannot read as text (photos of a textbook page, scanned PDFs)
# are handed to Gemini as-is: it reads images and PDFs natively, Arabic included.
MAX_ATTACHMENT_BYTES = 9 * 1024 * 1024  # all attachments together, before base64
IMAGE_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
               ".webp": "image/webp", ".heic": "image/heic", ".heif": "image/heif"}
MIN_PDF_CHARS_PER_PAGE = 80  # below this a PDF is treated as scanned


def _docx_text(data):
    """Paragraphs and table cells, in document order (curricula are often tables)."""
    doc = DocxDocument(io.BytesIO(data))
    body = doc.element.body
    out = []
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            t = "".join(n.text or "" for n in child.iter() if n.tag.endswith("}t"))
            if t.strip():
                out.append(t)
        elif tag == "tbl":
            for row in child.iter():
                if not row.tag.endswith("}tr"):
                    continue
                cells = []
                for cell in row.iterchildren():
                    if cell.tag.endswith("}tc"):
                        cells.append(" ".join("".join(n.text or "" for n in para.iter() if n.tag.endswith("}t"))
                                              for para in cell.iter() if para.tag.endswith("}p")).strip())
                if any(cells):
                    out.append(" | ".join(cells))
    return "\n".join(out)


def extract_text_from_file(file_storage):
    """Read an uploaded source. Returns (text, error, attachment): text for files
    we can read, or an attachment {mime, data, name, size} for images and scanned
    PDFs, which the model reads directly."""
    name = file_storage.filename or ""
    filename = name.lower()
    ext = "." + filename.rsplit(".", 1)[-1] if "." in filename else ""
    data = file_storage.read()
    if not data:
        return None, f"Fichier vide : {name}", None

    def attach(mime):
        return None, None, {"mime": mime, "name": name, "size": len(data),
                            "data": base64.b64encode(data).decode("ascii")}

    try:
        if ext in IMAGE_TYPES:
            return attach(IMAGE_TYPES[ext])
        if ext == ".pdf":
            reader = PdfReader(io.BytesIO(data))
            pages = len(reader.pages) or 1
            text = "\n".join((page.extract_text() or "") for page in reader.pages).strip()
            if len(text) < MIN_PDF_CHARS_PER_PAGE * min(pages, 5):
                return attach("application/pdf")  # scanned or image-only PDF
        elif ext == ".docx":
            text = _docx_text(data).strip()
        elif ext in (".txt", ".md"):
            text = data.decode("utf-8", errors="ignore").strip()
        elif ext == ".doc":
            return None, f"{name} : ancien format Word (.doc). Enregistrez-le en .docx ou en PDF.", None
        else:
            return None, f"Format non supporté : {name}", None
    except Exception as e:
        if ext == ".pdf":
            return attach("application/pdf")  # unreadable for us, the model may still read it
        return None, f"Erreur de lecture de {name} : {e}", None

    if not text:
        return None, f"Aucun texte trouvé dans {name}", None
    return text[:MAX_SOURCE_CHARS], None, None


def extract_text_from_url(url):
    """Fetch a web page and extract its readable text content."""
    if not re.match(r'^https?://', url):
        url = "https://" + url
    try:
        resp = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
            tag.decompose()
        text = soup.get_text(separator="\n")
        text = re.sub(r'\n{3,}', '\n\n', text).strip()
        if not text:
            return None, f"Aucun contenu extrait de {url}"
        return text[:MAX_SOURCE_CHARS], None
    except Exception as e:
        return None, f"Erreur d'accès à {url}: {e}"


def build_sources_context(files, urls):
    """
    files: list of werkzeug FileStorage
    urls:  list of str
    Returns (context_text, warnings_list, attachments_list)
    """
    parts, warnings, attachments = [], [], []
    attached_bytes = 0

    for f in files or []:
        if not f or not f.filename:
            continue
        text, err, att = extract_text_from_file(f)
        if err:
            warnings.append(err)
        elif att:
            if attached_bytes + att["size"] > MAX_ATTACHMENT_BYTES:
                warnings.append(f"{att['name']} : trop volumineux (9 Mo au total pour les photos et PDF scannés)")
                continue
            attached_bytes += att["size"]
            attachments.append(att)
            parts.append(f"### Source (fichier joint : {f.filename}, lu directement)")
        else:
            parts.append(f"### Source (fichier: {f.filename})\n{text}")

    for u in urls or []:
        u = (u or "").strip()
        if not u:
            continue
        text, err = extract_text_from_url(u)
        if err:
            warnings.append(err)
        else:
            parts.append(f"### Source (lien: {u})\n{text}")

    context = "\n\n".join(parts)
    if len(context) > MAX_TOTAL_CHARS:
        context = context[:MAX_TOTAL_CHARS] + "\n[...contenu tronqué...]"

    return context, warnings, attachments
