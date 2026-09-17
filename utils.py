import io
import re
import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader
from docx import Document as DocxDocument

MAX_SOURCE_CHARS   = 6000   # per source
MAX_TOTAL_CHARS    = 18000  # combined context sent to the AI


def extract_text_from_file(file_storage):
    """Extract plain text from an uploaded PDF, DOCX, TXT or MD file."""
    filename = (file_storage.filename or "").lower()
    data = file_storage.read()
    text = ""
    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(data))
            text = "\n".join((page.extract_text() or "") for page in reader.pages)
        elif filename.endswith(".docx"):
            doc = DocxDocument(io.BytesIO(data))
            text = "\n".join(p.text for p in doc.paragraphs)
        elif filename.endswith((".txt", ".md")):
            text = data.decode("utf-8", errors="ignore")
        else:
            return None, f"Format non supporté: {filename}"
    except Exception as e:
        return None, f"Erreur lecture {filename}: {e}"

    text = text.strip()
    if not text:
        return None, f"Aucun texte extrait de {filename}"
    return text[:MAX_SOURCE_CHARS], None


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
    Returns (context_text, warnings_list)
    """
    parts, warnings = [], []

    for f in files or []:
        if not f or not f.filename:
            continue
        text, err = extract_text_from_file(f)
        if err:
            warnings.append(err)
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

    return context, warnings
