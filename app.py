import os
import json
import uuid
import re
import logging
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

import anthropic
import PyPDF2
import docx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# App & extensions
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-in-production")
app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv("DATABASE_URL", "sqlite:///dociq.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_CONTENT_LENGTH", 16 * 1024 * 1024))

UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "uploads")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "outputs")
ALLOWED_EXTENSIONS = {"pdf", "docx", "doc", "txt", "pptx", "md"}

Path(UPLOAD_FOLDER).mkdir(exist_ok=True)
Path(OUTPUT_FOLDER).mkdir(exist_ok=True)

CORS(app)
db = SQLAlchemy(app)
limiter = Limiter(get_remote_address, app=app, default_limits=[], storage_uri="memory://")

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

class Document(db.Model):
    __tablename__ = "documents"
    id            = db.Column(db.String(8),   primary_key=True)
    original_name = db.Column(db.String(255), nullable=False)
    stored_name   = db.Column(db.String(300), nullable=False)
    filepath      = db.Column(db.String(500), nullable=False)
    ext           = db.Column(db.String(10),  nullable=False)
    size          = db.Column(db.Integer,     nullable=False)
    word_count    = db.Column(db.Integer,     default=0)
    char_count    = db.Column(db.Integer,     default=0)
    has_text      = db.Column(db.Boolean,     default=True)
    uploaded_at   = db.Column(db.DateTime,    default=datetime.utcnow)
    text          = db.Column(db.Text,        nullable=True)
    analyses      = db.relationship("AnalysisResult", backref="document", cascade="all, delete-orphan")
    messages      = db.relationship("ChatMessage",    backref="document", cascade="all, delete-orphan")

    def to_dict(self, include_text=False):
        d = {
            "id":             self.id,
            "original_name":  self.original_name,
            "ext":            self.ext,
            "size":           self.size,
            "word_count":     self.word_count,
            "char_count":     self.char_count,
            "has_text":       self.has_text,
            "uploaded_at":    self.uploaded_at.isoformat(),
            "analysis":       any(a.analysis_type == "full" for a in self.analyses),
            "analysis_types": [a.analysis_type for a in self.analyses],
        }
        if include_text:
            d["text"] = self.text or ""
        return d


class AnalysisResult(db.Model):
    __tablename__ = "analysis_results"
    id            = db.Column(db.Integer,  primary_key=True, autoincrement=True)
    doc_id        = db.Column(db.String(8), db.ForeignKey("documents.id"), nullable=False)
    analysis_type = db.Column(db.String(20), nullable=False)
    result        = db.Column(db.Text,    nullable=False)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)
    model_used    = db.Column(db.String(50), nullable=True)
    __table_args__ = (db.UniqueConstraint("doc_id", "analysis_type", name="uq_doc_atype"),)


class ChatMessage(db.Model):
    __tablename__ = "chat_messages"
    id         = db.Column(db.Integer,    primary_key=True, autoincrement=True)
    session_id = db.Column(db.String(36), nullable=False, index=True)
    doc_id     = db.Column(db.String(8),  db.ForeignKey("documents.id"), nullable=True)
    role       = db.Column(db.String(10), nullable=False)
    content    = db.Column(db.Text,       nullable=False)
    created_at = db.Column(db.DateTime,   default=datetime.utcnow)


class CustomTemplate(db.Model):
    __tablename__ = "custom_templates"
    id          = db.Column(db.Integer,    primary_key=True, autoincrement=True)
    name        = db.Column(db.String(100), nullable=False)
    description = db.Column(db.String(255), nullable=True)
    prompt      = db.Column(db.Text,       nullable=False)
    created_at  = db.Column(db.DateTime,   default=datetime.utcnow)


with app.app_context():
    db.create_all()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_client():
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise ValueError("ANTHROPIC_API_KEY not set. Add it to your .env file.")
    return anthropic.Anthropic(api_key=key)


def extract_text(filepath):
    ext = filepath.rsplit(".", 1)[-1].lower()
    text = ""
    if ext == "pdf":
        with open(filepath, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text += (page.extract_text() or "") + "\n"
    elif ext in ("docx", "doc"):
        d = docx.Document(filepath)
        text = "\n".join(p.text for p in d.paragraphs)
    elif ext == "pptx":
        prs = Presentation(filepath)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    elif ext in ("txt", "md"):
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    return text.strip()


def trunc(text, n=50000):
    if len(text) <= n:
        return text
    cut = text[:n].rfind(". ")
    cut = cut if cut > int(n * 0.8) else n
    return text[:cut] + "\n\n[... document truncated ...]"


def call_claude(system, user, model="claude-opus-4-6"):
    client = get_client()
    resp = client.messages.create(
        model=model, max_tokens=4096, system=system,
        messages=[{"role": "user", "content": user}],
    )
    return resp.content[0].text


def build_pptx(title, slides_data, output_path):
    DARK = RGBColor(0x0D, 0x1B, 0x2A); ACCENT = RGBColor(0x00, 0xB4, 0xD8)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF); LGRAY = RGBColor(0xCC, 0xCC, 0xCC)
    prs = Presentation()
    prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    sl = prs.slides.add_slide(blank)
    sl.background.fill.solid(); sl.background.fill.fore_color.rgb = DARK
    tb = sl.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(1.5))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    run = p.add_run(); run.text = title
    run.font.size = Pt(40); run.font.bold = True; run.font.color.rgb = WHITE
    bar = sl.shapes.add_shape(1, Inches(4.5), Inches(4.2), Inches(4.33), Emu(36000))
    bar.fill.solid(); bar.fill.fore_color.rgb = ACCENT; bar.line.fill.background()
    tb2 = sl.shapes.add_textbox(Inches(1), Inches(4.6), Inches(11.33), Inches(0.5))
    p2 = tb2.text_frame.paragraphs[0]; p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = datetime.now().strftime("%B %d, %Y")
    r2.font.size = Pt(16); r2.font.color.rgb = LGRAY

    for info in slides_data:
        sl = prs.slides.add_slide(blank)
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb = DARK
        hdr = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Emu(180000))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = ACCENT; hdr.line.fill.background()
        tbt = sl.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.9))
        pt = tbt.text_frame.paragraphs[0]; rt = pt.add_run()
        rt.text = info.get("title", ""); rt.font.size = Pt(26)
        rt.font.bold = True; rt.font.color.rgb = DARK
        bullets = info.get("content", [])
        if isinstance(bullets, str):
            bullets = [bullets]
        tbc = sl.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12.33), Inches(5.5))
        tfc = tbc.text_frame; tfc.word_wrap = True
        for i, b in enumerate(bullets):
            para = tfc.paragraphs[i] if i == 0 else tfc.add_paragraph()
            para.space_before = Pt(6); rc = para.add_run()
            rc.text = f"• {b}" if not str(b).startswith("•") else str(b)
            rc.font.size = Pt(18); rc.font.color.rgb = WHITE
    prs.save(output_path)

# ---------------------------------------------------------------------------
# Page routes
# ---------------------------------------------------------------------------

def _stats():
    return {
        "total_docs":    Document.query.count(),
        "analyzed":      db.session.query(db.func.count(db.func.distinct(AnalysisResult.doc_id))).scalar() or 0,
        "total_words":   db.session.query(db.func.sum(Document.word_count)).scalar() or 0,
        "presentations": len(list(Path(OUTPUT_FOLDER).glob("*.pptx"))),
        "chats":         db.session.query(db.func.count(db.func.distinct(ChatMessage.session_id))).scalar() or 0,
    }


@app.route("/")
def index():
    recent = [d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).limit(5).all()]
    return render_template("index.html", stats=_stats(), recent=recent)


@app.route("/analyze")
def analyze_page():
    return render_template("analyze.html", documents=[d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()])


@app.route("/research")
def research_page():
    return render_template("research.html", documents=[d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()])


@app.route("/present")
def present_page():
    pptx_files = [f.name for f in Path(OUTPUT_FOLDER).glob("*.pptx")]
    return render_template("present.html",
        documents=[d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()],
        presentations=pptx_files)


@app.route("/documents")
def documents_page():
    return render_template("documents.html", documents=[d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()])

# ---------------------------------------------------------------------------
# API – Document management
# ---------------------------------------------------------------------------

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files["file"]
    if not file.filename or not allowed_file(file.filename):
        return jsonify({"error": "Unsupported file type. Allowed: pdf, docx, txt, pptx, md"}), 400

    doc_id   = str(uuid.uuid4())[:8]
    filename = secure_filename(file.filename)
    stored   = f"{doc_id}_{filename}"
    filepath = os.path.join(UPLOAD_FOLDER, stored)
    file.save(filepath)

    warning = None
    try:
        text = extract_text(filepath)
    except Exception as e:
        logger.error(f"Text extraction failed for {filename}: {e}")
        text = ""
        warning = "Text could not be extracted from this file (may be image-based or corrupted)."

    if not text and not warning:
        warning = "No extractable text found. The file may be image-based or empty."

    doc = Document(
        id=doc_id, original_name=filename, stored_name=stored, filepath=filepath,
        ext=filename.rsplit(".", 1)[-1].lower(), size=os.path.getsize(filepath),
        word_count=len(text.split()) if text else 0, char_count=len(text),
        has_text=bool(text), text=text,
    )
    db.session.add(doc)
    db.session.commit()

    resp = {"id": doc_id, "name": filename, "word_count": doc.word_count, "size": doc.size, "has_text": doc.has_text}
    if warning:
        resp["warning"] = warning
    return jsonify(resp)


@app.route("/api/documents")
def list_docs():
    return jsonify({"documents": [d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()]})


@app.route("/api/documents/<doc_id>", methods=["DELETE"])
def delete_doc(doc_id):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    filepath = doc.filepath
    db.session.delete(doc)
    db.session.commit()
    try:
        os.remove(filepath)
    except OSError:
        pass
    return jsonify({"message": "Document deleted"})


@app.route("/api/documents/<doc_id>/text")
def get_doc_text(doc_id):
    """Return raw extracted text for the preview modal."""
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    return jsonify({
        "id": doc.id, "name": doc.original_name,
        "text": doc.text or "", "word_count": doc.word_count,
        "char_count": doc.char_count, "has_text": doc.has_text,
    })

# ---------------------------------------------------------------------------
# API – Analysis (with per-type caching)
# ---------------------------------------------------------------------------

ANALYSIS_PROMPTS = {
    "full": (
        "Perform a comprehensive analysis.\n"
        "## Executive Summary\n## Key Themes\n## Main Arguments\n"
        "## Evidence & Data\n## Strengths & Gaps\n## Key Takeaways\n## Recommended Actions"
    ),
    "summary":   "Write a concise 3-5 paragraph executive summary covering purpose, main points, and conclusions.",
    "insights":  "Extract the top 10 key insights as numbered items with bold titles and brief explanations.",
    "sentiment": ("Analyze tone, sentiment, and rhetorical style: overall tone, emotional language, "
                  "persuasion techniques, objectivity level, and intended audience."),
    "entities":  ("Extract and categorize all key entities:\n"
                  "**People** | **Organizations** | **Locations** | **Dates/Times** | **Key Terms** | **Statistics**"),
}


@app.route("/api/analyze/<doc_id>", methods=["POST"])
@limiter.limit("20 per hour")
def analyze(doc_id):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not doc.has_text or not doc.text:
        return jsonify({"error": "This document has no extractable text (may be image-based)."}), 400

    data  = request.json or {}
    atype = data.get("type", "full")
    force = data.get("force", False)

    if atype not in ANALYSIS_PROMPTS:
        return jsonify({"error": f"Unknown analysis type '{atype}'"}), 400

    # Cache check — skip if force=true
    if not force:
        cached = AnalysisResult.query.filter_by(doc_id=doc_id, analysis_type=atype).first()
        if cached:
            logger.info(f"Cache hit: doc={doc_id} type={atype}")
            return jsonify({"result": cached.result, "type": atype, "cached": True})

    prompt = f"{ANALYSIS_PROMPTS[atype]}\n\nDocument content:\n\n{trunc(doc.text)}"
    system = ("You are an expert document analyst. Provide structured, insightful analysis. "
              "Use markdown headers, bullet points, and bold text for clarity.")
    try:
        result = call_claude(system, prompt)

        # Upsert
        existing = AnalysisResult.query.filter_by(doc_id=doc_id, analysis_type=atype).first()
        if existing:
            existing.result = result; existing.created_at = datetime.utcnow()
        else:
            db.session.add(AnalysisResult(doc_id=doc_id, analysis_type=atype,
                                          result=result, model_used="claude-opus-4-6"))
        db.session.commit()
        return jsonify({"result": result, "type": atype, "cached": False})
    except Exception as e:
        logger.error(f"Analysis error doc={doc_id} type={atype}: {e}")
        return jsonify({"error": str(e)}), 500

# ---------------------------------------------------------------------------
# API – Research chat (persisted)
# ---------------------------------------------------------------------------

@app.route("/api/chat", methods=["POST"])
@limiter.limit("30 per hour")
def chat():
    data       = request.json or {}
    doc_ids    = data.get("doc_ids", [])
    question   = data.get("question", "").strip()
    session_id = data.get("session_id") or str(uuid.uuid4())

    if not question:
        return jsonify({"error": "Question is required"}), 400
    if len(question) > 4000:
        return jsonify({"error": "Question too long (max 4,000 characters)"}), 400

    ctx_parts    = []
    primary_did  = None
    for did in doc_ids:
        doc = Document.query.get(did)
        if doc and doc.text:
            ctx_parts.append(f"### {doc.original_name}\n\n{trunc(doc.text, 20000)}")
            if primary_did is None:
                primary_did = did
    context = "\n\n---\n\n".join(ctx_parts) or "No documents selected."

    system = (
        "You are an expert research assistant. You have access to the following document(s):\n\n"
        f"{context}\n\n"
        "Answer questions accurately based on the content. Cite specific sections when relevant. "
        "If information is not in the documents, say so clearly. Use markdown for clarity."
    )
    try:
        answer = call_claude(system, question)
        db.session.add(ChatMessage(session_id=session_id, role="user",      content=question, doc_id=primary_did))
        db.session.add(ChatMessage(session_id=session_id, role="assistant", content=answer,   doc_id=primary_did))
        db.session.commit()
        return jsonify({"answer": answer, "session_id": session_id})
    except Exception as e:
        logger.error(f"Chat error session={session_id}: {e}")
        return jsonify({"error": str(e)}), 500

# ---------------------------------------------------------------------------
# API – Presentation generator
# ---------------------------------------------------------------------------

@app.route("/api/generate-presentation", methods=["POST"])
@limiter.limit("5 per hour")
def gen_presentation():
    data       = request.json or {}
    doc_ids    = data.get("doc_ids", [])
    title      = (data.get("title") or "Presentation").strip()
    num_slides = max(3, min(int(data.get("num_slides", 8)), 20))
    style      = data.get("style", "professional")

    if not doc_ids:
        return jsonify({"error": "Select at least one document"}), 400

    combined = ""
    for did in doc_ids:
        doc = Document.query.get(did)
        if doc and doc.text:
            combined += f"\n\n=== {doc.original_name} ===\n\n{doc.text}"

    if not combined.strip():
        return jsonify({"error": "Selected documents have no extractable text"}), 400

    system = ("You are a professional presentation designer. "
              "Create compelling, well-structured slides that communicate key ideas clearly.")
    prompt = (
        f"Create a {style} presentation titled '{title}' with exactly {num_slides} content slides "
        f"based on:\n\n{trunc(combined, 40000)}\n\n"
        "Return ONLY valid JSON (no markdown fences, no extra text):\n"
        '{"slides":[{"title":"...","content":["bullet1","bullet2","bullet3"]}]}\n'
        "Each slide: 3-5 concise bullet points. Titles must be compelling and actionable."
    )
    try:
        raw = call_claude(system, prompt)
        raw = re.sub(r"```(?:json)?\s*", "", raw).strip()
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError("Could not parse slide structure from AI response")
        slides_data = json.loads(m.group())["slides"]

        pid   = str(uuid.uuid4())[:8]
        safe  = secure_filename(title)[:40] or "presentation"
        fname = f"{pid}_{safe}.pptx"
        build_pptx(title, slides_data, os.path.join(OUTPUT_FOLDER, fname))

        return jsonify({"message": "Presentation generated", "filename": fname,
                        "slides": len(slides_data), "download_url": f"/api/download/{fname}"})
    except Exception as e:
        logger.error(f"Presentation error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/download/<filename>")
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, secure_filename(filename), as_attachment=True)


@app.route("/api/stats")
def stats():
    return jsonify(_stats())

# ---------------------------------------------------------------------------
# Reading complexity (pure Python, no AI)
# ---------------------------------------------------------------------------

def _count_syllables(word):
    word = word.lower().strip(".,!?;:\"'()-")
    if not word:
        return 0
    count = 0
    prev_vowel = False
    for ch in word:
        v = ch in "aeiouy"
        if v and not prev_vowel:
            count += 1
        prev_vowel = v
    if word.endswith("e") and count > 1:
        count -= 1
    return max(1, count)


def _flesch(text):
    import re as _re
    sentences = [s.strip() for s in _re.split(r"[.!?]+", text) if s.strip()]
    words     = [w for w in text.split() if w.strip(".,!?;:\"'()-")]
    if not sentences or not words:
        return None
    ns = len(sentences); nw = len(words)
    nsyl = sum(_count_syllables(w) for w in words)
    fre   = 206.835 - 1.015 * (nw / ns) - 84.6 * (nsyl / nw)
    grade = 0.39 * (nw / ns) + 11.8 * (nsyl / nw) - 15.59
    ease_label = (
        "Very Easy" if fre >= 90 else "Easy" if fre >= 80 else
        "Fairly Easy" if fre >= 70 else "Standard" if fre >= 60 else
        "Fairly Difficult" if fre >= 50 else "Difficult" if fre >= 30 else "Very Confusing"
    )
    return {
        "flesch_reading_ease":     round(max(0, min(100, fre)), 1),
        "ease_label":              ease_label,
        "grade_level":             round(max(0, grade), 1),
        "words":                   nw,
        "sentences":               ns,
        "syllables":               nsyl,
        "avg_words_per_sentence":  round(nw / ns, 1),
        "avg_syllables_per_word":  round(nsyl / nw, 2),
    }


@app.route("/api/complexity/<doc_id>")
def reading_complexity(doc_id):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not doc.text:
        return jsonify({"error": "No extractable text"}), 400
    result = _flesch(doc.text)
    if not result:
        return jsonify({"error": "Could not compute readability (too short)"}), 400
    return jsonify(result)


# ---------------------------------------------------------------------------
# Semantic search
# ---------------------------------------------------------------------------

@app.route("/api/search", methods=["POST"])
@limiter.limit("30 per hour")
def semantic_search():
    data  = request.json or {}
    query = data.get("query", "").strip()
    if not query:
        return jsonify({"error": "Query is required"}), 400

    docs = Document.query.filter(Document.has_text == True).all()
    if not docs:
        return jsonify({"results": [], "message": "No documents with text found"})

    # Build compact index for Claude to rank
    index = "\n".join(
        f"[{d.id}] {d.original_name}: {trunc(d.text, 600)}"
        for d in docs
    )
    system = "You are a semantic search engine. Rank documents by relevance to the query."
    prompt = (
        f"Query: {query}\n\nDocuments:\n{index}\n\n"
        "Return ONLY valid JSON (no fences):\n"
        '{"results":[{"doc_id":"...","score":0-100,"snippet":"...","reason":"..."}]}\n'
        "Sort by score descending. Include only docs with score > 20. snippet = 1-2 sentences most relevant to query."
    )
    try:
        raw = call_claude(system, prompt)
        raw = re.sub(r"```(?:json)?\s*", "", raw).strip()
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError("No JSON in response")
        results = json.loads(m.group()).get("results", [])
        # Attach document names
        doc_map = {d.id: d.original_name for d in docs}
        for r in results:
            r["name"] = doc_map.get(r.get("doc_id"), "Unknown")
        return jsonify({"results": results, "query": query})
    except Exception as e:
        logger.error(f"Search error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Multi-document comparison
# ---------------------------------------------------------------------------

@app.route("/api/compare", methods=["POST"])
@limiter.limit("10 per hour")
def compare_docs():
    data    = request.json or {}
    doc_ids = data.get("doc_ids", [])
    focus   = data.get("focus", "").strip()

    if len(doc_ids) < 2:
        return jsonify({"error": "Select at least 2 documents to compare"}), 400

    parts = []
    for did in doc_ids:
        doc = Document.query.get(did)
        if doc and doc.text:
            parts.append(f"### Document: {doc.original_name}\n\n{trunc(doc.text, 15000)}")
    if len(parts) < 2:
        return jsonify({"error": "At least 2 documents must have extractable text"}), 400

    combined = "\n\n---\n\n".join(parts)
    focus_line = f"Focus specifically on: {focus}\n\n" if focus else ""
    system = "You are an expert comparative document analyst."
    prompt = (
        f"{focus_line}Compare these documents in depth:\n\n{combined}\n\n"
        "Provide:\n## Overview\n## Key Agreements\n## Key Differences\n"
        "## Conflicting Claims\n## Unique Contributions per Document\n## Synthesis & Recommendations"
    )
    try:
        result = call_claude(system, prompt)
        return jsonify({"result": result})
    except Exception as e:
        logger.error(f"Compare error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Fact verification
# ---------------------------------------------------------------------------

@app.route("/api/verify-facts", methods=["POST"])
@limiter.limit("10 per hour")
def verify_facts():
    data              = request.json or {}
    source_id         = data.get("source_doc_id")
    reference_ids     = data.get("reference_doc_ids", [])

    if not source_id or not reference_ids:
        return jsonify({"error": "Provide a source document and at least one reference"}), 400

    source = Document.query.get(source_id)
    if not source or not source.text:
        return jsonify({"error": "Source document not found or has no text"}), 404

    refs = []
    for rid in reference_ids:
        d = Document.query.get(rid)
        if d and d.text:
            refs.append(f"### {d.original_name}\n\n{trunc(d.text, 12000)}")

    if not refs:
        return jsonify({"error": "No valid reference documents with text"}), 400

    system = "You are a fact-checking expert who verifies claims across documents."
    prompt = (
        f"Source document to fact-check:\n\n### {source.original_name}\n\n{trunc(source.text, 15000)}\n\n"
        f"Reference documents:\n\n{'---'.join(refs)}\n\n"
        "For each significant claim in the source document, verify it against the references.\n"
        "## ✅ Supported Claims\n## ❌ Contradicted Claims\n## ⚠️ Unverifiable Claims\n## 📊 Verification Summary"
    )
    try:
        result = call_claude(system, prompt)
        return jsonify({"result": result})
    except Exception as e:
        logger.error(f"Fact verify error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Contradiction detection
# ---------------------------------------------------------------------------

@app.route("/api/contradictions/<doc_id>", methods=["POST"])
@limiter.limit("15 per hour")
def detect_contradictions(doc_id):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not doc.text:
        return jsonify({"error": "Document has no extractable text"}), 400

    system = "You are a logical consistency analyst specializing in finding contradictions."
    prompt = (
        f"Analyze this document for internal contradictions and inconsistencies:\n\n"
        f"{trunc(doc.text)}\n\n"
        "Find:\n## Direct Contradictions (statements that directly conflict)\n"
        "## Logical Inconsistencies (claims that can't both be true)\n"
        "## Ambiguities (statements with conflicting interpretations)\n"
        "## Data Inconsistencies (conflicting numbers, dates, or facts)\n"
        "## Overall Consistency Assessment\n\n"
        "For each issue, quote the conflicting passages and explain why they contradict. "
        "If no contradictions found, say so clearly."
    )
    try:
        result = call_claude(system, prompt)
        return jsonify({"result": result})
    except Exception as e:
        logger.error(f"Contradiction error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Timeline extraction
# ---------------------------------------------------------------------------

@app.route("/api/timeline/<doc_id>", methods=["POST"])
@limiter.limit("15 per hour")
def extract_timeline(doc_id):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not doc.text:
        return jsonify({"error": "Document has no extractable text"}), 400

    system = "You are an expert at extracting temporal information from documents."
    prompt = (
        f"Extract all dates, events, deadlines, and time-based information from:\n\n{trunc(doc.text)}\n\n"
        "Return ONLY valid JSON (no fences):\n"
        '{"events":[{"date":"...","event":"...","description":"...","type":"deadline|milestone|historical|scheduled"}]}\n'
        "Sort chronologically. date can be approximate (e.g. 'Q1 2024', 'Early 2023'). "
        "Include all time references even if approximate."
    )
    try:
        raw = call_claude(system, prompt)
        raw = re.sub(r"```(?:json)?\s*", "", raw).strip()
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError("No JSON in response")
        events = json.loads(m.group()).get("events", [])
        return jsonify({"events": events, "doc_name": doc.original_name})
    except Exception as e:
        logger.error(f"Timeline error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Action item extraction
# ---------------------------------------------------------------------------

@app.route("/api/actions/<doc_id>", methods=["POST"])
@limiter.limit("15 per hour")
def extract_actions(doc_id):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not doc.text:
        return jsonify({"error": "Document has no extractable text"}), 400

    system = "You are an expert at extracting actionable tasks, commitments, and responsibilities."
    prompt = (
        f"Extract all action items, tasks, commitments, and responsibilities from:\n\n{trunc(doc.text)}\n\n"
        "Return ONLY valid JSON (no fences):\n"
        '{"actions":[{"task":"...","owner":"...","deadline":"...","priority":"high|medium|low","context":"..."}]}\n'
        "owner = person/team responsible (or 'Unassigned'). deadline = date or 'Not specified'. "
        "context = brief quote or context from the document."
    )
    try:
        raw = call_claude(system, prompt)
        raw = re.sub(r"```(?:json)?\s*", "", raw).strip()
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError("No JSON in response")
        actions = json.loads(m.group()).get("actions", [])
        return jsonify({"actions": actions, "doc_name": doc.original_name})
    except Exception as e:
        logger.error(f"Action items error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Document scoring against custom criteria
# ---------------------------------------------------------------------------

@app.route("/api/score", methods=["POST"])
@limiter.limit("10 per hour")
def score_document():
    data     = request.json or {}
    doc_id   = data.get("doc_id")
    criteria = data.get("criteria", "").strip()

    if not doc_id or not criteria:
        return jsonify({"error": "Document and criteria are required"}), 400

    doc = Document.query.get(doc_id)
    if not doc or not doc.text:
        return jsonify({"error": "Document not found or has no text"}), 404

    system = "You are an expert document evaluator who scores documents against specific criteria."
    prompt = (
        f"Score this document against the following criteria:\n\n"
        f"**Criteria / Rubric:**\n{criteria}\n\n"
        f"**Document:**\n{trunc(doc.text)}\n\n"
        "Provide:\n## Overall Score (X/10)\n## Criteria Breakdown\n"
        "(Score each criterion 0-10 with justification and specific evidence from the document)\n"
        "## Strengths\n## Gaps & Missing Elements\n## Recommendations to Improve Score"
    )
    try:
        result = call_claude(system, prompt)
        return jsonify({"result": result})
    except Exception as e:
        logger.error(f"Scoring error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Batch analysis
# ---------------------------------------------------------------------------

@app.route("/api/batch-analyze", methods=["POST"])
@limiter.limit("5 per hour")
def batch_analyze():
    data    = request.json or {}
    doc_ids = data.get("doc_ids", [])
    atype   = data.get("type", "summary")

    if not doc_ids:
        return jsonify({"error": "No documents selected"}), 400
    if atype not in ANALYSIS_PROMPTS:
        return jsonify({"error": f"Unknown analysis type '{atype}'"}), 400

    results = []
    for did in doc_ids:
        doc = Document.query.get(did)
        if not doc or not doc.text:
            results.append({"doc_id": did, "name": did, "error": "No text", "cached": False})
            continue

        # Check cache first
        cached = AnalysisResult.query.filter_by(doc_id=did, analysis_type=atype).first()
        if cached:
            results.append({"doc_id": did, "name": doc.original_name,
                            "result": cached.result, "cached": True})
            continue

        # Run analysis
        prompt = f"{ANALYSIS_PROMPTS[atype]}\n\nDocument content:\n\n{trunc(doc.text)}"
        system = ("You are an expert document analyst. Provide structured, insightful analysis. "
                  "Use markdown headers, bullet points, and bold text for clarity.")
        try:
            result = call_claude(system, prompt)
            db.session.add(AnalysisResult(doc_id=did, analysis_type=atype,
                                          result=result, model_used="claude-opus-4-6"))
            db.session.commit()
            results.append({"doc_id": did, "name": doc.original_name,
                            "result": result, "cached": False})
        except Exception as e:
            logger.error(f"Batch analysis error for {did}: {e}")
            results.append({"doc_id": did, "name": doc.original_name,
                            "error": str(e), "cached": False})

    return jsonify({"results": results})


# ---------------------------------------------------------------------------
# Custom analysis templates
# ---------------------------------------------------------------------------

@app.route("/api/templates", methods=["GET"])
def list_templates():
    tmpls = CustomTemplate.query.order_by(CustomTemplate.created_at.desc()).all()
    return jsonify({"templates": [
        {"id": t.id, "name": t.name, "description": t.description,
         "prompt": t.prompt, "created_at": t.created_at.isoformat()}
        for t in tmpls
    ]})


@app.route("/api/templates", methods=["POST"])
def create_template():
    data = request.json or {}
    name   = data.get("name", "").strip()
    prompt = data.get("prompt", "").strip()
    if not name or not prompt:
        return jsonify({"error": "Name and prompt are required"}), 400
    t = CustomTemplate(name=name, description=data.get("description", ""), prompt=prompt)
    db.session.add(t)
    db.session.commit()
    return jsonify({"id": t.id, "name": t.name, "message": "Template created"})


@app.route("/api/templates/<int:tmpl_id>", methods=["DELETE"])
def delete_template(tmpl_id):
    t = CustomTemplate.query.get(tmpl_id)
    if not t:
        return jsonify({"error": "Template not found"}), 404
    db.session.delete(t)
    db.session.commit()
    return jsonify({"message": "Template deleted"})


@app.route("/api/analyze-with-template/<doc_id>", methods=["POST"])
@limiter.limit("20 per hour")
def analyze_with_template(doc_id):
    doc = Document.query.get(doc_id)
    if not doc or not doc.text:
        return jsonify({"error": "Document not found or has no text"}), 404

    data     = request.json or {}
    tmpl_id  = data.get("template_id")
    tmpl     = CustomTemplate.query.get(tmpl_id)
    if not tmpl:
        return jsonify({"error": "Template not found"}), 404

    system = ("You are an expert document analyst. Provide structured, insightful analysis. "
              "Use markdown headers, bullet points, and bold text for clarity.")
    prompt = f"{tmpl.prompt}\n\nDocument content:\n\n{trunc(doc.text)}"
    try:
        result = call_claude(system, prompt)
        return jsonify({"result": result, "template": tmpl.name})
    except Exception as e:
        logger.error(f"Template analysis error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Export analysis as markdown
# ---------------------------------------------------------------------------

@app.route("/api/export/<doc_id>/<atype>")
def export_analysis(doc_id, atype):
    doc = Document.query.get(doc_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    cached = AnalysisResult.query.filter_by(doc_id=doc_id, analysis_type=atype).first()
    if not cached:
        return jsonify({"error": "No cached analysis found. Run the analysis first."}), 404

    from flask import Response
    content = (
        f"# {atype.title()} — {doc.original_name}\n"
        f"_Generated: {cached.created_at.strftime('%Y-%m-%d %H:%M')} · Model: {cached.model_used}_\n\n"
        f"---\n\n{cached.result}"
    )
    safe_name = secure_filename(doc.original_name.rsplit(".", 1)[0])[:40]
    filename  = f"{safe_name}_{atype}.md"
    return Response(
        content,
        mimetype="text/markdown",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ---------------------------------------------------------------------------
# New page routes
# ---------------------------------------------------------------------------

@app.route("/compare")
def compare_page():
    docs = [d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()]
    return render_template("compare.html", documents=docs)


@app.route("/search")
def search_page():
    return render_template("search.html")


@app.route("/tools")
def tools_page():
    docs  = [d.to_dict() for d in Document.query.order_by(Document.uploaded_at.desc()).all()]
    tmpls = [{"id": t.id, "name": t.name, "description": t.description, "prompt": t.prompt}
             for t in CustomTemplate.query.order_by(CustomTemplate.created_at.desc()).all()]
    return render_template("tools.html", documents=docs, templates=tmpls)


# ---------------------------------------------------------------------------
# Error handlers
# ---------------------------------------------------------------------------

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Maximum size is 16 MB."}), 413

@app.errorhandler(429)
def rate_limited(e):
    return jsonify({"error": "Rate limit reached. Please wait before making more AI requests."}), 429

@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404


if __name__ == "__main__":
    debug = os.getenv("FLASK_DEBUG", "true").lower() == "true"
    port  = int(os.getenv("PORT", 5000))
    app.run(debug=debug, host="0.0.0.0", port=port)
