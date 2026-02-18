import os
import json
import uuid
import re
import logging
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_cors import CORS
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

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-in-production")
CORS(app)

UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "uploads")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "outputs")
ALLOWED_EXTENSIONS = {"pdf", "docx", "doc", "txt", "pptx", "md"}
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_CONTENT_LENGTH", 16 * 1024 * 1024))

Path(UPLOAD_FOLDER).mkdir(exist_ok=True)
Path(OUTPUT_FOLDER).mkdir(exist_ok=True)

document_store: dict = {}
chat_sessions: dict = {}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_client():
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise ValueError("ANTHROPIC_API_KEY not set")
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
    return text[:n] + "\n\n[... truncated ...]"


def claude(system, user, model="claude-opus-4-6"):
    client = get_client()
    resp = client.messages.create(
        model=model,
        max_tokens=4096,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return resp.content[0].text


def build_pptx(title, slides_data, output_path):
    DARK = RGBColor(0x0D, 0x1B, 0x2A)
    ACCENT = RGBColor(0x00, 0xB4, 0xD8)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    LGRAY = RGBColor(0xCC, 0xCC, 0xCC)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Title slide
    sl = prs.slides.add_slide(blank)
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = DARK

    tb = sl.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = WHITE

    bar = sl.shapes.add_shape(1, Inches(4.5), Inches(4.2), Inches(4.33), Emu(36000))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()

    tb2 = sl.shapes.add_textbox(Inches(1), Inches(4.6), Inches(11.33), Inches(0.5))
    p2 = tb2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = datetime.now().strftime("%B %d, %Y")
    r2.font.size = Pt(16)
    r2.font.color.rgb = LGRAY

    for info in slides_data:
        sl = prs.slides.add_slide(blank)
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = DARK

        hdr = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Emu(180000))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = ACCENT
        hdr.line.fill.background()

        tbt = sl.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.9))
        pt = tbt.text_frame.paragraphs[0]
        rt = pt.add_run()
        rt.text = info.get("title", "")
        rt.font.size = Pt(26)
        rt.font.bold = True
        rt.font.color.rgb = DARK

        bullets = info.get("content", [])
        if isinstance(bullets, str):
            bullets = [bullets]

        tbc = sl.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12.33), Inches(5.5))
        tfc = tbc.text_frame
        tfc.word_wrap = True
        for i, b in enumerate(bullets):
            para = tfc.paragraphs[i] if i == 0 else tfc.add_paragraph()
            para.space_before = Pt(6)
            rc = para.add_run()
            rc.text = f"• {b}" if not str(b).startswith("•") else str(b)
            rc.font.size = Pt(18)
            rc.font.color.rgb = WHITE

    prs.save(output_path)


# ---------------------------------------------------------------------------
# Page routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    stats = {
        "total_docs": len(document_store),
        "analyzed": sum(1 for d in document_store.values() if d.get("analysis")),
        "presentations": len(list(Path(OUTPUT_FOLDER).glob("*.pptx"))),
        "chats": len(chat_sessions),
        "total_words": sum(d.get("word_count", 0) for d in document_store.values()),
    }
    recent = sorted(document_store.values(), key=lambda d: d.get("uploaded_at", ""), reverse=True)[:5]
    return render_template("index.html", stats=stats, recent=recent)


@app.route("/analyze")
def analyze_page():
    return render_template("analyze.html", documents=list(document_store.values()))


@app.route("/research")
def research_page():
    return render_template("research.html", documents=list(document_store.values()))


@app.route("/present")
def present_page():
    pptx_files = [f.name for f in Path(OUTPUT_FOLDER).glob("*.pptx")]
    return render_template("present.html", documents=list(document_store.values()), presentations=pptx_files)


@app.route("/documents")
def documents_page():
    return render_template("documents.html", documents=list(document_store.values()))


# ---------------------------------------------------------------------------
# API – uploads & document management
# ---------------------------------------------------------------------------

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files["file"]
    if not file.filename or not allowed_file(file.filename):
        return jsonify({"error": "Invalid or unsupported file type"}), 400

    doc_id = str(uuid.uuid4())[:8]
    filename = secure_filename(file.filename)
    stored = f"{doc_id}_{filename}"
    filepath = os.path.join(UPLOAD_FOLDER, stored)
    file.save(filepath)

    try:
        text = extract_text(filepath)
    except Exception as e:
        logger.error(f"Extract error: {e}")
        text = ""

    doc = {
        "id": doc_id,
        "original_name": filename,
        "stored_name": stored,
        "filepath": filepath,
        "ext": filename.rsplit(".", 1)[-1].lower(),
        "size": os.path.getsize(filepath),
        "word_count": len(text.split()),
        "char_count": len(text),
        "uploaded_at": datetime.now().isoformat(),
        "text": text,
        "analysis": None,
        "summary": None,
    }
    document_store[doc_id] = doc
    return jsonify({"id": doc_id, "name": filename, "word_count": doc["word_count"], "size": doc["size"]})


@app.route("/api/documents")
def list_docs():
    docs = [{k: v for k, v in d.items() if k != "text"} for d in document_store.values()]
    return jsonify({"documents": docs})


@app.route("/api/documents/<doc_id>", methods=["DELETE"])
def delete_doc(doc_id):
    if doc_id not in document_store:
        return jsonify({"error": "Not found"}), 404
    doc = document_store.pop(doc_id)
    try:
        os.remove(doc["filepath"])
    except OSError:
        pass
    return jsonify({"message": "Deleted"})


# ---------------------------------------------------------------------------
# API – Analysis
# ---------------------------------------------------------------------------

ANALYSIS_PROMPTS = {
    "full": (
        "Perform a comprehensive analysis.\n"
        "## Executive Summary\n## Key Themes\n## Main Arguments\n"
        "## Evidence & Data\n## Strengths & Gaps\n## Key Takeaways\n## Recommended Actions"
    ),
    "summary": "Write a concise 3-5 paragraph executive summary covering purpose, main points, and conclusions.",
    "insights": "Extract the top 10 key insights as numbered items with bold titles and brief explanations.",
    "sentiment": (
        "Analyze tone, sentiment, and rhetorical style: overall tone, emotional language, "
        "persuasion techniques, objectivity, and intended audience."
    ),
    "entities": (
        "Extract and categorize all key entities:\n"
        "**People** | **Organizations** | **Locations** | **Dates/Times** | **Key Terms** | **Statistics**"
    ),
}


@app.route("/api/analyze/<doc_id>", methods=["POST"])
def analyze(doc_id):
    if doc_id not in document_store:
        return jsonify({"error": "Document not found"}), 404
    doc = document_store[doc_id]
    text = doc.get("text", "")
    if not text:
        return jsonify({"error": "Document has no extractable text"}), 400

    atype = (request.json or {}).get("type", "full")
    task = ANALYSIS_PROMPTS.get(atype, ANALYSIS_PROMPTS["full"])
    prompt = f"{task}\n\nDocument content:\n\n{trunc(text)}"

    system = (
        "You are an expert document analyst. Provide structured, insightful analysis. "
        "Use markdown headers, bullet points, and bold text for clarity."
    )
    try:
        result = claude(system, prompt)
        doc["analysis"] = result
        if atype == "summary":
            doc["summary"] = result
        return jsonify({"result": result, "type": atype})
    except Exception as e:
        logger.error(f"Analysis error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# API – Research chat
# ---------------------------------------------------------------------------

@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json or {}
    doc_ids = data.get("doc_ids", [])
    question = data.get("question", "").strip()
    session_id = data.get("session_id") or str(uuid.uuid4())

    if not question:
        return jsonify({"error": "Question is required"}), 400

    ctx_parts = []
    for did in doc_ids:
        if did in document_store:
            d = document_store[did]
            ctx_parts.append(f"### {d['original_name']}\n\n{trunc(d['text'], 20000)}")
    context = "\n\n---\n\n".join(ctx_parts) or "No documents selected."

    if session_id not in chat_sessions:
        chat_sessions[session_id] = []
    history = chat_sessions[session_id]
    history.append({"role": "user", "content": question})

    system = (
        "You are an expert research assistant. You have access to the following document(s):\n\n"
        f"{context}\n\n"
        "Answer questions accurately based on the content. Cite specific sections when relevant. "
        "If information is not in the documents, say so clearly. Use markdown for clarity."
    )
    try:
        answer = claude(system, question)
        history.append({"role": "assistant", "content": answer})
        chat_sessions[session_id] = history[-20:]
        return jsonify({"answer": answer, "session_id": session_id})
    except Exception as e:
        logger.error(f"Chat error: {e}")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# API – Presentation generator
# ---------------------------------------------------------------------------

@app.route("/api/generate-presentation", methods=["POST"])
def gen_presentation():
    data = request.json or {}
    doc_ids = data.get("doc_ids", [])
    title = data.get("title", "Presentation").strip() or "Presentation"
    num_slides = max(3, min(int(data.get("num_slides", 8)), 20))
    style = data.get("style", "professional")

    if not doc_ids:
        return jsonify({"error": "Select at least one document"}), 400

    combined = ""
    for did in doc_ids:
        if did in document_store:
            d = document_store[did]
            combined += f"\n\n=== {d['original_name']} ===\n\n{d['text']}"

    system = (
        "You are a professional presentation designer. "
        "Create compelling, well-structured slides that communicate key ideas clearly."
    )
    prompt = (
        f"Create a {style} presentation titled '{title}' with exactly {num_slides} content slides "
        f"based on:\n\n{trunc(combined, 40000)}\n\n"
        "Return ONLY valid JSON:\n"
        '{"slides":[{"title":"...","content":["bullet1","bullet2","bullet3"]}]}\n'
        "Each slide: 3-5 concise bullet points. Titles must be compelling and actionable."
    )
    try:
        raw = claude(system, prompt)
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError("No JSON found in AI response")
        slides_data = json.loads(m.group())["slides"]

        pid = str(uuid.uuid4())[:8]
        safe = secure_filename(title)[:40]
        fname = f"{pid}_{safe}.pptx"
        build_pptx(title, slides_data, os.path.join(OUTPUT_FOLDER, fname))

        return jsonify({
            "message": "Presentation generated",
            "filename": fname,
            "slides": len(slides_data),
            "download_url": f"/api/download/{fname}",
        })
    except Exception as e:
        logger.error(f"Presentation error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/download/<filename>")
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, secure_filename(filename), as_attachment=True)


@app.route("/api/stats")
def stats():
    return jsonify({
        "total_docs": len(document_store),
        "analyzed": sum(1 for d in document_store.values() if d.get("analysis")),
        "total_words": sum(d.get("word_count", 0) for d in document_store.values()),
        "presentations": len(list(Path(OUTPUT_FOLDER).glob("*.pptx"))),
        "chats": len(chat_sessions),
    })


@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Max 16 MB."}), 413


@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
