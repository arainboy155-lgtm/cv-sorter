"""
cv_engine.py - CV Sorter ka "dimagh" (core logic)
-------------------------------------------------
- CV file (PDF / DOCX / TXT) se text nikalta hai
- Keywords ke zariye detect karta hai ke CV kis subject ke teacher ki hai
- Candidate ka naam, email, phone, experience nikalta hai
- Subject ke hisaab se folder bana kar CV copy/move karta hai
- Duplicate CVs pakadta hai (file hash se)
- Har CV ka record SQLite database mein rakhta hai, Excel/CSV report bana sakta hai
"""

import csv
import hashlib
import json
import os
import re
import shutil
import sqlite3
from datetime import datetime

APP_DIR = os.path.join(os.path.expanduser("~"), "CVSorter")
CONFIG_PATH = os.path.join(APP_DIR, "settings.json")
DB_PATH = os.path.join(APP_DIR, "cv_records.db")

SUPPORTED_EXT = (".pdf", ".docx", ".txt")
UNSORTED = "Unsorted - Manual Review"
DUPLICATES = "Duplicates"

# ---------------------------------------------------------------------------
# Default subjects aur keywords. Strong keywords ka weight zyada hai.
# User Settings se inhein badal / naye subject add kar sakta hai.
# ---------------------------------------------------------------------------
DEFAULT_SUBJECTS = {
    "English Teacher": {
        "strong": ["english teacher", "english lecturer", "teacher of english", "esl", "efl",
                   "tesol", "tefl", "celta", "ma english", "m.a english", "m.a. english",
                   "bs english", "english literature", "english language teaching", "elt"],
        "weak": ["grammar", "literature", "linguistics", "phonetics", "essay writing",
                 "creative writing", "spoken english", "ielts", "toefl", "poetry", "novel",
                 "reading comprehension", "vocabulary"],
    },
    "Math Teacher": {
        "strong": ["math teacher", "maths teacher", "mathematics teacher", "mathematics lecturer",
                   "teacher of mathematics", "msc mathematics", "m.sc mathematics", "m.sc. mathematics",
                   "bs mathematics", "mphil mathematics", "bs math", "msc math"],
        "weak": ["algebra", "calculus", "geometry", "trigonometry", "statistics", "arithmetic",
                 "differential equations", "linear algebra", "probability", "mathematical",
                 "numerical analysis", "mathematics", "maths"],
    },
    "Physics Teacher": {
        "strong": ["physics teacher", "physics lecturer", "teacher of physics", "msc physics",
                   "m.sc physics", "bs physics", "mphil physics"],
        "weak": ["mechanics", "thermodynamics", "electromagnetism", "optics", "quantum",
                 "physics lab", "physics"],
    },
    "Chemistry Teacher": {
        "strong": ["chemistry teacher", "chemistry lecturer", "teacher of chemistry",
                   "msc chemistry", "m.sc chemistry", "bs chemistry", "mphil chemistry"],
        "weak": ["organic chemistry", "inorganic", "biochemistry", "titration",
                 "chemistry lab", "periodic table", "chemistry"],
    },
    "Biology Teacher": {
        "strong": ["biology teacher", "biology lecturer", "teacher of biology", "msc biology",
                   "msc zoology", "msc botany", "bs biology", "bs zoology", "bs botany"],
        "weak": ["zoology", "botany", "genetics", "microbiology", "anatomy", "ecology",
                 "cell biology", "biology"],
    },
    "Computer Science Teacher": {
        "strong": ["computer teacher", "computer science teacher", "computer science lecturer",
                   "it teacher", "ict teacher", "bscs", "bs computer science", "mcs", "bsit",
                   "bs information technology"],
        "weak": ["programming", "python", "java", "c++", "html", "database", "ms office",
                 "networking", "computer science", "coding", "software"],
    },
    "Urdu Teacher": {
        "strong": ["urdu teacher", "urdu lecturer", "teacher of urdu", "ma urdu", "m.a urdu",
                   "m.a. urdu", "bs urdu"],
        "weak": ["urdu literature", "urdu adab", "ghazal", "nazm", "iqbaliyat", "urdu"],
    },
    "Islamiat Teacher": {
        "strong": ["islamiat teacher", "islamic studies teacher", "teacher of islamiat",
                   "ma islamiat", "m.a islamiat", "ma islamic studies", "bs islamic studies",
                   "nazra teacher", "quran teacher"],
        "weak": ["islamiat", "islamic studies", "hadith", "fiqh", "tafseer", "seerah",
                 "nazra", "tajweed", "hifz", "arabic"],
    },
    "Science Teacher (General)": {
        "strong": ["science teacher", "general science teacher", "teacher of science"],
        "weak": ["general science", "science fair", "science lab"],
    },
    "Social Studies Teacher": {
        "strong": ["social studies teacher", "pak studies teacher", "history teacher",
                   "geography teacher", "ma history", "ma political science", "ma geography"],
        "weak": ["pakistan studies", "pak studies", "history", "geography", "civics",
                 "social studies", "political science"],
    },
    "Primary / Montessori Teacher": {
        "strong": ["montessori teacher", "primary teacher", "pre-school teacher",
                   "preschool teacher", "early childhood", "nursery teacher", "ecce",
                   "montessori diploma", "kindergarten teacher"],
        "weak": ["montessori", "kindergarten", "playgroup", "phonics", "primary classes",
                 "early years"],
    },
}

DEFAULT_CONFIG = {
    "output_folder": os.path.join(os.path.expanduser("~"), "Documents", "Sorted CVs"),
    "watch_folder": "",
    "mode": "copy",              # "copy" ya "move"
    "min_score": 4,              # is se kam score ho to Unsorted mein jaye
    "rename_files": True,        # file ko "Candidate Name - Subject.pdf" naam do
    "subjects": DEFAULT_SUBJECTS,
}

STRONG_WEIGHT = 5
WEAK_WEIGHT = 1


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
def load_config():
    os.makedirs(APP_DIR, exist_ok=True)
    if not os.path.exists(CONFIG_PATH):
        save_config(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, json.JSONDecodeError):
        cfg = {}
    merged = json.loads(json.dumps(DEFAULT_CONFIG))
    merged.update(cfg)
    return merged


def save_config(cfg):
    os.makedirs(APP_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------
def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    text = ""
    if ext == ".pdf":
        text = _pdf_text(path)
    elif ext == ".docx":
        text = _docx_text(path)
    elif ext == ".txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    return text


def _pdf_text(path):
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        reader = PdfReader(path)
        return "\n".join((p.extract_text() or "") for p in reader.pages)
    except Exception:
        return ""


def _docx_text(path):
    try:
        import docx
    except ImportError:
        return ""
    try:
        d = docx.Document(path)
        parts = [p.text for p in d.paragraphs]
        for table in d.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(parts)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------
def _count(keyword, text):
    pattern = r"(?<![a-z0-9])" + re.escape(keyword.lower()) + r"(?![a-z0-9])"
    return len(re.findall(pattern, text))


def classify(text, subjects, min_score=4):
    """Return (best_subject, confidence_percent, scores_dict)."""
    t = re.sub(r"\s+", " ", text.lower())
    t = t.replace("m. sc", "m.sc").replace("m. a", "m.a")
    scores = {}
    for subject, kw in subjects.items():
        score = 0
        for k in kw.get("strong", []):
            score += min(_count(k, t), 3) * STRONG_WEIGHT
        for k in kw.get("weak", []):
            score += min(_count(k, t), 3) * WEAK_WEIGHT
        scores[subject] = score

    if not scores:
        return UNSORTED, 0, scores
    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    best, best_score = ranked[0]
    second_score = ranked[1][1] if len(ranked) > 1 else 0
    total = sum(s for s in scores.values()) or 1

    if best_score < min_score:
        return UNSORTED, 0, scores
    confidence = round(100 * best_score / total)
    # Agar do subjects bilkul barabar hain to insaan ko check karna chahiye
    if best_score == second_score:
        return UNSORTED, confidence, scores
    return best, confidence, scores


# ---------------------------------------------------------------------------
# Candidate details
# ---------------------------------------------------------------------------
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+92|0092|0)[\s-]?3\d{2}[\s-]?\d{7}|(?:\+?\d[\d\s-]{9,14}\d)")
EXP_RE = re.compile(r"(\d{1,2})\s*\+?\s*(?:years?|yrs?)\s*(?:of)?\s*(?:teaching\s*)?(?:experience|exp)", re.I)

SKIP_NAME_WORDS = {"curriculum", "vitae", "resume", "cv", "profile", "contact", "objective",
                   "personal", "information", "details", "address", "email", "phone", "mobile"}


def extract_details(text):
    email_m = EMAIL_RE.search(text)
    email = email_m.group(0) if email_m else ""
    phone_m = PHONE_RE.search(text)
    phone = re.sub(r"[\s-]", "", phone_m.group(0)) if phone_m else ""
    exp_matches = [int(x) for x in EXP_RE.findall(text)]
    experience = max(exp_matches) if exp_matches else ""

    name = ""
    m = re.search(r"name\s*[:\-]\s*([A-Za-z][A-Za-z .]{2,40})", text, re.I)
    if m:
        name = m.group(1).strip()
    else:
        for line in text.splitlines()[:8]:
            line = line.strip()
            words = line.split()
            if not (2 <= len(words) <= 4):
                continue
            if any(w.lower().strip(":") in SKIP_NAME_WORDS for w in words):
                continue
            if "@" in line or any(c.isdigit() for c in line):
                continue
            if all(w[0].isalpha() for w in words):
                name = line
                break
    name = re.split(r"\s{2,}|\n", name)[0].strip().title() if name else ""
    return {"name": name, "email": email, "phone": phone, "experience": experience}


# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------
def _db():
    os.makedirs(APP_DIR, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("""CREATE TABLE IF NOT EXISTS cvs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT, original_name TEXT, saved_path TEXT, subject TEXT,
        confidence INTEGER, name TEXT, email TEXT, phone TEXT,
        experience TEXT, file_hash TEXT UNIQUE)""")
    return con


def find_by_hash(file_hash):
    con = _db()
    row = con.execute("SELECT saved_path FROM cvs WHERE file_hash=?", (file_hash,)).fetchone()
    con.close()
    return row[0] if row else None


def all_records(search=""):
    con = _db()
    q = ("SELECT id, date, name, subject, confidence, email, phone, experience, "
         "original_name, saved_path FROM cvs")
    args = ()
    if search:
        like = f"%{search}%"
        q += " WHERE name LIKE ? OR subject LIKE ? OR email LIKE ? OR phone LIKE ? OR original_name LIKE ?"
        args = (like,) * 5
    q += " ORDER BY id DESC"
    rows = con.execute(q, args).fetchall()
    con.close()
    return rows


def update_subject(record_id, new_subject, output_folder):
    """Galat detection ho to user manually subject badal sakta hai; file bhi move hoti hai."""
    con = _db()
    row = con.execute("SELECT saved_path FROM cvs WHERE id=?", (record_id,)).fetchone()
    if not row:
        con.close()
        return None
    old_path = row[0]
    new_dir = os.path.join(output_folder, safe_name(new_subject))
    os.makedirs(new_dir, exist_ok=True)
    new_path = unique_path(os.path.join(new_dir, os.path.basename(old_path)))
    if os.path.exists(old_path):
        shutil.move(old_path, new_path)
    con.execute("UPDATE cvs SET subject=?, saved_path=?, confidence=100 WHERE id=?",
                (new_subject, new_path, record_id))
    con.commit()
    con.close()
    return new_path


def delete_record(record_id):
    con = _db()
    con.execute("DELETE FROM cvs WHERE id=?", (record_id,))
    con.commit()
    con.close()


def export_report(path):
    rows = all_records()
    headers = ["ID", "Date", "Candidate Name", "Subject", "Confidence %", "Email",
               "Phone", "Experience (Years)", "Original File", "Saved At"]
    if path.lower().endswith(".xlsx"):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill
            wb = Workbook()
            ws = wb.active
            ws.title = "CVs"
            ws.append(headers)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E78")
            for r in rows:
                ws.append(list(r))
            widths = [6, 18, 24, 26, 12, 28, 16, 12, 30, 60]
            for i, w in enumerate(widths):
                ws.column_dimensions[chr(65 + i)].width = w
            ws.freeze_panes = "A2"
            # Summary sheet
            s = wb.create_sheet("Summary")
            s.append(["Subject", "Total CVs"])
            for c in s[1]:
                c.font = Font(bold=True)
            for subj, n in subject_counts().items():
                s.append([subj, n])
            s.column_dimensions["A"].width = 32
            wb.save(path)
            return path
        except ImportError:
            path = os.path.splitext(path)[0] + ".csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)
    return path


def subject_counts():
    con = _db()
    rows = con.execute("SELECT subject, COUNT(*) FROM cvs GROUP BY subject ORDER BY 2 DESC").fetchall()
    con.close()
    return dict(rows)


# ---------------------------------------------------------------------------
# File helpers
# ---------------------------------------------------------------------------
def file_hash(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def safe_name(s):
    return re.sub(r'[<>:"/\\|?*]', "", s).strip() or "Unknown"


def unique_path(path):
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(path):
        path = f"{base} ({i}){ext}"
        i += 1
    return path


# ---------------------------------------------------------------------------
# Main entry: aik CV process karo
# ---------------------------------------------------------------------------
def process_cv(path, cfg):
    """Returns dict with result info. status: 'sorted' | 'unsorted' | 'duplicate' | 'error'."""
    result = {"file": os.path.basename(path), "status": "error", "subject": "",
              "confidence": 0, "saved_path": "", "message": ""}
    try:
        if not path.lower().endswith(SUPPORTED_EXT):
            result["message"] = "Ye file type support nahi hai (sirf PDF, DOCX, TXT)"
            return result

        out_root = cfg["output_folder"]
        os.makedirs(out_root, exist_ok=True)
        h = file_hash(path)
        existing = find_by_hash(h)
        if existing:
            result.update(status="duplicate", subject=DUPLICATES, saved_path=existing,
                          message="Ye CV pehle se save hai")
            return result

        text = extract_text(path)
        if len(text.strip()) < 30:
            subject, conf, details = UNSORTED, 0, {"name": "", "email": "", "phone": "", "experience": ""}
            result["message"] = "Text nahi mila (shayad scanned image PDF hai)"
        else:
            subject, conf, _ = classify(text, cfg["subjects"], cfg.get("min_score", 4))
            details = extract_details(text)

        dest_dir = os.path.join(out_root, safe_name(subject))
        os.makedirs(dest_dir, exist_ok=True)
        ext = os.path.splitext(path)[1].lower()
        if cfg.get("rename_files") and details["name"] and subject != UNSORTED:
            fname = safe_name(f"{details['name']} - {subject}") + ext
        else:
            fname = os.path.basename(path)
        dest = unique_path(os.path.join(dest_dir, fname))

        if cfg.get("mode") == "move":
            shutil.move(path, dest)
        else:
            shutil.copy2(path, dest)

        con = _db()
        con.execute("""INSERT INTO cvs (date, original_name, saved_path, subject, confidence,
                       name, email, phone, experience, file_hash) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (datetime.now().strftime("%Y-%m-%d %H:%M"), os.path.basename(path), dest,
                     subject, conf, details["name"], details["email"], details["phone"],
                     str(details["experience"]), h))
        con.commit()
        con.close()

        result.update(status="unsorted" if subject == UNSORTED else "sorted",
                      subject=subject, confidence=conf, saved_path=dest, **details)
        if not result["message"]:
            result["message"] = "OK"
        return result
    except Exception as e:  # koi bhi error app ko crash na kare
        result["message"] = str(e)
        return result
