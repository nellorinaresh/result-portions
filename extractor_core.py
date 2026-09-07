"""
court_result_extractor.py — core extraction engine.

Implements the rules established through testing against real case files:
  - Content anchor: "doth Order and (Preliminary) Decree as follows" marks the
    start of the Decree/Decretal Order content (used whether or not an explicit
    DECREE heading label follows it).
  - Fallback anchor: "delivered the following" marks the start of the Order/
    Judgment section, used when no Decree transition exists anywhere in the file.
  - JUDGMENT is never used as a content source for Mode 1.
  - Case number, party names, and disposal date are always read from the FIRST
    occurrence (the top caption) — never from the trailing "Decree in ... dt. ..."
    footer, which has been found to carry copy-paste errors in real files.
  - Mode 2 (Final Result Portion in Judgment) pulls the paragraph starting
    "In the result" from within the Judgment/Order-acting-as-judgment section.
    If there's no Judgment at all, the app shows the plain Order text itself,
    labeled using the order's own last line (e.g. "Hence, ... dismissed ...")
    rather than a guessed category.
"""

import re
import os
import sys
import glob
import bisect
import subprocess
import tempfile
import shutil
from datetime import date

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


STOP_PHRASES = ["given under my hand", "pronounced by", "dictated to"]
DECREE_TRANSITION = re.compile(
    r'doth\s+order\s+and\s+(preliminary\s+)?(decree|decretal\s+order)\s+as\s+follows',
    re.IGNORECASE)
ORDER_TRANSITION = re.compile(r'delivered\s+the\s+following', re.IGNORECASE)
IN_THE_RESULT = re.compile(r'\bIn\s+the\s+result\b.*', re.IGNORECASE | re.DOTALL)
HEADING_WORDS = ["DECREE", "DECRETALORDER", "ORDER", "JUDGMENT"]
CASE_NO_PATTERN = re.compile(
    r'([A-Za-z\.]{2,15}\s*No\.?\s*\d+\s*(?:/\s*\d{2,4}|\s+of\s+\d{2,4}))',
    re.IGNORECASE)
MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]


def norm(text):
    return re.sub(r'[^A-Za-z]', '', text).upper()


# ---------------------------------------------------------------------------
# Format conversion
# ---------------------------------------------------------------------------

class LibreOfficeNotFound(Exception):
    """Raised once, upfront, when soffice can't be located anywhere sensible —
    so the app can stop with one clear message instead of a cryptic per-file
    error (this was a real bug found when a Windows machine had LibreOffice
    installed but not on PATH — see convert_to_docx below)."""
    pass


_SOFFICE_PATH_CACHE = None


def find_soffice():
    """Locate the LibreOffice 'soffice' executable. Checks PATH first (covers
    Linux/Cloud Run, where installing LibreOffice normally puts it on PATH),
    then falls back to common install locations on Windows/macOS, since
    installing LibreOffice there does not add it to PATH by default."""
    global _SOFFICE_PATH_CACHE
    if _SOFFICE_PATH_CACHE:
        return _SOFFICE_PATH_CACHE

    candidates = [shutil.which('soffice'), shutil.which('soffice.exe')]

    if sys.platform.startswith('win'):
        for base in [os.environ.get('ProgramFiles', r'C:\Program Files'),
                     os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)')]:
            candidates.append(os.path.join(base, 'LibreOffice', 'program', 'soffice.exe'))
    elif sys.platform == 'darwin':
        candidates.append('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    else:
        candidates += ['/usr/bin/soffice', '/opt/libreoffice/program/soffice',
                        '/snap/bin/libreoffice']

    for c in candidates:
        if c and os.path.isfile(c):
            _SOFFICE_PATH_CACHE = c
            return c

    raise LibreOfficeNotFound(
        "LibreOffice was not found on this server. It's needed to read "
        ".doc and .odt files. Please contact the site administrator."
    )


def convert_to_docx(path, workdir):
    """Convert .doc/.odt to .docx via headless LibreOffice. Returns new path,
    or the original path if it's already .docx."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        return path
    if ext in ('.doc', '.odt'):
        soffice = find_soffice()
        result = subprocess.run(
            [soffice, '--headless', '--convert-to', 'docx', '--outdir', workdir, path],
            capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
        base = os.path.splitext(os.path.basename(path))[0]
        out_path = os.path.join(workdir, base + '.docx')
        if not os.path.exists(out_path):
            raise RuntimeError("Conversion did not produce expected output file")
        return out_path
    raise ValueError(f"Unsupported extension: {ext}")


# ---------------------------------------------------------------------------
# Text extraction (format-specific -> flat paragraph list)
# ---------------------------------------------------------------------------

def get_paragraphs_docx(path):
    doc = DocxDocument(path)
    return [p.text for p in doc.paragraphs]


def get_paragraphs_pdf(path):
    """Extract text page by page, drop lines that repeat identically across
    most pages (running headers/footers, CNR barcodes etc.), return a flat
    paragraph list similar in spirit to the docx paragraph list."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber not installed — required for PDF input")
    pages_lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            pages_lines.append([l for l in text.split('\n') if l.strip()])

    if len(pages_lines) > 2:
        from collections import Counter
        counts = Counter()
        for lines in pages_lines:
            for l in set(lines):
                counts[l.strip()] += 1
        # Only strip a line that repeats on EVERY page — a real, confirmed
        # bug showed a case number appearing as a running header on some
        # (not all) pages getting wrongly stripped at a 60% threshold, which
        # silently deleted the very thing extraction needed. A genuine page
        # header/footer/barcode appears on literally every page; a case
        # number used as a section-boundary running header does not.
        threshold = len(pages_lines)
        repeated = {l for l, c in counts.items() if c >= threshold}
    else:
        repeated = set()

    flat = []
    for lines in pages_lines:
        for l in lines:
            if l.strip() in repeated:
                continue
            flat.append(l)
    # merge into paragraph-ish chunks: blank-separated already not present in PDF
    # extraction, so treat each non-empty line as its own paragraph — the
    # regex-based rules below don't require exact paragraph boundaries.
    return flat


def get_paragraphs(path, workdir):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.pdf':
        return get_paragraphs_pdf(path), 'pdf'
    docx_path = convert_to_docx(path, workdir)
    return get_paragraphs_docx(docx_path), 'docx'


# ---------------------------------------------------------------------------
# Field extraction
# ---------------------------------------------------------------------------

def extract_case_no(paras_text):
    for t in paras_text[:20]:
        m = CASE_NO_PATTERN.search(t)
        if m:
            val = m.group(1)
            val = re.sub(r'\s*/\s*', '/', val)
            return val.strip()
    return None


def extract_case_type(case_no):
    if not case_no:
        return "UNCLASSIFIED"
    m = re.match(r'([A-Za-z\.]+)\s*No\.?', case_no, re.IGNORECASE)
    if not m:
        return "UNCLASSIFIED"
    letters = re.sub(r'[^A-Za-z]', '', m.group(1)).upper()
    return letters


def extract_parties(paras_text):
    try:
        b_idx = next(i for i, t in enumerate(paras_text) if t.strip().lower().startswith('between'))
    except StopIteration:
        return None, None
    try:
        and_idx = next(i for i in range(b_idx + 1, len(paras_text)) if paras_text[i].strip().lower() == 'and')
    except StopIteration:
        return None, None
    try:
        resp_end = next(i for i in range(and_idx + 1, len(paras_text))
                         if re.search(r'(Respondents?|Defendants?)\s*$', paras_text[i].strip(), re.IGNORECASE))
    except StopIteration:
        resp_end = and_idx + 3

    pet_block = [t for t in paras_text[b_idx + 1:and_idx] if t.strip()]
    resp_block = [t for t in paras_text[and_idx + 1:resp_end + 1] if t.strip()]

    def names_from_block(block):
        if len(block) == 1 and block[0].strip().lower().startswith('nil'):
            return ['Nil']
        names = []
        cur = []
        for line in block:
            line_s = line.strip()
            if re.match(r'^\d+\.', line_s) and cur:
                names.append(cur[0])
                cur = []
            cur.append(line_s)
        if cur:
            names.append(cur[0])
        cleaned = []
        for n in names:
            n = re.sub(r'^\d+\.\s*', '', n)
            n = re.split(r',|\bS/o\b|\bD/o\b|\bW/o\b', n)[0].strip()
            n = re.sub(r'\.*\s*(Petitioners?|Respondents?|Plaintiffs?|Defendants?)\s*$', '', n, flags=re.IGNORECASE).strip()
            if n and n.lower() != 'nil':
                cleaned.append(n)
        return cleaned if cleaned else ['Nil']

    return names_from_block(pet_block), names_from_block(resp_block)


def extract_date(paras_text):
    for t in paras_text:
        low = t.lower()
        if 'pronounced by' in low or 'given under my hand' in low or 'dictated to' in low:
            m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+day\s+of\s+(\w+),?\s+(\d{4})', t, re.IGNORECASE)
            if m:
                month_num = MONTH_NAMES.index(m.group(2).capitalize()) if m.group(2).capitalize() in MONTH_NAMES else None
                return {
                    "display": f"{int(m.group(1)):02d}-{month_num:02d}-{m.group(3)}" if month_num else t,
                    "month": month_num,
                    "year": int(m.group(3)),
                }
    return None


def is_stop_para(text):
    low = text.lower()
    return any(p in low for p in STOP_PHRASES)


def find_anchor(paras_text, pattern):
    # Search a flattened, position-mapped version of the text so a match
    # spanning two lines/paragraphs (a real, confirmed PDF line-wrap case —
    # "doth Order and Decree as follows" split across two lines) is still
    # found. The match's END position is mapped back to a line index, so
    # content capture (which starts on the next line) doesn't include the
    # tail of the anchor phrase itself when it was split.
    import bisect
    flat = ''
    line_starts = []
    for t in paras_text:
        line_starts.append(len(flat))
        flat += t + ' '
    m = pattern.search(flat)
    if not m:
        return None
    return bisect.bisect_right(line_starts, m.end()) - 1


def join_into_paragraphs(lines):
    """Merge line-wrap fragments back into flowing paragraphs — needed for
    PDF input, where each visual line is a separate array entry (unlike
    .docx, where each paragraph is already complete). A new paragraph starts
    only at a numbered point (i), ii), 1., a) ...); everything else is a
    continuation of the previous line. For .docx input this is a safe no-op,
    since each entry there is already a complete paragraph."""
    new_point = re.compile(r'^(?:[ivxlcdm]+\)|\d+\)|[a-z]\)|\d+\.)\s', re.IGNORECASE)
    paragraphs = []
    for line in lines:
        if not paragraphs or new_point.match(line):
            paragraphs.append(line)
        else:
            paragraphs[-1] += ' ' + line
    return paragraphs


def capture_content_from(paras_text, start_idx):
    """Starting right after start_idx, skip blanks and one bare heading label,
    then collect lines until a stop phrase or another heading word."""
    i = start_idx + 1
    while i < len(paras_text) and not paras_text[i].strip():
        i += 1
    if i < len(paras_text) and norm(paras_text[i]) in HEADING_WORDS:
        i += 1
        while i < len(paras_text) and not paras_text[i].strip():
            i += 1

    lines = []
    for t in paras_text[i:]:
        t = t.strip()
        if not t:
            continue
        if is_stop_para(t):
            break
        if norm(t) in HEADING_WORDS + ["APPENDIXOFEVIDENCE", "STATEMENTOFCOSTS"]:
            break
        lines.append(t)
    return join_into_paragraphs(lines)


def mode1_decree_result(paras_text):
    """Decreetal Order Result: Decree/Decretal Order transition > Order fallback."""
    idx = find_anchor(paras_text, DECREE_TRANSITION)
    if idx is not None:
        return "Decree", capture_content_from(paras_text, idx)
    idx = find_anchor(paras_text, ORDER_TRANSITION)
    if idx is not None:
        return "Order", capture_content_from(paras_text, idx)
    return None, None


def mode2_final_result(paras_text):
    """Final Result Portion in Judgment: 'In the result' paragraph inside the
    Judgment/Order-acting-as-judgment section. Falls back to the plain Order
    text (labeled with its own last line) when there's no Judgment at all."""
    order_idx = find_anchor(paras_text, ORDER_TRANSITION)
    decree_idx = find_anchor(paras_text, DECREE_TRANSITION)

    if order_idx is None:
        return "Not applicable — no Order/Judgment section found", []

    end = decree_idx if decree_idx is not None else len(paras_text)
    window = paras_text[order_idx + 1:end]  # stop-phrase lines kept IN here — they're the boundary marker

    # Find where "in the result" starts (tolerating a match split across two
    # lines — the same real scenario as the Decree anchor), then capture
    # forward from there using the same bounded logic as capture_content_from,
    # stopping at the next real stop-phrase/heading rather than running to
    # the end of the whole window (a real bug caught by testing: without a
    # stop boundary, this used to swallow the appendix, second case caption,
    # and everything up to the next Decree section).
    flat, line_starts = '', []
    for t in window:
        line_starts.append(len(flat))
        flat += t + ' '
    m = re.search(r'\bin\s+the\s+result\b', flat, re.IGNORECASE)
    if m:
        start_line = bisect.bisect_right(line_starts, m.start()) - 1
        lines = []
        for t in window[start_line:]:
            t = t.strip()
            if not t:
                continue
            if is_stop_para(t):
                break
            if norm(t) in HEADING_WORDS + ["APPENDIXOFEVIDENCE", "STATEMENTOFCOSTS"]:
                break
            lines.append(t)
        combined = ' '.join(lines)
        m2 = re.search(r'\bin\s+the\s+result\b.*', combined, re.IGNORECASE | re.DOTALL)
        if m2:
            return 'Judgment — paragraph beginning "In the result"', [m2.group(0).strip()]

    # no "In the result" paragraph -> this category never had a Judgment;
    # show the plain order text, label = the order's own last line
    order_content = capture_content_from(paras_text, order_idx)
    if not order_content:
        return "Not applicable", []
    last_line = order_content[-1]
    m = re.search(r'([^.]*\.)\s*$', last_line)
    label_line = m.group(1).strip() if m else last_line
    return f'No Judgment — Order shown; last line: "{label_line}"', order_content


def process_file(path, mode, workdir):
    try:
        paras_text, fmt = get_paragraphs(path, workdir)
    except Exception as e:
        return {"file": os.path.basename(path), "error": f"could not read file ({e})"}

    case_no = extract_case_no(paras_text)
    if not case_no:
        return {"file": os.path.basename(path), "error": "no case number found"}

    petitioners, respondents = extract_parties(paras_text)
    date_info = extract_date(paras_text)
    case_type = extract_case_type(case_no)

    if mode == 1:
        source, content = mode1_decree_result(paras_text)
    else:
        source, content = mode2_final_result(paras_text)

    if content is None:
        return {"file": os.path.basename(path), "case_no": case_no,
                "error": "no recognizable Order/Decree/Judgment structure found"}

    return {
        "file": os.path.basename(path),
        "case_no": case_no,
        "case_type": case_type,
        "petitioners": petitioners or ["Nil"],
        "respondents": respondents or ["Nil"],
        "date_info": date_info,
        "source": source,
        "content": content,
    }
