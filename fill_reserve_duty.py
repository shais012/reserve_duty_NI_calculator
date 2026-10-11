#!/usr/bin/env python3
"""Fill the reserve-duty (ימ"מ) workbook from Hilan pay-slip PDFs.

Per pay-slip it writes (into the sheet of the slip's year, column of its month):
    row 9   ברוטו לב.לאומי             (נתונים נוספים)
    row 11  sum of code 198            (פרוט תשלומים)
    row 12  sum of code 193            (פרוט תשלומים)
    helper sheet, col E   ברוטו לב.ל ב.ק. (נתונים נוספים), 0 if absent
Reserve days (row 27) come from the service certificate (form 3010), the value NI paid per day and the days it paid for
(rows 22/23) from the NI annual letter, and the current day value (F36) from --current-day-value; otherwise they are manual.

Usage:
    python fill_reserve_duty.py slip1.pdf [slip2.pdf ...] [--out filled.xlsx]
    python fill_reserve_duty.py slip.pdf --debug        # no Excel, masked structure dump
"""
import argparse
import datetime as dt
import importlib.util
import json
import os
import re
import sys
import threading
from pathlib import Path

HERE = Path(__file__).resolve().parent
HEB = re.compile(r"[א-ת]")
NUM = re.compile(r"^\(?-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?-?$|^\(?-?\d+(?:\.\d+)?\)?-?$")


# ----------------------------------------------------------------- text helpers
def fix_encoding(s):
    """Some Hilan PDFs store Hebrew as Windows-1255 bytes shown as Latin-1 (ï ð ø ...).
    Map U+00E0..U+00FA back to Hebrew letters."""
    return "".join(
        bytes([ord(c)]).decode("cp1255") if "à" <= c <= "ú" else c for c in s
    )


def canon(s):
    """Hebrew letters only - ignores spaces, dots, quotes, digits."""
    return "".join(HEB.findall(s))


def to_number(tok, rev=False):
    t = tok.strip().replace("\u20aa", "").strip()        # "₪1,234.00" or "1,234.00₪"
    if rev:
        t = t[::-1]
    if not NUM.match(t):
        return None
    neg = t.endswith("-") or t.startswith("-") or t.startswith("(")
    v = float(t.strip("()-").replace(",", ""))
    return -v if neg else v


def shape(tok):
    """Masked shape of a token for --debug output (no real digits)."""
    if HEB.search(tok):
        return "H%d" % len(canon(tok))
    return re.sub(r"\d", "#", tok)


def variants(words):
    """Yield (name, [(word, text)]) for the ways Hebrew text might be stored."""
    desc = sorted(words, key=lambda w: -w["x0"])
    for oname, seq in (("rtl", desc), ("ltr", desc[::-1])):
        for flip in (False, True):
            yield "%s/%s" % (oname, "flip" if flip else "as-is"), [
                (w, w["text"][::-1] if flip else w["text"]) for w in seq
            ]


# --------------------------------------------------------------------- PDF I/O
_PDF_LOCK = threading.RLock()        # PyMuPDF must not be used from two threads at the same time


def _load_words_pymupdf(pdf_path):
    import pymupdf

    pages = []
    with _PDF_LOCK, pymupdf.open(pdf_path) as doc:
        if doc.needs_pass:
            raise PdfReadError(PASSWORD_MESSAGE)
        for page in doc:
            words = [
                {"text": w[4], "x0": w[0], "x1": w[2], "top": w[1]}
                for w in page.get_text("words")
                if w[4].strip()
            ]
            pages.append({"width": float(page.rect.width), "words": words})
    return pages


def _load_words_pdfplumber(pdf_path):
    import pdfplumber

    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
            pages.append({"width": float(page.width), "words": words})
    return pages


def inside_program_folder(path):
    """True if `path` is in (or under) the folder this program lives in."""
    try:
        Path(path).resolve().relative_to(HERE)
        return True
    except ValueError:
        return False


PASSWORD_MESSAGE = ("The PDF is password-protected. Open it, and save or print it to a new PDF without a password "
                    "(the copy stays on your computer), then add the copy.")


class PdfReadError(RuntimeError):
    """A PDF could not be read (missing libraries, or no text layer)."""


def _has_module(name):
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def missing_modules(need_excel=True):
    """Libraries that are not installed in THIS Python. One PDF reader is enough (either one);
    openpyxl is needed to write Excel."""
    missing = []
    if not (_has_module("pymupdf") or _has_module("pdfplumber")):
        missing += ["pymupdf", "pdfplumber"]
    if need_excel and not _has_module("openpyxl"):
        missing.append("openpyxl")
    return missing


def console_python():
    """The Python that is running now, as python.exe even when started via pythonw.exe (so output can be captured)."""
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe"):
        candidate = exe[: -len("pythonw.exe")] + "python.exe"
        if os.path.exists(candidate):
            exe = candidate
    return exe


def install_command():
    """How to install the libraries: install.bat (a private environment next to the program) when it is there,
    otherwise the pip command for the Python that is running now."""
    if (HERE / "install.bat").is_file():
        return '"%s"' % (HERE / "install.bat")
    return '"%s" -m pip install pymupdf pdfplumber openpyxl' % console_python()


RELAUNCH_ENV = "RESERVE_DUTY_RELAUNCHED"


def local_env_python(windowed=False):
    """The Python of a private environment next to the program, if there is one: .venv made by install.bat, or the
    environment the program folder sits in (../Scripts). None if there is none, or it is the one running now."""
    exe = ("pythonw.exe" if windowed else "python.exe") if os.name == "nt" else "python"
    sub = "Scripts" if os.name == "nt" else "bin"
    for base in (HERE / ".venv", HERE.parent):
        candidate = base / sub / exe
        if candidate.is_file():
            try:
                if candidate.resolve() == Path(sys.executable).resolve():
                    return None
            except OSError:
                pass
            return candidate
    return None


def relaunch_with_local_env(script, argv, windowed=False, wait=False):
    """If the libraries are missing here but a private environment exists next to the program, run `script` again with
    that environment's Python and return True (the caller then exits; with `wait` the child's exit code is returned
    instead). Never loops: the child is marked, and is not relaunched again."""
    import subprocess

    if os.environ.get(RELAUNCH_ENV) or not missing_modules(need_excel=True):
        return False
    python = local_env_python(windowed)
    if python is None:
        return False
    env = dict(os.environ, **{RELAUNCH_ENV: "1"})
    cmd = [str(python), str(script)] + list(argv)
    if wait:
        return subprocess.call(cmd, env=env)  # noqa: S603 - our own script, with the Python next to it
    subprocess.Popen(cmd, env=env, creationflags=0x08000000 if os.name == "nt" else 0)  # noqa: S603
    return True


def load_words(pdf_path):
    """Words with positions, per page. PyMuPDF first (decodes more font types);
    pdfplumber if PyMuPDF is missing or finds no text at all."""
    pages, notes = [], []
    try:
        pages = _load_words_pymupdf(pdf_path)
    except ImportError:
        notes.append("pymupdf is not installed")
    if not any(p["words"] for p in pages):
        try:
            pages = _load_words_pdfplumber(pdf_path)
        except ImportError:
            notes.append("pdfplumber is not installed")
        except Exception as e:                                  # pdfplumber's own error for an encrypted file
            if "password" in type(e).__name__.lower() or "password" in str(e).lower():
                raise PdfReadError(PASSWORD_MESSAGE)
            raise
    if not any(p["words"] for p in pages):
        msg = "No text could be read from this PDF."
        if notes:
            msg += " (%s in %s). Install with: %s." % ("; ".join(notes), sys.executable, install_command())
        msg += " If the libraries are installed, the PDF may be a scan or an image with no text layer."
        raise PdfReadError(msg)
    for p in pages:
        for w in p["words"]:
            w["text"] = fix_encoding(w["text"])
    return pages


# ------------------------------------------------ image-only certificates: OCR (optional add-on)
OCR_ZOOM = 1.6                      # render scale; 1.5-2.0 all read every row of a real certificate identically
OCR_WORKERS = 4                     # pages read at the same time
_OCR_ENGINE = None
_OCR_LOCK = threading.Lock()


def ocr_available():
    """The optional OCR add-on (rapidocr + onnxruntime + numpy) and PyMuPDF (to render the pages) are installed."""
    return all(_has_module(m) for m in ("rapidocr", "onnxruntime", "numpy", "pymupdf"))


def ocr_install_command():
    return '"%s" -m pip install rapidocr onnxruntime' % console_python()


def _ocr_engine():
    global _OCR_ENGINE
    with _OCR_LOCK:
        if _OCR_ENGINE is None:
            import logging

            logging.getLogger("RapidOCR").setLevel(logging.ERROR)       # no model-loading chatter in the console
            from rapidocr import RapidOCR

            _OCR_ENGINE = RapidOCR(params={"Global.use_cls": False})    # page-angle step is not needed (pages are upright)
        return _OCR_ENGINE


def _table_band(image):
    """Rows (top, bottom) of the certificate table: from its black header bar down 55% of the page. None if no bar."""
    dark = (image.mean(axis=2) < 90).mean(axis=1)             # share of dark pixels in every image row
    bar = [i for i, v in enumerate(dark) if v > 0.55]
    if len(bar) < 6:
        return None
    height = image.shape[0]
    return max(0, bar[0] - 6), min(height, bar[0] + int(0.55 * height))


def _ocr_read(engine, image, zoom, y_offset=0):
    result = engine(image)
    words = []
    if result is not None and result.boxes is not None:
        for box, text in zip(result.boxes, result.txts):
            xs, ys = [pt[0] for pt in box], [pt[1] + y_offset for pt in box]
            x0, x1, top = min(xs) / zoom, max(xs) / zoom, min(ys) / zoom
            tokens = text.split()
            total = sum(len(t) for t in tokens) or 1
            cursor = x0
            for token in tokens:                                      # share the box between its words
                width = (x1 - x0) * len(token) / total
                words.append({"text": token, "x0": cursor, "x1": cursor + width, "top": top})
                cursor += width
    return words


def _ocr_page_words(engine, image, zoom):
    """Read only the table when its header bar is found (about 35% faster); read the whole page if that fails,
    or if the cut leaves less than one row of room below the last line (the table might continue)."""
    band = _table_band(image)
    if band:
        words = _ocr_read(engine, image[band[0]:band[1]], zoom, band[0])
        tops = sorted({round(w["top"] * zoom / 6) for w in words})       # one value per text line (6px buckets)
        if len(tops) >= 2:
            pitch = 6 * (tops[-1] - tops[0]) / (len(tops) - 1)
            if band[1] - tops[-1] * 6 > 1.5 * pitch:
                return words
    return _ocr_read(engine, image, zoom)


def ocr_words(pdf_path, zoom=None, workers=None):
    """Read an image-only PDF: render every page, then recognise the pages in parallel. Returns pages like
    load_words, marked "ocr": True. Words get positions in PDF points, so the normal parsing works unchanged."""
    from concurrent.futures import ThreadPoolExecutor

    import numpy as np
    import pymupdf

    zoom = zoom or OCR_ZOOM
    images, widths = [], []
    with _PDF_LOCK, pymupdf.open(str(pdf_path)) as doc:                 # rendering is quick; do it under the lock
        for page in doc:
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            images.append(np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n).copy())
            widths.append(float(page.rect.width))
    engine = _ocr_engine()
    with ThreadPoolExecutor(max(1, min(workers or OCR_WORKERS, len(images)))) as pool:
        all_words = list(pool.map(lambda im: _ocr_page_words(engine, im, zoom), images))   # keeps page order
    return [{"width": w, "words": ws, "ocr": True} for w, ws in zip(widths, all_words)]


def load_pages(pdf_path, cfg, ocr_allowed=False):
    """Words per page. A PDF without any text is read by OCR when `ocr_allowed` (used for the reserve-service
    certificate only: it is mostly digits, which OCR reads well; Hebrew pay-slips are not read this way)."""
    try:
        return load_words(pdf_path)
    except PdfReadError as err:
        if not ocr_allowed or not _has_module("pymupdf") or str(err) == PASSWORD_MESSAGE:
            raise
        if not ocr_available():
            raise PdfReadError("%s If this is the reserve-service certificate saved as a picture, install the OCR add-on to read it: %s"
                               % (err, ocr_install_command()))
        pages = ocr_words(pdf_path)
        tol = cfg["parsing"]["line_tolerance"]
        if not any(parse_form_rows(group_lines(pg["words"], tol), 0) or parse_letter_rows(group_lines(pg["words"], tol), 0)
                   for pg in pages):
            raise PdfReadError("The PDF is a picture, and image recognition (OCR) found no reserve-service table in it.")
        return pages


def group_lines(words, tol):
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(lines[-1]["top"] - w["top"]) <= tol:
            lines[-1]["words"].append(w)
            n = len(lines[-1]["words"])
            lines[-1]["top"] += (w["top"] - lines[-1]["top"]) / n
        else:
            lines.append({"top": w["top"], "words": [w]})
    return lines


# --------------------------------------------------------------------- parsing
def find_window(line, key_canon, max_words=8, with_next=False):
    """Find consecutive words whose Hebrew letters concatenate exactly to key.
    Pure punctuation fragments (e.g. a lone '.' that a PDF reader split off) are skipped
    and don't count toward max_words; a number in between breaks the match.
    with_next: also return the Hebrew letters of the word right after the match ("" if a number or nothing follows),
    so "ברוטו לב"ל" can be told apart from "ברוטו לב"ל ב.ק."."""
    for vname, seq in variants(line["words"]):
        for i in range(len(seq)):
            acc, used = "", 0
            for j in range(i, len(seq)):
                text = seq[j][1]
                c = canon(text)
                if not c:
                    if any(ch.isdigit() for ch in text) or used == 0:
                        break          # a number, or a punctuation start, ends/forbids the window
                    continue           # punctuation fragment inside the window
                acc += c
                used += 1
                if acc == key_canon:
                    win = [w for w, _ in seq[i:j + 1]]
                    if not with_next:
                        return vname, win
                    nxt = ""
                    for _, t in seq[j + 1:]:
                        if any(ch.isdigit() for ch in t):
                            break
                        if canon(t):
                            nxt = canon(t)
                            break
                    return vname, win, nxt
                if not key_canon.startswith(acc) or used >= max_words:
                    break
    return None


MONTH_YEAR_RE = re.compile(r"(0?[1-9]|1[0-2])[/.\-](20\d\d|\d\d)")


def _title_keys(cfg):
    return [canon(k) for k in (cfg.get("title_keys") or [cfg["title_key"]])]


def _month_year_in(words, months):
    """Month and year in a line: a month name ("מרץ") and a 4-digit year, or a numeric "03/2026" / "3.26"."""
    month = year = None
    rev_digits = False
    for w in words:
        t = w["text"].strip(" :")
        # month and year may be glued into one token (e.g. "2026" + Hebrew month)
        letters = canon(t)
        for cand in (letters, letters[::-1]):
            if cand in months and month is None:
                month = months[cand]
        m = re.search(r"(?<!\d)20\d\d(?!\d)", t)
        if m and not MONTH_YEAR_RE.fullmatch(t[::-1]):
            year = int(m.group())
        else:
            m = re.search(r"(?<!\d)\d\d02(?!\d)", t)
            if m and not MONTH_YEAR_RE.fullmatch(t):
                year, rev_digits = int(m.group()[::-1]), True
    if month is None:
        for w in words:
            t = w["text"].strip(" :")
            for cand, rev in ((t, False), (t[::-1], True)):
                m = MONTH_YEAR_RE.fullmatch(cand)
                if m:
                    y = int(m.group(2))
                    return int(m.group(1)), (y if y > 99 else 2000 + y), rev
    return month, year, rev_digits


def find_title(lines, cfg):
    """(month, year, digits_reversed) from the slip title, e.g. "תלוש שכר לחודש מרץ 2026" or "תלוש משכורת 03/2026".
    The wordings are listed in config.json ("title_keys"); month and year may also sit on the line below the title."""
    keys = _title_keys(cfg)
    months = {canon(k): v for k, v in cfg["months"].items()}
    tops = [l["top"] for l in lines]
    top_zone = min(tops) + 0.35 * (max(tops) - min(tops)) if tops else 0     # the title is in the top part of the page,
    title_seen = False
    for i, line in enumerate(lines):                                         # never a retro-pay row further down
        if line["top"] > top_zone or not any(canon(w["text"]) for w in line["words"]):
            continue
        hit = False
        for _, seq in variants(line["words"]):
            joined = "".join(canon(t) for _, t in seq)
            if any(k in joined for k in keys):
                hit = True
                break
        if not hit:
            continue
        month, year, rev_digits = _month_year_in(line["words"], months)
        if not (month and year) and i + 1 < len(lines):
            m2, y2, r2 = _month_year_in(lines[i + 1]["words"], months)
            month, year, rev_digits = month or m2, year or y2, rev_digits or r2
        if month and year:
            return month, year, rev_digits
        title_seen = True
    # e.g. "משכורת לחודש" in one box and "אוקטובר 2014" in a corner, or a slip whose title wording is not listed but
    # whose header says "דצמבר 2024": a month NAME and a year in the header area (a numeric month only with a title,
    # since the header also holds dates like a start date)
    for line in lines:
        if line["top"] <= top_zone:
            month, year, rev_digits = _month_year_in(line["words"], months)
            named = any(c in months or c[::-1] in months for c in (canon(w["text"].strip(" :")) for w in line["words"]))
            if month and year and (named or title_seen):
                return month, year, rev_digits
    return None


def find_value_for_label(line, key_canon, rev_digits, not_followed_by=()):
    """Number sitting next to a label (in RTL tables the value is left of it). A match whose next word starts with one
    of `not_followed_by` (Hebrew letters only) is a different item - e.g. "ברוטו לב"ל ב.ק." for "ברוטו לב"ל" - and is skipped."""
    hit = find_window(line, key_canon, with_next=True)
    if not hit:
        return None
    vname, win, nxt = hit
    if nxt and any(nxt.startswith(x) for x in not_followed_by):
        return None
    label_x0 = min(w["x0"] for w in win)
    label_ids = {id(w) for w in win}
    nums = [(w, to_number(w["text"], rev_digits)) for w in line["words"] if id(w) not in label_ids]
    nums = [(w, v) for w, v in nums if v is not None]
    left = [(w, v) for w, v in nums if w["x1"] <= label_x0 + 1]
    if left:
        return vname, max(left, key=lambda p: p[0]["x0"])[1], "left"
    right = [(w, v) for w, v in nums if w["x0"] >= max(x["x1"] for x in win) - 1]
    if right:
        return vname, min(right, key=lambda p: p[0]["x0"])[1], "right"
    return vname, None, None


HEADER_CODE_TOL = 30.0      # points: a code under the code-column header
HEADER_AMOUNT_TOL = 45.0    # points: an amount under the amount-column header


def _cx(w):
    return (w["x0"] + w["x1"]) / 2


def find_table_headers(lines, cfg):
    """Header rows of payment tables: a code-column word ("סמל", "קוד") and an amount-column word ("סכום") on one line.
    A line may hold two tables side by side; each code word is paired with the amount word of its own table (RTL: the
    amount column is left of the code column, and right of the next table's code column)."""
    hcfg = cfg["parsing"].get("table_header", {})
    code_keys = {canon(k) for k in hcfg.get("code", ["סמל", "קוד"])}
    amount_keys = {canon(k) for k in hcfg.get("amount", ["סכום"])}
    payments_keys = [canon(k) for k in hcfg.get("payments_title", ["תשלומים"])]
    other_keys = [canon(k) for k in hcfg.get("other_titles", ["ניכויים", "ניכוי", "הפרשות", "זקיפות"])]
    headers = []
    for idx, line in enumerate(lines):
        code_ws, amount_ws = [], []
        for w in line["words"]:
            for t in (w["text"], w["text"][::-1]):
                c = canon(t)
                if c in code_keys:
                    code_ws.append(w)
                    break
                if c in amount_keys:
                    amount_ws.append(w)
                    break
        code_ws.sort(key=lambda w: -w["x0"])
        for k, cw in enumerate(code_ws):
            limit = _cx(code_ws[k + 1]) if k + 1 < len(code_ws) else float("-inf")         # next table to the left
            upper = _cx(code_ws[k - 1]) if k > 0 else float("inf")                         # previous table to the right
            left = [a for a in amount_ws if limit < _cx(a) < _cx(cw)]
            right = [a for a in amount_ws if _cx(cw) < _cx(a) < upper]
            # RTL tables usually end with the amount (left of the code); some put it first, right of the code
            mine = left or right
            if mine:
                aw = max(mine, key=_cx) if left else min(mine, key=_cx)
                beyond = ([w for w in line["words"] if limit < _cx(w) < aw["x0"]] if left else [])   # more columns after the amount
                headers.append({"top": line["top"], "code_x": _cx(cw), "amount_x": _cx(aw), "amount_is_last": not beyond,
                                "kind": _table_kind(lines[max(0, idx - 3):idx], line["top"],
                                                    max(limit, min(_cx(aw), _cx(cw)) - 80), min(upper, max(_cx(aw), _cx(cw)) + 80),
                                                    payments_keys, other_keys),
                                "id": len(headers)})
    return headers


def _table_kind(lines_above, header_top, x_lo, x_hi, payments_keys, other_keys):
    """"payments", "other" (deductions and the like) or None, from a table title just above the header row."""
    for line in reversed(lines_above):
        if header_top - line["top"] > 60:
            continue
        words = [w for w in line["words"] if x_lo < _cx(w) < x_hi]
        for _, seq in variants(words):
            joined = "".join(canon(t) for _, t in seq)
            if any(k in joined for k in other_keys):
                return "other"
            if any(k in joined for k in payments_keys):
                return "payments"
    return None


def _is_code(text, code, rev_digits):
    t = text.strip().strip(".:-")
    if rev_digits:
        t = t[::-1]
    return t.isdigit() and t.lstrip("0") == code.lstrip("0")


def find_code_amounts(lines, code, split_x, rev_digits, headers=()):
    """Amounts of every payment row whose code column equals `code`. Returns [(amount or None, method, table)];
    table identifies an untitled table (None otherwise), so a code found in two untitled tables can be reported.

    With a table header above the row (see find_table_headers) the amount is the number under the amount header
    ("column"). Otherwise the payments table is taken to be on the right: the 'נתונים נוספים' panel sits left of
    `split_x` on the same lines, so words left of it are ignored, and the leftmost number of the row is the amount
    ("row end"). When both ways work and disagree, the method is "disagree"."""
    found = []
    for line in lines:
        for w in line["words"]:
            if not _is_code(w["text"], code, rev_digits):
                continue
            nums = [(x, to_number(x["text"], rev_digits)) for x in line["words"] if x is not w]
            nums = [(x, v) for x, v in nums if v is not None]
            column = None
            above = [h for h in headers if h["top"] < line["top"] and abs(h["code_x"] - _cx(w)) <= HEADER_CODE_TOL]
            if not above and any(h["top"] < line["top"] for h in headers):
                continue                       # the tables have headers, and this number is not under a code column
            h = max(above, key=lambda h: h["top"]) if above else None
            if h and h["kind"] == "other":
                continue                       # a deductions table: the same code number means something else there
            if h:
                near = [(abs(_cx(x) - h["amount_x"]), v) for x, v in nums if abs(_cx(x) - h["amount_x"]) <= HEADER_AMOUNT_TOL]
                column = min(near)[1] if near else None
            row_end = None
            if w["x0"] >= split_x and (h is None or h["amount_is_last"]):     # the row end is only the amount if nothing follows it
                row = [(x, v) for x, v in nums if split_x <= x["x0"] < w["x0"]]
                row_end = min(row, key=lambda p: p[0]["x0"])[1] if row else None
            table = h["id"] if h and h["kind"] is None else None    # untitled table: remember which, to spot doubles
            if column is not None:
                found.append((column, "column" if row_end in (None, column) else "disagree", table))
            elif h is not None:
                found.append((None, "column", table))
            elif w["x0"] >= split_x:
                found.append((row_end, "row end", None))
    return found


def unknown_reserve_codes(lines, known, rev_digits, cfg):
    """Codes of payment rows whose description mentions reserve duty ("מילואים", "מיל'") but that are not configured.
    A row's code is its rightmost word when that is a 1-5 digit number (RTL tables start with the code)."""
    keys = [canon(k) for k in cfg.get("reserve_row_words", ["מילואים"])]
    found = set()
    for line in lines:
        words = sorted(line["words"], key=lambda w: -w["x0"])
        if not words:
            continue
        first = words[0]["text"].strip().strip(".:-")
        first = first[::-1] if rev_digits else first
        if not re.fullmatch(r"\d{1,5}", first) or first.lstrip("0") in known:
            continue
        texts = [canon(w["text"]) for w in words[1:]] + [canon(w["text"][::-1]) for w in words[1:]]
        if any(t == k or t.endswith(k) for t in texts for k in keys if t):         # "מילואים", "תגמולימילואים", "יו.מילואים"
            found.add(first)
    return found


def find_page_footer(lines):
    """Read the page number "עמוד X מתוך N" wherever it is on the page: the footer of a pay-slip, the HEADER of
    the reserve-service certificate. Returns (X, N) or None.

    The numbers are picked by position, not "the first two on the line": X is the number between the words
    "עמוד" and "מתוך", N the number just beyond "מתוך". Other numbers on the same line (a form number, a date)
    are ignored."""
    keys = ("עמודמתוך", "דףמתוך")
    cx = lambda w: (w["x0"] + w["x1"]) / 2
    for line in reversed(lines):
        words = line["words"]
        page_w = of_w = None
        for w in words:
            for t in (w["text"], w["text"][::-1]):               # logical or visually reversed letters
                c = canon(t)
                if c in ("עמוד", "דף") and page_w is None:
                    page_w = w
                elif c == "מתוך" and of_w is None:
                    of_w = w
        ints = [w for w in words if re.fullmatch(r"\d{1,3}", w["text"].strip())]
        if page_w is not None and of_w is not None and ints:
            a, b = cx(page_w), cx(of_w)
            between = [w for w in ints if min(a, b) < cx(w) < max(a, b)]
            beyond = [w for w in ints if (cx(w) < b if a > b else cx(w) > b)]
            if between and beyond:
                x = int(min(between, key=lambda w: abs(cx(w) - (a + b) / 2))["text"])
                n = int(min(beyond, key=lambda w: abs(cx(w) - b))["text"])
                return (x, n) if x <= n else None
        for _, seq in variants(words):                           # fallback: words glued together, odd layouts
            if any(k in "".join(canon(t) for _, t in seq) for k in keys):
                ordered = sorted(ints, key=lambda w: -w["x0"])   # RTL: X is right of N
                if len(ordered) >= 2:
                    x, n = int(ordered[0]["text"]), int(ordered[1]["text"])
                    return (x, n) if x <= n else (n, x)
                return None
    return None


def diagnose(pdf_path):
    """Content-free report on what kind of PDF this is (counts only, no text or numbers)."""
    try:
        import pdfplumber
    except ImportError:
        print("pdfplumber is not installed - skipping its report (%s)" % install_command())
        pdfplumber = None
    print("== %s" % Path(pdf_path).name)
    if pdfplumber is None:
        _diagnose_pymupdf(pdf_path)
        return
    with pdfplumber.open(pdf_path) as pdf:
        print("pages: %d   encrypted: %s" % (len(pdf.pages), bool(getattr(pdf, "doc", None) and pdf.doc.encryption)))
        for no, page in enumerate(pdf.pages, 1):
            chars = page.chars
            words = page.extract_words()
            text = "".join(c["text"] for c in chars)
            cats = {
                "hebrew letters": sum("א" <= ch <= "ת" for ch in text),
                "latin1 e0-fa (cp1255 style)": sum("à" <= ch <= "ú" for ch in text),
                "ascii letters": sum(("a" <= ch.lower() <= "z") for ch in text),
                "digits": sum(ch.isdigit() for ch in text),
                "(cid:N) markers": text.count("(cid:"),
                "private-use/odd (U+E000-F8FF, U+FFFD)": sum("" <= ch <= "" or ch == "�" for ch in text),
            }
            fonts = sorted({c.get("fontname", "?").split("+")[-1] for c in chars})[:6]
            print("page %d: %.0fx%.0f  chars=%d  words=%d  images=%d  rects=%d  lines=%d" % (
                no, page.width, page.height, len(chars), len(words), len(page.images), len(page.rects), len(page.lines)))
            print("   char classes:", {k: v for k, v in cats.items() if v})
            print("   fonts:", fonts)
            if not chars and page.images:
                print("   -> pdfplumber sees no text layer on this page")
            if no >= 3:
                print("(remaining pages skipped)")
                break

    _diagnose_pymupdf(pdf_path)


def _diagnose_pymupdf(pdf_path):
    try:
        import pymupdf
    except ImportError:
        print("PyMuPDF not installed - skipping its report (%s)" % install_command())
        return
    doc = pymupdf.open(pdf_path)
    print("--- PyMuPDF second opinion (counts only) ---")
    for no, page in enumerate(doc, 1):
        text = page.get_text("text")
        words = page.get_text("words")
        heb = sum("א" <= ch <= "ת" for ch in text)
        lat = sum("à" <= ch <= "ú" for ch in text)
        digs = sum(ch.isdigit() for ch in text)
        fonts = sorted({f[3].split("+")[-1] for f in page.get_fonts()})[:6]
        print("page %d: text chars=%d words=%d hebrew=%d latin1(cp1255-style)=%d digits=%d images=%d fonts=%s" % (
            no, len(text.strip()), len(words), heb, lat, digs, len(page.get_images()), fonts))
        if no >= 3:
            break


def split_into_slips(pages, cfg):
    """Group the pages of a PDF into slips.
    1) Footer "עמוד X מתוך N": X == 1 starts a new slip, X > 1 continues the current one.
    2) Pages without a footer: a page with the title ("תלוש שכר לחודש <month> <year>") starts a slip
       (merged with the previous one if it is the same month/year); other pages continue it.
    Returns [(label, pages, note)] where note is a warning text or ''."""
    groups = []  # {first, last, pages, ym, expected}
    for no, p in enumerate(pages, 1):
        lines = group_lines(p["words"], cfg["parsing"]["line_tolerance"])
        t = find_title(lines, cfg)
        ym = (t[1], t[0]) if t else None
        footer = find_page_footer(lines)
        if footer:
            x, n = footer
            new = x == 1 or not groups
        elif ym is None:
            new = not groups
        else:
            new = not groups or groups[-1]["ym"] != ym
        if new:
            groups.append({"first": no, "last": no, "pages": [p], "ym": ym, "expected": footer[1] if footer else None})
        else:
            g = groups[-1]
            g["last"] = no
            g["pages"].append(p)
            g["ym"] = g["ym"] or ym
            if footer and not g["expected"]:
                g["expected"] = footer[1]
    out = []
    for g in groups:
        label = "page %d" % g["first"] if g["first"] == g["last"] else "pages %d-%d" % (g["first"], g["last"])
        note = ""
        if g["expected"] and g["expected"] != len(g["pages"]):
            note = "footer says %d page(s) but %d page(s) were grouped" % (g["expected"], len(g["pages"]))
        out.append((label, g["pages"], note))
    return out


def parse_slip(source, cfg, debug=False, force_month=None, force_year=None, split_ratio=None, text_dump=False):
    """`source` is a PDF path, or a list of pages as returned by load_words."""
    pages = source if isinstance(source, list) else load_words(source)
    all_lines, width = [], 0.0
    for p in pages:
        width = max(width, p["width"])
        all_lines += group_lines(p["words"], cfg["parsing"]["line_tolerance"])
    split_x = width * (split_ratio or cfg["parsing"]["left_panel_ratio"])
    out = {"warnings": []}

    if debug:
        print("page width: %.0f   left-panel boundary x: %.0f   lines: %d" % (width, split_x, len(all_lines)))
        if text_dump:
            known = {"שכר", "תלוש", "ברוטו", "סהכ", "תשלומים", "מילואים", "הכנסה", "מס", "נתונים", "פרוט", "נוספים", "זקיפות"}
            hits_asis = sum(canon(w["text"]) in known for l in all_lines for w in l["words"])
            hits_flip = sum(canon(w["text"][::-1]) in known for l in all_lines for w in l["words"])
            flip = hits_flip > hits_asis
            print("Hebrew words stored %s; shown below in reading order. Digits are masked (#); "
                  "3-digit codes in the payments table are kept." % ("reversed (fixed for display)" if flip else "normally"))
            print("REVIEW BEFORE SHARING: delete any line that contains a name or other personal text.")
            for n, line in enumerate(all_lines):
                ws = sorted(line["words"], key=lambda w: -w["x0"])
                parts = []
                for i, w in enumerate(ws):
                    t = w["text"]
                    if i == 0 and w["x0"] >= split_x and re.fullmatch(r"\d{3}", t.strip()):
                        parts.append(t.strip())          # payments-table code
                        continue
                    t = re.sub(r"\d", "#", t)
                    parts.append(t[::-1] if (flip and HEB.search(t)) else t)
                print("%3d | %s" % (n, "  ".join(parts)))
        else:
            for n, line in enumerate(all_lines):
                ws = sorted(line["words"], key=lambda w: -w["x0"])
                print("%3d | %s" % (n, "  ".join("%s@%d" % (shape(w["text"]), w["x0"]) for w in ws)))

    title = find_title(all_lines, cfg)
    rev = False
    if title:
        out["month"], out["year"], rev = title
    if force_month:
        out["month"] = force_month
    if force_year:
        out["year"] = force_year
    if "month" not in out or "year" not in out:
        out["warnings"].append("could not read month/year from the title (use --month/--year)")
    out["digits_reversed"] = rev

    if debug:
        counts = {}
        for line in all_lines:
            right = [w for w in line["words"] if w["x0"] >= split_x]
            if right:
                first = max(right, key=lambda w: w["x0"])["text"].strip()
                if re.fullmatch(r"\d{3}", first):
                    counts[first] = counts.get(first, 0) + 1
        print("payment-table codes seen (rightmost 3-digit token per row): %s" % (
            ", ".join("%s x%d" % kv for kv in sorted(counts.items())) or "none"))

    headers = find_table_headers(all_lines, cfg)
    if debug:
        print("payment-table headers found: %d" % len(headers))
    for field, codes in cfg["codes"].items():
        codes = [codes] if isinstance(codes, str) else codes      # a field may sum several codes
        amounts = []
        for code in codes:
            got = find_code_amounts(all_lines, code, split_x, rev, headers)
            if any(a is None for a, _, _ in got):
                out["warnings"].append("code %s: row found but no amount read" % code)
            if any(m == "disagree" for _, m, _ in got):
                out["warnings"].append("code %s: the amount column and the end of the row give different amounts - check it" % code)
            if len({t for _, _, t in got if t is not None}) > 1:
                out["warnings"].append("code %s: found in more than one table (payments and deductions?) - all were added, check it" % code)
            amounts += [a for a, _, _ in got]
            if debug:
                print("code %s: %d row(s), amounts %s, via %s" % (code, len(got), ["#" if a is not None else None for a, _, _ in got],
                                                                  sorted({m for _, m, _ in got}) or "-"))
        out[field] = sum(a for a in amounts if a is not None)
        out[field + "_rows"] = len(amounts)
        out[field + "_codes"] = "+".join(codes)

    known = {c.lstrip("0") for v in cfg["codes"].values() for c in ([v] if isinstance(v, str) else v)}
    # Only for a slip that has NONE of the configured codes (another payroll system): on a slip that has them, other
    # rows mentioning reserve duty are something else and need no warning.
    found_known = any(out.get(f + "_rows") for f in cfg["codes"])
    unknown = [] if found_known else sorted(unknown_reserve_codes(all_lines, known, rev, cfg))
    if unknown:
        out["warnings"].append("rows that mention reserve duty (מילואים) have code(s) %s, which are not in config.json - "
                               "they were NOT included; check whether they are reserve payments" % ", ".join(unknown))

    for field in ("btl_gross", "btl_variable"):
        value = method = None
        avoid = [canon(x) for x in cfg.get("label_not_followed_by", {}).get(field, [])]
        for label in cfg["labels"][field]:
            key = canon(label)
            for line in all_lines:
                r = find_value_for_label(line, key, rev, avoid)
                if r and r[1] is not None:
                    method, value = "%s/%s" % (r[0], r[2]), r[1]
                    break
            if value is not None:
                break
        out[field] = value
        out[field + "_how"] = method
        if debug:
            if value is not None:
                print("%s: found via %s" % (field, method))
            elif field == "btl_variable":
                print("%s: not on this slip -> will be written as 0" % field)
            else:
                print("%s: NOT FOUND" % field)
            if value is None and field != "btl_variable":
                for n, line in enumerate(all_lines):
                    for _, seq in variants(line["words"]):
                        joined = "".join(canon(t) for _, t in seq)
                        if "ברוטו" in joined:
                            print("   candidate line %d: %s" % (n, " ".join(shape(t) if not HEB.search(t) else t for _, t in seq)))
                            break
    if out["btl_gross"] is None:
        out["warnings"].append("ברוטו לב.לאומי not found")
    return out


# ------------------------------------------------------------------- workbook
def open_template(path=None, with_summary=True, cfg=None):
    """Open a workbook: from `path` if given, else the built-in template embedded in template_data.py.
    With `with_summary` the "total National Insurance owes me" block is added to the helper sheet if it is missing
    (so the built-in template has it, and files made by older versions get it when they are updated)."""
    import base64
    import io

    from openpyxl import load_workbook

    if path:
        wb = load_workbook(str(path), keep_vba=False)
    else:
        try:
            from template_data import TEMPLATE_B64
        except ImportError:
            raise SystemExit("ERROR: the built-in template is missing (template_data.py). Run embed_template.py.")
        wb = load_workbook(io.BytesIO(base64.b64decode(TEMPLATE_B64)), keep_vba=False)
    if with_summary:
        ensure_summary_table(wb, cfg or json.loads((HERE / "config.json").read_text(encoding="utf-8")))
    return wb


def ensure_summary_table(wb, cfg):
    """Add the small block "סה"כ ביטוח לאומי חייב לי" to the helper sheet: a blue label and, below it, a green cell
    that sums every year sheet's annual difference (F38). Idempotent; never overwrites other content.
    Returns True if the block was added."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    scfg, wcfg = cfg["summary"], cfg["workbook"]
    if wcfg["helper_sheet"] not in wb.sheetnames:
        return False
    hs = wb[wcfg["helper_sheet"]]
    label_cell, value_cell = scfg["label_range"].split(":")[0], scfg["value_range"].split(":")[0]
    if hs[label_cell].value == scfg["label"]:
        return False                                                # already there
    area = [c for rng in (scfg["label_range"], scfg["value_range"]) for row in hs[rng] for c in row]
    if any(c.value not in (None, "") for c in area):
        return False                                                # something else is there: leave it alone
    years = [y for y in wcfg["year_sheets"] if y in wb.sheetnames]
    if not years:
        return False
    merged = {str(r) for r in hs.merged_cells.ranges}
    for rng in (scfg["label_range"], scfg["value_range"]):
        if rng not in merged:                           # the template may already have the (empty) merged cells
            hs.merge_cells(rng)
    hs[label_cell].value = scfg["label"]
    hs[value_cell].value = "=" + "+".join("'%s'!%s" % (y, scfg["year_total_cell"]) for y in years)
    edge = Side(style="thin", color="2F5F8F")
    for rng, color, fmt in ((scfg["label_range"], "4A90E2", None), (scfg["value_range"], "70AD47", "[$₪-40D]\\ #,##0.00")):
        for row in hs[rng]:
            for c in row:
                c.fill = PatternFill("solid", fgColor=color)
                c.font = Font(bold=True, size=14, color="000000")
                c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
                c.border = Border(top=edge, bottom=edge, left=edge, right=edge)
                if fmt:
                    c.number_format = fmt
    for rng in (scfg["label_range"], scfg["value_range"]):
        hs.row_dimensions[hs[rng][0][0].row].height = 26
    return True


def _date_of(cell_value):
    if isinstance(cell_value, dt.datetime):
        return cell_value.date()
    if isinstance(cell_value, dt.date):
        return cell_value
    if isinstance(cell_value, (int, float)):
        from openpyxl.utils.datetime import from_excel
        return from_excel(cell_value).date()
    return None


def write_slip(wb, slip, cfg, overwrite, report):
    from openpyxl.utils import get_column_letter

    wcfg = cfg["workbook"]
    year, month = slip["year"], slip["month"]
    if str(year) not in wb.sheetnames:
        report.append("  ERROR: no sheet '%s' in the workbook - nothing written" % year)
        return False
    ws = wb[str(year)]
    first = ws[wcfg["first_month_column"] + "1"].column
    col = get_column_letter(first + month - 1)

    targets = [
        (ws, "%s%d" % (col, wcfg["rows"]["btl_gross"]), slip["btl_gross"], "ברוטו לב.לאומי"),
        (ws, "%s%d" % (col, wcfg["rows"]["current_payment"]), slip["current_payment"], "קוד " + slip.get("current_payment_codes", "198+197")),
        (ws, "%s%d" % (col, wcfg["rows"]["completion"]), slip["completion"], "קוד " + slip.get("completion_codes", "193")),
    ]
    hs = wb[wcfg["helper_sheet"]]
    hrow = None
    for r in range(1, hs.max_row + 1):
        d = _date_of(hs["%s%d" % (wcfg["helper_date_column"], r)].value)
        if d and d.year == year and d.month == month:
            hrow = r
            break
    if hrow is None:
        report.append("  ERROR: month %02d/%d not found in helper sheet" % (month, year))
        return False
    targets.append((hs, "%s%d" % (wcfg["helper_value_column"], hrow), slip["btl_variable"] or 0.0, "ברוטו בלתי קבוע"))

    ok = True
    for sheet, addr, value, name in targets:
        if value is None:                              # not found on this slip: leave the cell for manual input
            report.append("  MISSING %-18s %s!%s  not found on the slip - fill it in by hand" % (name, sheet.title, addr))
            ok = False
            continue
        cell = sheet[addr]
        cur = cell.value
        same = isinstance(cur, (int, float)) and abs(cur - value) < 0.005
        if cur not in (None, "") and not same and not overwrite:
            shown = "{:,.2f}".format(cur) if isinstance(cur, (int, float)) else repr(cur)
            report.append("  SKIPPED %-18s %s!%s  cell has: %s   slip has: %s   (use --overwrite to replace)"
                          % (name, sheet.title, addr, shown, "{:,.2f}".format(value)))
            ok = False
            continue
        cell.value = round(value, 2)
        report.append("  wrote   %-18s %s!%s = %s" % (name, sheet.title, addr, "{:,.2f}".format(value)))
    return ok


# ------------------------------------------------------------------------ main
# ------------------------------------------------- reserve-service certificate (form 3010)
DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
DAYS_RE = re.compile(r"^\d+\.\d+$")


def parse_form_date(token):
    """dd/mm/yyyy -> (date, digits_reversed) or None. A token stored digit-reversed (6202/50/82) is repaired."""
    t = token.strip(" ,;:()[]")
    m, reversed_ = DATE_RE.match(t), False
    if not m:
        m, reversed_ = DATE_RE.match(t[::-1]), True
    if not m:
        return None
    try:
        return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))), reversed_
    except ValueError:
        return None


def parse_form_rows(lines, page_no):
    """Rows of the service table: a line with exactly two dates (start, end) and one 'N.N' days figure."""
    rows = []
    for line in lines:
        dates, days = [], []
        for w in line["words"]:
            parsed = parse_form_date(w["text"])
            if parsed:
                dates.append(parsed)
                continue
            t = w["text"].strip()
            if DAYS_RE.match(t) or DAYS_RE.match(t[::-1]):
                days.append(t)
        if len(dates) != 2 or len(days) != 1:
            continue                                     # headings, body text, page header
        text = days[0][::-1] if dates[0][1] else days[0]
        start, end = sorted(d for d, _ in dates)         # a period always runs from the earlier to the later date
        rows.append({"start": start, "end": end, "days": float(text), "page": page_no})
    return rows


def is_reserve_form_page(lines, cfg):
    """True for a page of the reserve-service certificate: its subject line, or at least one service-table row."""
    key = canon(cfg["reserve_form"]["phrase"])
    for line in lines:
        for _, seq in variants(line["words"]):
            if key in "".join(canon(t) for _, t in seq):
                return True
    return bool(parse_form_rows(lines, 0))


# ------------------------------------- NI annual letter: "אישור שנתי בגין תשלומים שהועברו למעסיק"
LETTER_AMOUNT_RE = re.compile(r"^\d{1,3}(?:,\d{3})+(?:\.\d+)?$|^\d+(?:\.\d+)?$")


def _letter_columns(lines, cfg):
    """x centres of the "ימי שירות" and "שולם למעסיק" columns from the table header of a page, or None."""
    days_x = paid_x = None
    for line in lines:
        for w in line["words"]:
            for t in (w["text"], w["text"][::-1]):
                c = canon(t)
                if c == "ימי" and days_x is None:
                    days_x = _cx(w)
                elif c == "שולם" and paid_x is None:
                    paid_x = _cx(w)
        if days_x is not None and paid_x is not None:
            return days_x, paid_x
    return None


def parse_letter_rows(lines, page_no, columns=None):
    """Rows of the letter's table: two dates (start, end of service), the days and the amount paid to the employer.
    The employer's name in between is ignored (kept only as text, to tell two employers apart). With the column
    positions from the header the two numbers are picked by column; without them, a row must have exactly two numbers
    and the larger one is the amount (a day is always paid far more than 1 shekel)."""
    rows = []
    for line in lines:
        dates, nums, text = [], [], []
        for w in line["words"]:
            parsed = parse_form_date(w["text"])
            if parsed:
                dates.append(parsed)
                continue
            t = w["text"].strip()
            if LETTER_AMOUNT_RE.match(t) or LETTER_AMOUNT_RE.match(t[::-1]):
                nums.append((t, w))
            elif canon(t):
                text.append(canon(t))
        if len(dates) != 2 or len(nums) < 2:
            continue
        rev = dates[0][1]                                # digits stored reversed: the dates tell (as in the certificate)
        nums = [((t[::-1] if rev else t), w) for t, w in nums]
        if not all(LETTER_AMOUNT_RE.match(t) for t, _ in nums):
            continue
        nums = [(float(t.replace(",", "")), w) for t, w in nums]
        if columns:
            days_x, paid_x = columns
            days = min(nums, key=lambda n: abs(_cx(n[1]) - days_x))
            paid = min(nums, key=lambda n: abs(_cx(n[1]) - paid_x))
            if days is paid:
                continue
        elif len(nums) == 2:
            days, paid = sorted(nums, key=lambda n: n[0])
        else:
            continue
        if days[0] != int(days[0]) or days[0] <= 0:
            continue
        start, end = sorted(d for d, _ in dates)
        rows.append({"start": start, "end": end, "days": int(days[0]), "amount": paid[0], "page": page_no,
                     "employer": "".join(sorted(text))})
    return rows


def letter_total(lines):
    """The amount on the "סה"כ לתשלום" line, or None."""
    for line in lines:
        for _, seq in variants(line["words"]):
            if "סהכלתשלום" in "".join(canon(t) for _, t in seq):
                for w in line["words"]:
                    t = w["text"].strip(" :;₪")                     # ":12,345" - the colon may be glued to it
                    for cand in (t, t[::-1]):
                        if cand and LETTER_AMOUNT_RE.match(cand.strip(" :;")):
                            cand = cand.strip(" :;")
                            return float(cand.replace(",", ""))
    return None


def is_letter_page(lines, cfg):
    """A page of the NI annual letter: its subject line or the table header "שולם למעסיק"."""
    keys = [canon(cfg["nis_letter"]["phrase"]), canon(cfg["nis_letter"]["table_header"])]
    for line in lines:
        for _, seq in variants(line["words"]):
            joined = "".join(canon(t) for _, t in seq)
            if any(k in joined for k in keys):
                return True
    return False


def parse_nis_letter(pages, cfg, name=""):
    """Read the payment table of one NI letter (all its pages)."""
    rows, seen, totals, total_amount = [], [], set(), None
    columns = None
    for no, p in enumerate(pages, 1):
        lines = group_lines(p["words"], cfg["parsing"]["line_tolerance"])
        footer = find_page_footer(lines)
        if footer:
            seen.append(footer[0])
            totals.add(footer[1])
        columns = _letter_columns(lines, cfg) or columns
        rows += parse_letter_rows(lines, footer[0] if footer else no, columns)
        total_amount = letter_total(lines) or total_amount
    return {"name": name, "rows": rows, "pages_seen": seen, "totals": sorted(totals), "total_amount": total_amount,
            "ocr": any(p.get("ocr") for p in pages), "page_count": len(pages)}


def combine_letters(letters, sheet_years):
    """Per-month payments from all NI letters of a run: {(year, month): {"days", "amount"}}.

    Identical rows (the same letter uploaded twice) count once. A row that crosses a month end is split by calendar days
    at its own daily rate. Days are counted once even when two employers were paid for the same day (the amounts add up),
    so amount / days is the value paid per day. A row whose stated days differ from its calendar span is kept (it is
    what was paid) and reported."""
    seen, duplicates, mismatches, total_checks = {}, 0, [], []
    for letter in letters:
        for row in letter["rows"]:
            key = (row["start"], row["end"], row["days"], row["amount"], row["employer"])
            if key in seen:
                duplicates += 1
                continue
            seen[key] = row
            span = (row["end"] - row["start"]).days + 1
            if span != row["days"]:
                mismatches.append(dict(row, span=span, file=letter["name"]))
        if letter["total_amount"] is not None:
            got = sum(r["amount"] for r in letter["rows"])
            if abs(got - letter["total_amount"]) > 0.5:
                total_checks.append({"file": letter["name"], "stated": letter["total_amount"], "read": got})

    days_of, amount_of, ignored = {}, {}, {}
    for row in seen.values():
        span = (row["end"] - row["start"]).days + 1
        per_month = {}
        day = row["start"]
        while day <= row["end"]:
            per_month.setdefault((day.year, day.month), []).append(day)
            day += dt.timedelta(days=1)
        for (year, month), days in per_month.items():
            share = len(days) / span
            if year not in sheet_years:
                ignored[year] = ignored.get(year, 0) + row["amount"] * share
                continue
            days_of.setdefault((year, month), set()).update(days)
            amount_of[(year, month)] = amount_of.get((year, month), 0.0) + row["amount"] * share
    months = {key: {"days": len(days), "amount": round(amount_of[key], 2)} for key, days in days_of.items()}

    warnings, missing_pages = [], []
    for total in sorted({t for f in letters for t in f["totals"]}):
        have = {x for f in letters if total in f["totals"] for x in f["pages_seen"]}
        missing = sorted(set(range(1, total + 1)) - have)
        if missing:
            missing_pages.append({"total": total, "missing": missing})
    all_days = sorted({d for row in seen.values() for d in
                       (row["start"] + dt.timedelta(days=i) for i in range((row["end"] - row["start"]).days + 1))})
    return {"months": months, "rows": len(seen), "duplicates": duplicates, "mismatches": mismatches, "all_days": all_days,
            "missing_pages": missing_pages, "total_checks": total_checks, "ignored": ignored,
            "ocr_files": [f["name"] for f in letters if f.get("ocr")]}


def write_nis_payments(wb, months, cfg):
    """Rows 22 (value paid per day = amount / days, 2 decimals) and 23 (days paid) from the NI letter.
    Like the reserve days, the letter's figures replace what is in those cells."""
    from openpyxl.utils import get_column_letter, column_index_from_string

    wcfg = cfg["workbook"]
    first = column_index_from_string(wcfg["first_month_column"])
    out = []
    for (year, month), m in sorted(months.items()):
        if str(year) not in wb.sheetnames or not m["days"]:
            continue
        ws = wb[str(year)]
        col = get_column_letter(first + month - 1)
        for key, value in (("nis_paid_value", round(m["amount"] / m["days"], 2)), ("nis_paid_days", m["days"])):
            cell = ws["%s%d" % (col, wcfg["rows"][key])]
            prev = cell.value
            cell.value = value
            out.append({"year": year, "month": month, "cell": "%s!%s" % (year, cell.coordinate), "value": value,
                        "previous": prev, "replaced": prev not in (None, "") and prev != value})
    return out


def split_by_kind(pages, cfg):
    """[(kind, pages)] with kind 'letter' (NI annual letter), 'form' (service certificate) or 'slip'; consecutive pages
    of one kind stay together."""
    runs = []
    for p in pages:
        lines = group_lines(p["words"], cfg["parsing"]["line_tolerance"])
        kind = "letter" if is_letter_page(lines, cfg) else "form" if is_reserve_form_page(lines, cfg) else "slip"
        if kind == "slip" and runs and runs[-1][0] != "slip" and not find_title(lines, cfg):
            kind = runs[-1][0]          # a closing page of the letter / certificate (total, signature, contact details)
        if runs and runs[-1][0] == kind:
            runs[-1][1].append(p)
        else:
            runs.append((kind, [p]))
    return runs


def parse_reserve_form(pages, cfg, name=""):
    """Read the service table of one certificate (all its pages)."""
    rows, seen, totals = [], [], set()
    for no, p in enumerate(pages, 1):
        lines = group_lines(p["words"], cfg["parsing"]["line_tolerance"])
        footer = find_page_footer(lines)
        if footer:
            seen.append(footer[0])
            totals.add(footer[1])
        rows += parse_form_rows(lines, footer[0] if footer else no)
    return {"name": name, "rows": rows, "pages_seen": seen, "totals": sorted(totals),
            "ocr": any(p.get("ocr") for p in pages), "page_count": len(pages)}


def combine_reserve(forms, sheet_years):
    """Per-month service days from all certificates of a run.

    Every calendar day is counted once: identical rows are merged, overlapping rows never count a day twice.
    A period is split at month ends (and year ends); months of years without a sheet are ignored.
    A row whose stated days differ from its calendar span is set aside for the user to check."""
    seen_rows, duplicates, mismatches = {}, 0, []
    for form in forms:
        for row in form["rows"]:
            span = (row["end"] - row["start"]).days + 1
            if abs(span - row["days"]) > 1e-9:
                mismatches.append(dict(row, span=span, file=form["name"]))
                continue
            key = (row["start"], row["end"])
            if key in seen_rows:
                duplicates += 1
                continue
            seen_rows[key] = row

    owner, contrib, overlaps = {}, {}, []
    for (start, end) in sorted(seen_rows):
        day = start
        while day <= end:
            if day in owner:
                if (owner[day], (start, end)) not in overlaps:
                    overlaps.append((owner[day], (start, end)))
            else:
                owner[day] = (start, end)
                part = contrib.setdefault((day.year, day.month), {})
                part[(start, end)] = part.get((start, end), 0) + 1
            day += dt.timedelta(days=1)

    months, ignored = {}, {}
    for (year, month), parts in contrib.items():
        n = sum(parts.values())
        if year in sheet_years:
            months[(year, month)] = n
        else:
            ignored[year] = ignored.get(year, 0) + n

    warnings, missing_pages = [], []
    for total in sorted({t for f in forms for t in f["totals"]}):
        have = {x for f in forms if total in f["totals"] for x in f["pages_seen"]}
        missing = sorted(set(range(1, total + 1)) - have)
        if missing:
            missing_pages.append({"total": total, "missing": missing})
            warnings.append("missing page(s) %s of %d - upload the whole certificate in one run, "
                            "otherwise months may be incomplete" % (", ".join(map(str, missing)), total))
    if overlaps:
        warnings.append("%d overlapping period(s) - each calendar day was counted once" % len(overlaps))
    if mismatches:
        warnings.append("%d row(s) where the stated days differ from the calendar days - left out, please check" % len(mismatches))
    notes = []
    ocr_files = [f["name"] for f in forms if f.get("ocr")]
    if ocr_files:
        notes.append("read by image recognition (OCR): %s - each row was checked against the calendar, but compare the totals "
                     "with the certificate, and make sure all of its pages are in the file (page numbers cannot be read from a picture)"
                     % ", ".join(ocr_files))
    parts_out = {k: [{"start": s, "end": e, "days": n} for (s, e), n in sorted(v.items())] for k, v in contrib.items() if k in months}
    return {"months": months, "parts": parts_out, "ignored": ignored, "rows": len(seen_rows), "duplicates": duplicates,
            "all_days": sorted(owner),
            "overlaps": overlaps, "mismatches": mismatches, "warnings": warnings, "missing_pages": missing_pages,
            "notes": notes, "ocr_files": ocr_files}


def write_reserve_days(wb, months, cfg):
    """Set row 27 (reserve days by the 3010) of each month that has service. Existing values are REPLACED, never added to."""
    from openpyxl.utils import get_column_letter

    wcfg, out = cfg["workbook"], []
    for (year, month), days in sorted(months.items()):
        if str(year) not in wb.sheetnames:
            continue
        ws = wb[str(year)]
        first = ws[wcfg["first_month_column"] + "1"].column
        addr = "%s%d" % (get_column_letter(first + month - 1), wcfg["rows"]["reserve_days"])
        value = int(days) if float(days).is_integer() else round(float(days), 2)
        previous = ws[addr].value
        ws[addr].value = value
        out.append({"year": year, "month": month, "cell": "%s!%s" % (year, addr), "value": value,
                    "previous": previous, "replaced": previous not in (None, "") and previous != value})
    return out


def letter_problems(res):
    return bool(res["missing_pages"] or res["mismatches"] or res["total_checks"])


def letter_summary_lines(res):
    """Plain-text lines for the command line: per month days paid, amount, value per day; then the problems."""
    out = ["  %02d/%d: %d day(s) paid, %s, %s per day" % (m, y, v["days"], "{:,.2f}".format(v["amount"]),
                                                     "{:,.2f}".format(v["amount"] / v["days"]))
           for (y, m), v in sorted(res["months"].items()) if v["days"]]
    for mp in res["missing_pages"]:
        out.append("  WARNING: missing page(s) %s of %d" % (", ".join(map(str, mp["missing"])), mp["total"]))
    for t in res["total_checks"]:
        out.append("  WARNING: %s: the rows add up to %s but the letter's total is %s - check that every page was read"
                   % (t["file"], "{:,.2f}".format(t["read"]), "{:,.2f}".format(t["stated"])))
    for mm in res["mismatches"]:
        out.append("  NOTE: %s - %s: %d day(s) stated, %d calendar days" % (mm["start"].strftime("%d/%m/%Y"), mm["end"].strftime("%d/%m/%Y"),
                                                                          mm["days"], mm["span"]))
    for y, a in sorted(res["ignored"].items()):
        out.append("  ignored: %d (no sheet for that year), %s" % (y, "{:,.2f}".format(a)))
    return out


def reserve_summary_lines(res):
    lines = ["Reserve service certificate: %d period(s), %d duplicate(s) merged" % (res["rows"], res["duplicates"])]
    for (year, month), days in sorted(res["months"].items()):
        lines.append("  %d-%02d : %s ימי מילואים" % (year, month, int(days) if float(days).is_integer() else days))
    for year, days in sorted(res["ignored"].items()):
        lines.append("  %d : %s days ignored (no sheet for that year)" % (year, days))
    for w in res["warnings"]:
        lines.append("  WARNING: %s" % w)
    for n in res.get("notes", []):
        lines.append("  NOTE: %s" % n)
    for m in res["mismatches"]:
        lines.append("  CHECK: %s - %s states %s day(s) but spans %d (%s)" % (m["start"], m["end"], m["days"], m["span"], m["file"]))
    return lines


def parse_day_value(text):
    """The "ערך ימ"מ נוכחי" typed by the user -> float >= 0. Empty means 0, i.e. "not entered".
    Raises ValueError with a Hebrew message for anything that is not a sensible amount."""
    if text is None or isinstance(text, bool):
        return 0.0
    if isinstance(text, (int, float)):
        value = float(text)
    else:
        cleaned = str(text).strip().replace(",", "").replace("₪", "").strip()
        if not cleaned:
            return 0.0
        try:
            value = float(cleaned)
        except ValueError:
            raise ValueError('ערך ימ"מ נוכחי חייב להיות מספר, למשל 1650.50')
    if value != value or value in (float("inf"), float("-inf")):
        raise ValueError('ערך ימ"מ נוכחי חייב להיות מספר, למשל 1650.50')
    if value < 0:
        raise ValueError('ערך ימ"מ נוכחי לא יכול להיות שלילי')
    if value > 100000:
        raise ValueError('ערך ימ"מ נוכחי גדול מדי. בדוק שהזנת סכום ליום אחד.')
    return round(value, 2)


def day_value_range_warning(value, cfg):
    """A warning (never an error) if the current day value is outside the expected range.
    Returns None, or {"kind": "above" | "below", "limit": the boundary that was crossed, "value": value}.
    0 means "not entered" and is handled separately, so it is not reported here."""
    limits = cfg.get("day_value_limits")
    if not limits or not value or value <= 0:
        return None
    if value > limits["max"]:
        return {"kind": "above", "limit": limits["max"], "value": value}
    if value < limits["min"]:
        return {"kind": "below", "limit": limits["min"], "value": value}
    return None


def day_value_range_text(w):
    """One line for the command line (and for logs)."""
    return ("the current day value %s is %s the %s of %s - check that it is the value of ONE day; it was still written"
            % (w["value"], "above" if w["kind"] == "above" else "below", "maximum" if w["kind"] == "above" else "minimum", w["limit"]))


def write_current_day_value(wb, value, cfg):
    """Write the current day value into the cell the year sheets' formulas read (F36) in EVERY year sheet,
    replacing what is there."""
    wcfg, out = cfg["workbook"], []
    for year in wcfg["year_sheets"]:
        if year not in wb.sheetnames:
            continue
        ws, addr = wb[year], wcfg["current_value_cell"]
        previous = ws[addr].value
        ws[addr].value = int(value) if float(value).is_integer() else value
        out.append({"year": year, "cell": "%s!%s" % (year, addr), "value": ws[addr].value, "previous": previous,
                    "replaced": previous not in (None, "", 0) and previous != value})
    return out


def years_missing_day_value(wb, cfg):
    """Year sheets with neither a current day value (F36) nor any value paid per day from the NI letter (row 22)."""
    from openpyxl.utils import get_column_letter, column_index_from_string

    wcfg = cfg["workbook"]
    first = column_index_from_string(wcfg["first_month_column"])
    out = []
    for y in wcfg["year_sheets"]:
        if y not in wb.sheetnames:
            continue
        ws = wb[y]
        if _number(ws[wcfg["current_value_cell"]].value):
            continue
        row = wcfg["rows"].get("nis_paid_value")
        if row and any(_number(ws["%s%d" % (get_column_letter(first + m), row)].value) for m in range(12)):
            continue
        out.append(y)
    return out


def inspect_workbook(path, cfg):
    """Number-free overview of which input cells of a result workbook hold a value.
    # = a number, 0 = zero, . = empty, ? = text or a formula. Safe to share: no figures are shown."""
    from openpyxl.utils import get_column_letter

    wb = open_template(path)
    wcfg = cfg["workbook"]
    first = wb[wcfg["year_sheets"][0]][wcfg["first_month_column"] + "1"].column if wcfg["year_sheets"][0] in wb.sheetnames else 6
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    def sym(v):
        if v is None or v == "":
            return "."
        if isinstance(v, bool):
            return "?"
        if isinstance(v, (int, float)):
            return "0" if v == 0 else "#"
        return "?"

    lines = []
    hs = wb[wcfg["helper_sheet"]] if wcfg["helper_sheet"] in wb.sheetnames else None
    for sheet in wcfg["year_sheets"]:
        if sheet not in wb.sheetnames:
            lines.append("%s: sheet missing" % sheet)
            continue
        ws = wb[sheet]
        lines.append("%s    %s" % (sheet, " ".join("%3s" % m for m in months)))
        rows = wcfg["rows"]
        for label, row in (("gross (row %d)" % rows["btl_gross"], rows["btl_gross"]),
                           ("198+197 (row %d)" % rows["current_payment"], rows["current_payment"]),
                           ("193 (row %d)" % rows["completion"], rows["completion"]),
                           ("NI paid/day (row %d)" % rows["nis_paid_value"], rows["nis_paid_value"]),
                           ("NI paid days (row %d)" % rows["nis_paid_days"], rows["nis_paid_days"]),
                           ("days by 3010 (row %d)" % rows["reserve_days"], rows["reserve_days"])):
            cells = [ws["%s%d" % (get_column_letter(first + i), row)].value for i in range(12)]
            lines.append("  %-22s %s" % (label, " ".join("%3s" % sym(v) for v in cells)))
        if hs is not None:
            var = {}
            for r in range(1, hs.max_row + 1):
                d = _date_of(hs["%s%d" % (wcfg["helper_date_column"], r)].value)
                if d and str(d.year) == sheet:
                    var[d.month] = hs["%s%d" % (wcfg["helper_value_column"], r)].value
            lines.append("  %-22s %s" % ("variable (helper E)", " ".join("%3s" % sym(var.get(m + 1)) for m in range(12))))
        lines.append("  %s (current day value): %s" % (wcfg["current_value_cell"], sym(ws[wcfg["current_value_cell"]].value)))
    lines.append("legend: # number, 0 zero, . empty, ? text or formula (no figures shown)")
    return lines


def count_cells(messages):
    """(written, skipped) cell counts from the lines write_slip reports."""
    written = sum(1 for m in messages if m.lstrip().startswith("wrote"))
    skipped = sum(1 for m in messages if m.lstrip().startswith("SKIPPED"))
    return written, skipped


# ------------------------------------- existing result files: compatibility check, hidden backup, restore
class IncompatibleResult(Exception):
    """The chosen result file does not match the built-in template. `problems` explains why (Hebrew)."""

    def __init__(self, problems):
        super().__init__("; ".join(problems[:3]))
        self.problems = problems


def _template_strings(wb, cfg):
    """{(sheet, cell): text} of the fixed text cells (labels, headers) of the sheets the program relies on."""
    wcfg = cfg["workbook"]
    out = {}
    for name in list(wcfg["year_sheets"]) + [wcfg["helper_sheet"]]:
        if name not in wb.sheetnames:
            continue
        for row in wb[name].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value and not c.value.startswith("="):
                    out[(name, c.coordinate)] = c.value
    return out


def check_result_compatibility(path, cfg, max_problems=8):
    """Does an existing workbook have the structure of the built-in template? Returns (problems, hidden_count);
    an empty list means it can be updated safely."""
    from openpyxl import load_workbook

    try:
        existing = load_workbook(str(path))
    except Exception:
        return ["הקובץ אינו קובץ Excel תקין מסוג xlsx, או שאי אפשר לפתוח אותו."], 0
    template = open_template(None, with_summary=False)
    wcfg = cfg["workbook"]
    problems = []
    for name in list(wcfg["year_sheets"]) + [wcfg["helper_sheet"]]:
        if name not in existing.sheetnames:
            problems.append('חסר הגיליון "%s".' % name)
    wanted = _template_strings(template, cfg)
    summary_label = (wcfg["helper_sheet"], cfg["summary"]["label_range"].split(":")[0])
    wanted.pop(summary_label, None)                     # the summary block's wording is the user's to change
    for (sheet, coord), text in sorted(wanted.items()):
        if sheet not in existing.sheetnames:
            continue
        value = existing[sheet][coord].value
        if value != text:
            shown = "ריק" if value in (None, "") else '"%s"' % str(value)[:30]
            problems.append('%s!%s: בתבנית כתוב "%s", בקובץ %s.' % (sheet, coord, text[:30], shown))
    return problems[:max_problems], max(0, len(problems) - max_problems)


def old_layout_of(path, cfg):
    """The entry of config.json "old_layouts" whose marker labels the file has (a result made by an older version of the
    program), or None."""
    from openpyxl import load_workbook

    try:
        wb = load_workbook(str(path), read_only=True)
    except Exception:
        return None
    try:
        for layout in cfg["workbook"].get("old_layouts", []):
            sheets = [y for y in cfg["workbook"]["year_sheets"] if y in wb.sheetnames]
            if sheets and all(wb[y][a].value == v for y in sheets for a, v in layout["markers"].items()):
                return layout
        return None
    finally:
        wb.close()


def migrate_old_layout(path, layout, cfg):
    """A fresh workbook from the built-in template with every input of an older-layout result copied over: the slip
    rows, the reserve days, the current day value, and the helper sheet's unfixed gross (by month) and yearly limits.
    Formulas and labels come from the template. Nothing is saved here."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter, column_index_from_string

    old = load_workbook(str(path))
    wb = open_template(None, cfg=cfg)
    wcfg = cfg["workbook"]
    first = column_index_from_string(wcfg["first_month_column"])
    for y in wcfg["year_sheets"]:
        if y not in old.sheetnames or y not in wb.sheetnames:
            continue
        for key, old_row in layout["rows"].items():
            for m in range(12):
                col = get_column_letter(first + m)
                v = old[y]["%s%d" % (col, old_row)].value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    wb[y]["%s%d" % (col, wcfg["rows"][key])].value = v
        v = old[y][layout["current_value_cell"]].value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            wb[y][wcfg["current_value_cell"]].value = v
    h = wcfg["helper_sheet"]
    if h in old.sheetnames and h in wb.sheetnames:
        new_rows = {}
        for r in range(1, wb[h].max_row + 1):
            d = _date_of(wb[h]["%s%d" % (wcfg["helper_date_column"], r)].value)
            if d:
                new_rows[(d.year, d.month)] = r
        for r in range(1, old[h].max_row + 1):
            d = _date_of(old[h]["%s%d" % (wcfg["helper_date_column"], r)].value)
            v = old[h]["%s%d" % (wcfg["helper_value_column"], r)].value
            if d and (d.year, d.month) in new_rows and isinstance(v, (int, float)):
                wb[h]["%s%d" % (wcfg["helper_value_column"], new_rows[(d.year, d.month)])].value = v
        for rng in wcfg.get("editable_helper_ranges", []):
            for row in old[h][rng]:
                for c in row:
                    if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                        wb[h][c.coordinate].value = c.value
    return wb


def backup_path(result):
    """Hidden backup next to the result file: .<name>.bkup"""
    p = Path(result)
    return p.with_name("." + p.name + ".bkup")


def _set_hidden(path, hidden):
    if os.name != "nt":
        return                                          # a leading dot already hides it
    import ctypes

    ctypes.windll.kernel32.SetFileAttributesW(str(path), 2 if hidden else 0x80)   # HIDDEN / NORMAL


def _write_hidden(path, data):
    path = Path(path)
    if path.exists():
        _set_hidden(path, False)                        # Windows refuses to overwrite a hidden file
    path.write_bytes(data)
    _set_hidden(path, True)


def write_backup(result):
    """Keep a copy of the result file as it is now (one backup per file, the latest state before a run)."""
    b = backup_path(result)
    _write_hidden(b, Path(result).read_bytes())
    return b


def backup_info(result):
    """(exists, modified datetime or None) of the hidden backup."""
    b = backup_path(result)
    if not b.is_file():
        return False, None
    return True, dt.datetime.fromtimestamp(b.stat().st_mtime)


def restore_backup(result):
    """Put the backup back. The file being replaced becomes the new backup, so a second restore undoes the first.
    Returns True if a backup existed."""
    r, b = Path(result), backup_path(result)
    if not b.is_file():
        return False
    _set_hidden(b, False)
    previous = b.read_bytes()
    current = r.read_bytes() if r.is_file() else None
    r.write_bytes(previous)
    if current is None:
        b.unlink()
    else:
        _write_hidden(b, current)
    return True


_REF = re.compile(r"(?:'([^']+)'!)?\$?([A-Z]{1,2})\$?(\d+)(?::\$?([A-Z]{1,2})\$?(\d+))?")


def _number(value):
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _helper_value(hs, ref):
    """A number from the helper sheet: a plain value, or a formula of the form =N14/30."""
    v = hs[ref].value
    if isinstance(v, str):
        m = re.fullmatch(r"=\$?([A-Z]+)\$?(\d+)/(\d+)", v.replace(" ", ""))
        return _helper_value(hs, m.group(1) + m.group(2)) / float(m.group(3)) if m else None
    return _number(v)


def compute_total_owed(wb, cfg):
    """What Excel will show in "סה"כ ביטוח לאומי חייב לי" for this workbook, calculated here by following the template's own
    formulas (monthly figures, 3-month average capped by the yearly maximum, daily value floored by the yearly minimum,
    difference to the current day value times the reserve days). Returns {"total", "years": {year: annual difference}}, or
    None when the workbook is not shaped like the template."""
    from openpyxl.utils import column_index_from_string, get_column_letter

    wcfg = cfg["workbook"]
    if wcfg["helper_sheet"] not in wb.sheetnames:
        return None
    hs = wb[wcfg["helper_sheet"]]
    # helper table: month -> weighted unfixed gross (SUMIFS over this month and the 11 after it, / 12)
    dates = {}
    for row in range(1, hs.max_row + 1):
        d = _date_of(hs[wcfg["helper_date_column"] + str(row)].value)
        if d:
            dates[row] = (d.year * 12 + d.month, _number(hs[wcfg["helper_value_column"] + str(row)].value) or 0.0)
    weighted = {row: sum(v for _, (n, v) in dates.items() if n0 <= n <= n0 + 11) / 12 for row, (n0, _) in dates.items()}
    rows = wcfg["rows"]
    first = column_index_from_string(wcfg["first_month_column"])
    sheets = [y for y in wcfg["year_sheets"] if y in wb.sheetnames]
    r13, r18, r24 = {}, {}, {}

    def cell(sheet, col, row):
        return _number(wb[sheet].cell(row=row, column=col).value) or 0.0

    for y in sheets:                                       # month values (row 13) first: later months look back at them
        ws = wb[y]
        for m in range(12):
            col = first + m
            if not str(ws.cell(row=13, column=col).value or "").startswith("="):
                continue
            ref = _REF.search(str(ws.cell(row=10, column=col).value or ""))
            fixed = weighted.get(int(ref.group(3)), 0.0) if ref else 0.0
            r13[(y, col)] = (cell(y, col, rows["btl_gross"]) + fixed
                             - (cell(y, col, rows["current_payment"]) + cell(y, col, rows["completion"])))
    gaps = set()                                           # (year, month) of reserve months whose look-back has an empty month
    unpaid, overridden = set(), set()
    daily_of = {}                                          # (year, month) -> row 18, the correct daily value
    totals = {}
    for y in sheets:
        ws, annual = wb[y], 0.0
        day_value = _number(ws[wcfg["current_value_cell"]].value) or 0.0
        for m in range(12):
            col = first + m
            f17 = str(ws.cell(row=17, column=col).value or "")
            f18 = str(ws.cell(row=18, column=col).value or "")
            if not (f17.startswith("=") and f18.startswith("=")):
                continue
            hi = [h for h in re.findall(r"'[^']+'!\$([A-Z]+)\$(\d+)", f17) if h]
            lo = re.findall(r"'[^']+'!\$([A-Z]+)\$(\d+)", f18)
            cap = _helper_value(hs, hi[0][0] + hi[0][1]) if hi else None
            floor = _helper_value(hs, lo[0][0] + lo[0][1]) if lo else None
            average_args = re.search(r"AVERAGE\((.*?)\)(?:,|\)$)", f17)
            values = []
            for sheet, c1, r1, c2, r2 in _REF.findall(average_args.group(1) if average_args else ""):
                a, b = column_index_from_string(c1), column_index_from_string(c2 or c1)
                values += [r13[(sheet or y, c)] for c in range(a, b + 1) if (sheet or y, c) in r13]
            if not values or cap is None or floor is None:
                return None
            daily = max(min(sum(values) / len(values), cap) / 30, floor)
            daily_of[(int(y), m + 1)] = daily
            days = cell(y, col, rows["reserve_days"])
            paid_value = _number(ws.cell(row=rows["nis_paid_value"], column=col).value) if "nis_paid_value" in rows else None
            if day_value > 0:                                  # IF(N($F$36)>0, F18-$F$36, F18-F22)
                annual += days * (daily - day_value)
                if paid_value:
                    overridden.add(int(y))
            else:
                annual += days * (daily - (paid_value or 0.0))
                if days and not paid_value:
                    unpaid.add((int(y), m + 1))
            if days:
                for sheet, c1, r1, c2, r2 in _REF.findall(average_args.group(1)):
                    for c in range(column_index_from_string(c1), column_index_from_string(c2 or c1) + 1):
                        if wb[sheet or y].cell(row=rows["btl_gross"], column=c).value in (None, ""):
                            gaps.add((int(sheet or y), c - first + 1))
        totals[y] = annual
    return {"total": sum(totals.values()), "years": totals, "missing_months": sorted(gaps),
            "unpaid_months": sorted(unpaid), "overridden_years": sorted(overridden), "daily": daily_of}


# ------------------------------------------- service-period bonus: length of a continuous period mod 7
def service_periods(days):
    """Continuous periods [(start, end)] from calendar days: days that follow each other without a gap are one period
    (20-21.09 and 22-24.09 are 20-24.09), also across months and years."""
    periods = []
    for day in sorted(set(days)):
        if periods and day - periods[-1][1] == dt.timedelta(days=1):
            periods[-1][1] = day
        elif not periods or day > periods[-1][1]:
            periods.append([day, day])
    return [tuple(p) for p in periods]


def bonus_factor(days, rate=0.4):
    """Day values added to a continuous period: remainder 1-5 of days / 7 -> rate x remainder; 6 -> one day value;
    0 -> nothing. A 12-day period (remainder 5) gets 0.4 x 5 = 2 day values."""
    rem = days % 7
    return 1.0 if rem == 6 else round(rate * rem, 6)


def bonus_rows(periods, daily, cfg):
    """One dict per period. The whole period is valued by the daily value (row 18) of the month it STARTS in, and its
    bonus is counted in that month. daily: {(year, month): value} (None/missing = no value for that month)."""
    rate = cfg.get("bonus", {}).get("rate", 0.4)
    out = []
    for start, end in periods:
        n = (end - start).days + 1
        value = daily.get((start.year, start.month))
        factor = bonus_factor(n, rate)
        out.append({"start": start, "end": end, "days": n, "remainder": n % 7, "factor": factor,
                    "year": start.year, "month": start.month, "daily": value,
                    "amount": factor * value if value is not None else None})
    return out


def bonus_summary(rows):
    months, years = {}, {}
    for r in rows:
        if r["amount"] is None:
            continue
        months[(r["year"], r["month"])] = months.get((r["year"], r["month"]), 0.0) + r["amount"]
        years[r["year"]] = years.get(r["year"], 0.0) + r["amount"]
    return {"rows": rows, "months": months, "years": years, "total": round(sum(years.values()), 2),
            "no_value": [r for r in rows if r["amount"] is None]}


def write_bonus_sheet(wb, periods, cfg):
    """(Re)build the separate sheet of the service-period bonus: one row per continuous period, then a summary per month
    and per year and the grand total. Live formulas: the daily value is read from row 18 of the start month's year sheet.
    It is NOT part of "סה"כ ביטוח לאומי חייב לי". Only the start/end dates are editable."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.worksheet.table import Table, TableStyleInfo

    bcfg, wcfg = cfg.get("bonus", {}), cfg["workbook"]
    name = bcfg.get("sheet", "תוספת לתקופות שירות")
    rate = bcfg.get("rate", 0.4)
    if name in wb.sheetnames:
        del wb[name]
    at = wb.sheetnames.index(wcfg["helper_sheet"]) + 1 if wcfg["helper_sheet"] in wb.sheetnames else len(wb.sheetnames)
    ws = wb.create_sheet(name, at)
    ws.sheet_view.rightToLeft = True
    first = column_index_from_string(wcfg["first_month_column"])
    bold, head_fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    money, edge = "[$₪-40D]\\ #,##0.00", Side(style="thin", color="9BB7D4")
    border = Border(top=edge, bottom=edge, left=edge, right=edge)

    ws["B2"] = "תוספת לתקופות שירות (שארית ימי התקופה חלקי 7)"
    ws["B2"].font = Font(bold=True, size=14)
    ws["B3"] = ("תקופות רצופות מאוחדות לתקופה אחת. שארית 1–5: תוספת של %s × השארית ערכי יום. שארית 6: ערך יום אחד נוסף. "
                "שארית 0: אין תוספת. ערך היום הוא הערך היומי המחושב (שורה 18) של חודש תחילת התקופה, והתוספת נספרת בחודש זה. "
                "החישוב נפרד ואינו נכלל בסה״כ ביטוח לאומי חייב לי." % ("{:g}".format(rate)))
    ws["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B3:J3")
    ws.row_dimensions[3].height = 48

    heads = ["תחילת תקופה", "סוף תקופה", "ימים", "שארית (חלקי 7)", "ערכי יום נוספים", "חודש החישוב", "שנה",
             "ערך יומי (שורה 18)", "סכום התוספת"]
    top = 5
    for i, h in enumerate(heads):
        c = ws.cell(row=top, column=2 + i, value=h)
        c.font, c.fill, c.border = bold, head_fill, border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[top].height = 30
    rows = list(periods) or [(None, None)]
    for k, (start, end) in enumerate(rows):
        r = top + 1 + k
        ws.cell(row=r, column=2, value=start).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=3, value=end).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=4, value='=IF(AND(ISNUMBER(B{0}),ISNUMBER(C{0})),C{0}-B{0}+1,"")'.format(r))
        ws.cell(row=r, column=5, value='=IF(D{0}="","",MOD(D{0},7))'.format(r))
        ws.cell(row=r, column=6, value='=IF(E{0}="","",IF(E{0}=6,1,{1}*E{0}))'.format(r, rate))
        ws.cell(row=r, column=7, value='=IF(ISNUMBER(B{0}),DATE(YEAR(B{0}),MONTH(B{0}),1),"")'.format(r)).number_format = "mm/yyyy"
        ws.cell(row=r, column=8, value='=IF(ISNUMBER(B{0}),YEAR(B{0}),"")'.format(r))
        # the start month's row 18, found by year and month (CHOOSE over the year sheets, INDEX over the 12 months)
        sheets = [y for y in wcfg["year_sheets"] if y in wb.sheetnames]
        pick = ",".join("'{0}'!${1}$18:${2}$18".format(y, get_column_letter(first), get_column_letter(first + 11)) for y in sheets)
        found = "INDEX(CHOOSE(MATCH(YEAR(B{0}),{{{1}}},0),{2}),1,MONTH(B{0}))".format(r, ",".join(sheets), pick)
        ws.cell(row=r, column=9, value='=IFERROR(IF(ISNUMBER(B{0}),IF(N({1})>0,{1},""),""),"")'.format(r, found)).number_format = money
        ws.cell(row=r, column=10, value='=IF(ISNUMBER(I{0}),F{0}*I{0},"")'.format(r)).number_format = money
        for c in range(2, 11):
            ws.cell(row=r, column=c).border = border
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    last = top + len(rows)
    ref = "B%d:J%d" % (top, last)
    tbl = Table(displayName="tbl_bonus_periods", ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(tbl)

    # summary per month and per year, and the grand total (to the left of the periods)
    L = 13                                                            # column M
    ws.cell(row=top - 1, column=L, value="סיכום לפי חודש (חודש תחילת התקופה)").font = bold
    for i, h in enumerate(("חודש", "סכום התוספת")):
        c = ws.cell(row=top, column=L + i, value=h)
        c.font, c.fill, c.border = bold, head_fill, border
    r = top + 1
    amount_rng, month_rng, year_rng = "$J$%d:$J$%d" % (top + 1, last), "$G$%d:$G$%d" % (top + 1, last), "$H$%d:$H$%d" % (top + 1, last)
    for y in wcfg["year_sheets"]:
        for m in range(1, 13):
            ws.cell(row=r, column=L, value=dt.date(int(y), m, 1)).number_format = "mm/yyyy"
            ws.cell(row=r, column=L + 1, value="=SUMIFS(%s,%s,%s%d)" % (amount_rng, month_rng, get_column_letter(L), r)).number_format = money
            for c in (L, L + 1):
                ws.cell(row=r, column=c).border = border
            r += 1
    Y = L + 3                                                         # column P
    ws.cell(row=top - 1, column=Y, value="סיכום לפי שנה").font = bold
    for i, h in enumerate(("שנה", "סכום התוספת")):
        c = ws.cell(row=top, column=Y + i, value=h)
        c.font, c.fill, c.border = bold, head_fill, border
    for k, y in enumerate(wcfg["year_sheets"]):
        ws.cell(row=top + 1 + k, column=Y, value=int(y)).border = border
        c = ws.cell(row=top + 1 + k, column=Y + 1, value="=SUMIFS(%s,%s,%s%d)" % (amount_rng, year_rng, get_column_letter(Y), top + 1 + k))
        c.number_format, c.border = money, border
    t = top + 2 + len(wcfg["year_sheets"])
    lab = ws.cell(row=t, column=Y, value="סה\"כ תוספת לכל התקופות")
    tot = ws.cell(row=t, column=Y + 1, value="=SUM(%s)" % amount_rng)
    for c, color in ((lab, "4A90E2"), (tot, "70AD47")):
        c.fill, c.font, c.border = PatternFill("solid", fgColor=color), Font(bold=True, size=12), border
    tot.number_format = money
    for col, width in zip("BCDEFGHIJ", (13, 13, 7, 10, 11, 11, 7, 15, 15)):
        ws.column_dimensions[col].width = width
    for col, width in (("M", 11), ("N", 15), ("P", 20), ("Q", 15)):
        ws.column_dimensions[col].width = width
    return name


def preview_total_owed(slips, base_path, overwrite, cfg, reserve_months=None, current_day_value=None, nis_months=None, periods=None):
    """The "סה"כ ביטוח לאומי חייב לי" amount the result file would show after a fill, without writing anything.
    base_path: the result file the user chose (used when it exists and matches the template), else the built-in template."""
    wb = None
    if base_path and Path(base_path).exists():
        problems, _hidden = check_result_compatibility(Path(base_path), cfg)
        old = old_layout_of(Path(base_path), cfg) if problems else None
        if not problems:
            wb = open_template(Path(base_path))
        elif old:
            wb = migrate_old_layout(Path(base_path), old, cfg)
    if wb is None:
        wb = open_template(None)
    apply_to_workbook(wb, slips, overwrite, cfg, reserve_months, current_day_value, nis_months)
    res = compute_total_owed(wb, cfg)
    if res is not None and periods:
        res["bonus"] = bonus_summary(bonus_rows(periods, res["daily"], cfg))   # separate: not part of res["total"]
    return res


SHEET_PASSWORD_ENV = "RESERVE_DUTY_SHEET_PASSWORD"


def input_cells(wb, cfg):
    """{sheet title: set of cell addresses} the user types into; every other cell is locked."""
    from openpyxl.utils import get_column_letter, column_index_from_string

    wcfg = cfg["workbook"]
    first = column_index_from_string(wcfg["first_month_column"])
    cells = {}
    for y in wcfg["year_sheets"]:
        if y in wb.sheetnames:
            mine = cells.setdefault(y, {wcfg["current_value_cell"]})
            for row in wcfg["rows"].values():
                mine.update("%s%d" % (get_column_letter(first + m), row) for m in range(12))
    if wcfg["helper_sheet"] in wb.sheetnames:
        hs = wb[wcfg["helper_sheet"]]
        mine = cells.setdefault(hs.title, set())
        for r in range(1, hs.max_row + 1):
            if _date_of(hs["%s%d" % (wcfg["helper_date_column"], r)].value):
                mine.add("%s%d" % (wcfg["helper_value_column"], r))
        for rng in wcfg.get("editable_helper_ranges", []):          # the yearly maximum/minimum tables
            mine.update(c.coordinate for row in hs[rng] for c in row)
    bonus = cfg.get("bonus", {}).get("sheet")
    if bonus in wb.sheetnames:                                      # the service periods' start and end dates
        bs = wb[bonus]
        mine = cells.setdefault(bonus, set())
        for r in range(6, bs.max_row + 1):
            if _date_of(bs.cell(row=r, column=2).value) or _date_of(bs.cell(row=r, column=3).value):
                mine.update(("B%d" % r, "C%d" % r))
    return cells


def add_credit(wb, cfg):
    """Author and license in the file properties, the disclaimer as the first sheet (the one that opens), and both in the
    printed footer of every sheet."""
    credit, disc = cfg.get("credit", {}), cfg.get("disclaimer", {})
    if credit.get("author"):
        wb.properties.creator = credit["author"]
    if disc.get("he"):
        write_disclaimer_sheet(wb, disc)
    footer = " · ".join(t for t in (credit.get("footer"), disc.get("footer")) if t)
    if footer:
        for ws in wb.worksheets:
            ws.oddFooter.center.text = footer
            ws.evenFooter.center.text = footer


def write_disclaimer_sheet(wb, disc):
    """(Re)build the disclaimer sheet, first in the workbook and active, so it is what a file opens on."""
    from openpyxl.styles import Alignment, Font, PatternFill

    name = disc.get("sheet", "הבהרה")
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, 0)
    ws.sheet_view.rightToLeft = True
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    wrap = Alignment(wrap_text=True, vertical="top", horizontal="right", readingOrder=2)
    row = 2
    for title, lines, rtl in ((disc.get("title_he"), disc.get("he", []), True), (disc.get("title_en"), disc.get("en", []), False)):
        cell = ws.cell(row=row, column=2, value=title)
        cell.font = Font(bold=True, size=14, color="9C0006")
        cell.fill = PatternFill("solid", fgColor="FFC7CE")
        cell.alignment = wrap if rtl else Alignment(wrap_text=True, vertical="top", horizontal="left", readingOrder=1)
        row += 1
        for line in lines:
            c = ws.cell(row=row, column=2, value=line)
            c.font = Font(size=12)
            c.alignment = wrap if rtl else Alignment(wrap_text=True, vertical="top", horizontal="left", readingOrder=1)
            ws.row_dimensions[row].height = 36
            row += 1
        row += 1
    for other in wb.worksheets:
        other.sheet_view.tabSelected = False
    ws.sheet_view.tabSelected = True
    wb.active = 0
    return name


def protect_formulas(wb, cfg, password=None):
    """Lock every cell and hide every formula from the formula bar; only the input cells stay editable.
    This is sheet protection, not encryption: it keeps the calculation out of sight and safe from accidental edits.
    The password comes from the environment variable RESERVE_DUTY_SHEET_PASSWORD (to be able to unprotect your own
    files); otherwise a random one is used and not kept."""
    import secrets

    from openpyxl.styles import Protection

    password = password or os.environ.get(SHEET_PASSWORD_ENV) or secrets.token_urlsafe(12)
    editable = input_cells(wb, cfg)
    for ws in wb.worksheets:
        mine = editable.get(ws.title, set())
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.protection = Protection(locked=True, hidden=True)
                elif c.coordinate in mine:
                    c.protection = Protection(locked=False, hidden=False)
        for addr in mine:                                           # input cells that are still empty
            ws[addr].protection = Protection(locked=False, hidden=False)
        ws.protection.sheet = True
        ws.protection.password = password
        ws.protection.formatColumns = False                         # widening a column is allowed
        ws.protection.formatRows = False


def apply_to_workbook(wb, slips, overwrite, cfg, reserve_months=None, current_day_value=None, nis_months=None, periods=None):
    """Put slips, the certificate's reserve days, the NI letter's payments and the current day value into an open
    workbook (nothing is saved). Returns (results, failures, reserve_results, day_value); day_value["nis"] lists the
    cells written from the NI letter."""
    results, failures, seen = [], 0, {}
    ordered = sorted(slips, key=lambda s: (s[1].get("year") or 0, s[1].get("month") or 0, Path(s[0]).name))
    for name, slip in ordered:
        short = Path(name).name
        res = {"name": short, "month": slip.get("month"), "year": slip.get("year"), "status": "ok", "messages": []}
        if "month" not in slip or "year" not in slip:
            res.update(status="skipped", headline="%s: skipped (month/year not found - it cannot be placed)" % short)
            failures += 1
        elif (slip["year"], slip["month"]) in seen:
            key = (slip["year"], slip["month"])
            res.update(status="skipped", headline="%s: skipped - same month %02d/%d as %s" % (short, key[1], key[0], seen[key]))
            failures += 1
        else:
            seen[(slip["year"], slip["month"])] = short
            res["headline"] = "%s  ->  %02d/%d" % (short, slip["month"], slip["year"])
            if not write_slip(wb, slip, cfg, overwrite, res["messages"]):
                res["status"] = "partial"
                failures += 1
            res["written"], res["skipped"] = count_cells(res["messages"])
        results.append(res)

    reserve_results = write_reserve_days(wb, reserve_months, cfg) if reserve_months else []
    nis_results = write_nis_payments(wb, nis_months, cfg) if nis_months else []

    entered = bool(current_day_value and current_day_value > 0)
    day_value = {"entered": entered, "value": current_day_value or 0,
                 "written": write_current_day_value(wb, current_day_value, cfg) if entered else []}
    day_value["missing_years"] = years_missing_day_value(wb, cfg)
    day_value["range_warning"] = day_value_range_warning(current_day_value, cfg)
    day_value["nis"] = nis_results
    day_value["bonus_sheet"] = write_bonus_sheet(wb, periods, cfg) if periods else None

    return results, failures, reserve_results, day_value


def fill_workbook(slips, out_path, overwrite, cfg, reserve_months=None, replace_incompatible=False, current_day_value=None,
                  nis_months=None, periods=None):
    """Write parsed slips into the result workbook and save it.

    slips: [(display_name, slip_dict)]; out_path: result file. An existing file that matches the built-in
    template is extended; one that does not raises IncompatibleResult unless `replace_incompatible` (then it
    is replaced by a fresh template). Before touching an existing file a hidden backup is written. `reserve_months` ({(year, month): days}, from the
    service certificate) is written to row 27, replacing what is there; `nis_months` ({(year, month): {days, amount}}, from the
    NI annual letter) to rows 22/23. `current_day_value` (> 0) is written to F36 of every year sheet; years with neither F36
    nor a value from the letter are reported.
    Returns (results, base_note, failures, reserve_results, day_value); results is one dict per slip: name, month,
    year, status (ok | partial | skipped), headline, messages; day_value is {"entered", "value", "written",
    "missing_years"}."""
    out_path = Path(out_path)
    if out_path.exists():
        problems, _hidden = check_result_compatibility(out_path, cfg)
        old = old_layout_of(out_path, cfg) if problems else None
        if problems and not old and not replace_incompatible:
            raise IncompatibleResult(problems)
        write_backup(out_path)                           # hidden .bkup: the file as it was before this run
        if old:
            base_note = ("base workbook: %s converted from the older layout (%s) - your inputs were copied; "
                         "a hidden backup of the old file was kept" % (out_path.name, old["name"]))
            wb = migrate_old_layout(out_path, old, cfg)
        elif problems:
            base_note = "base workbook: built-in template (the file did not match the template and was replaced; a hidden backup was kept)"
            wb = open_template(None)
        else:
            base_note = "base workbook: %s (existing result file, slips are added to it; a hidden backup was kept)" % out_path.name
            wb = open_template(out_path)
    else:
        base_note = "base workbook: built-in template"
        wb = open_template(None)
    wb.template = False

    results, failures, reserve_results, day_value = apply_to_workbook(wb, slips, overwrite, cfg, reserve_months, current_day_value,
                                                                      nis_months, periods)

    add_credit(wb, cfg)
    protect_formulas(wb, cfg)
    wb.calculation.fullCalcOnLoad = True
    wb.save(str(out_path))
    return results, base_note, failures, reserve_results, day_value


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="pay-slip PDF(s)")
    ap.add_argument("--out", help="result .xlsx (default: filled_reserve_duty.xlsx). An existing file that matches the "
                                  "built-in template is updated; a hidden backup (.<name>.bkup) is kept.")
    ap.add_argument("--replace-incompatible", action="store_true",
                    help="if the result file does not match the built-in template, replace it with a fresh one (backup kept)")
    ap.add_argument("--restore-backup", action="store_true",
                    help="the files given are result workbooks: put back their hidden backup (a second call undoes it)")
    ap.add_argument("--config", default=str(HERE / "config.json"))
    ap.add_argument("--overwrite", action="store_true", help="overwrite cells that already hold a different value")
    ap.add_argument("--current-day-value", metavar="AMOUNT",
                    help='ערך ימ"מ נוכחי: written to F36 of every year sheet; it overrides the values from the NI letter (empty or 0 = not written)')
    ap.add_argument("--debug", action="store_true", help="masked structure dump, no Excel output")
    ap.add_argument("--diagnose", action="store_true", help="content-free report of what kind of PDF each file is")
    ap.add_argument("--ocr", action="store_true",
                    help="read a PDF that is only a picture (the reserve-service certificate) with image recognition; needs the OCR add-on")
    ap.add_argument("--inspect", action="store_true", help="the files given are result workbooks: show which input cells hold a value (no figures)")
    ap.add_argument("--check", action="store_true", help="print only the extracted values (on screen), write nothing")
    ap.add_argument("--debug-text", action="store_true", help="like --debug but prints the real Hebrew words (digits still masked)")
    ap.add_argument("--month", type=int, help="override month (1-12)")
    ap.add_argument("--year", type=int, help="override year")
    ap.add_argument("--split", type=float, help="left-panel boundary as a fraction of page width (default from config.json)")
    args = ap.parse_args(argv)
    if args.debug_text:
        args.debug = True
    if args.check and args.debug:
        ap.error("--check cannot be combined with --debug / --debug-text")

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        day_value_input = parse_day_value(args.current_day_value)
    except ValueError as e:
        print("ERROR: %s" % e)
        return 2
    range_warning = day_value_range_warning(day_value_input, cfg)
    if range_warning:
        print("WARNING: %s" % day_value_range_text(range_warning))

    if args.restore_backup:
        rc = 0
        for p in args.pdfs:
            try:
                done = restore_backup(p)
            except PermissionError:
                print("%s: cannot write the file. Is it open in Excel?" % Path(p).name)
                rc = 2
                continue
            print("%s: %s" % (Path(p).name, "backup restored (run again to undo)" if done else "no backup found"))
            rc = rc or (0 if done else 1)
        return rc

    if args.inspect:
        for p in args.pdfs:
            print("== %s" % Path(p).name)
            if not Path(p).is_file():
                print("  ERROR: file not found")
                continue
            try:
                print("\n".join(inspect_workbook(p, cfg)))
            except Exception as e:                       # unreadable / not an Excel file
                print("  ERROR: could not open as a workbook (%s)" % type(e).__name__)
        return 0

    expanded = []
    for p in args.pdfs:
        if Path(p).is_dir():
            found = sorted(Path(p).glob("*.pdf"), key=lambda q: q.name.lower())
            if not found:
                print("WARNING: no PDF files in folder %s" % p)
            expanded += [str(q) for q in found]
        else:
            expanded.append(p)
    args.pdfs = expanded
    if not args.pdfs:
        print("ERROR: no pay-slips to process.")
        return 2
    if args.diagnose:
        for p in args.pdfs:
            diagnose(p)
        return 0
    if not (args.check or args.debug) and inside_program_folder(args.out or "filled_reserve_duty.xlsx"):
        print("ERROR: the result file would be written inside the program's own folder:")
        print("  %s" % Path(args.out or "filled_reserve_duty.xlsx").resolve())
        print("Result files contain real pay figures, and this folder may be zipped or published.")
        print("Choose another folder for the result file (for example Documents), with --out.")
        return 2
    miss = missing_modules(need_excel=not (args.check or args.debug))
    if miss:
        print("ERROR: missing libraries in this Python (%s): %s" % (sys.executable, ", ".join(miss)))
        print("Install them with:  %s" % install_command())
        return 2
    print("%d pay-slip file(s) to process" % len(args.pdfs))

    missing = [p for p in args.pdfs if not Path(p).is_file()]
    if missing:
        for p in missing:
            print("ERROR: file not found: %s" % p)
        print("Tip: use the real path of your PDF, e.g. drag the file into the PowerShell window to paste its path.")
        return 2

    slips = []
    jobs = []  # (display name, pages)
    unreadable = 0
    forms = []                                           # reserve-service certificates (form 3010)
    letters = []                                         # NI annual letters (payments to the employer)
    for pdf in args.pdfs:
        try:
            pages = load_pages(pdf, cfg, ocr_allowed=args.ocr)
        except PdfReadError as e:
            print("== %s" % Path(pdf).name)
            print("  ERROR: %s" % e)
            unreadable += 1
            continue
        groups = []
        for kind, segment in split_by_kind(pages, cfg):
            if kind == "form":
                forms.append(parse_reserve_form(segment, cfg, Path(pdf).name))
            elif kind == "letter":
                letters.append(parse_nis_letter(segment, cfg, Path(pdf).name))
            else:
                groups += split_into_slips(segment, cfg)
        if len(groups) > 1:
            print("%s: found %d slips in this PDF" % (Path(pdf).name, len(groups)))
        for label, grp, note in groups:
            name = Path(pdf).name if len(groups) == 1 else "%s [%s]" % (Path(pdf).name, label)
            jobs.append((name, grp, note))

    for name, pdf_pages, note in jobs:
        print("== %s" % name)
        if note:
            print("  WARNING: %s" % note)
        pdf = name
        slip = parse_slip(pdf_pages, cfg, debug=args.debug, force_month=args.month, force_year=args.year, split_ratio=args.split, text_dump=args.debug_text)
        slips.append((pdf, slip))
        if args.check:
            fmt = lambda v: "NOT FOUND" if v is None else "{:,.2f}".format(v)
            print("  month/year             : %s/%s" % (slip.get("month"), slip.get("year")))
            print("  ברוטו לב.לאומי         : %s" % fmt(slip["btl_gross"]))
            print("  ברוטו בלתי קבוע (ב.ק.) : %s" % ("0.00 (not on slip)" if slip["btl_variable"] is None else fmt(slip["btl_variable"])))
            print("  קוד %s (%d rows) : %s" % (slip["current_payment_codes"], slip["current_payment_rows"], fmt(slip["current_payment"])))
            print("  קוד %s (%d rows) : %s" % (slip["completion_codes"], slip["completion_rows"], fmt(slip["completion"])))
            for w in slip["warnings"]:
                print("  WARNING:", w)
            continue
        if args.debug:
            print("month/year found: %s/%s   digits reversed: %s" % (slip.get("month"), slip.get("year"), slip["digits_reversed"]))
            for w in slip["warnings"]:
                print("  WARNING:", w)
            continue
        print("  month %s/%s   198 rows: %d   193 rows: %d" % (slip.get("month"), slip.get("year"), slip["current_payment_rows"], slip["completion_rows"]))
        for w in slip["warnings"]:
            print("  WARNING:", w)
    reserve = combine_reserve(forms, {int(y) for y in cfg["workbook"]["year_sheets"]}) if forms else None
    if reserve:
        print("\n== reserve service certificate")
        print("\n".join(reserve_summary_lines(reserve)))
    nis = combine_letters(letters, {int(y) for y in cfg["workbook"]["year_sheets"]}) if letters else None
    if nis:
        print("\n== NI annual letter (payments to the employer)")
        print("\n".join(letter_summary_lines(nis)))
    reserve_issues = bool(reserve and (reserve["mismatches"] or reserve["warnings"])) or bool(nis and letter_problems(nis))
    periods = service_periods((reserve or {}).get("all_days") or (nis or {}).get("all_days") or [])
    if periods:
        print("\n== service periods (continuous periods joined; remainder of days / 7 -> bonus day values)")
        for st, en in periods:
            n = (en - st).days + 1
            print("  %s - %s: %d day(s), remainder %d, bonus %s day value(s) of %02d/%d"
                  % (st.strftime("%d/%m/%Y"), en.strftime("%d/%m/%Y"), n, n % 7, "{:g}".format(bonus_factor(n, cfg.get("bonus", {}).get("rate", 0.4))),
                     st.month, st.year))
    if args.debug or args.check:
        return 1 if (unreadable or reserve_issues) else 0
    out_path = Path(args.out or "filled_reserve_duty.xlsx")
    try:
        results, base_note, failures, reserve_results, day_info = fill_workbook(
            slips, out_path, args.overwrite, cfg, reserve["months"] if reserve else None, args.replace_incompatible,
            day_value_input, nis["months"] if nis else None, periods)
    except IncompatibleResult as e:
        print("ERROR: %s does not match the built-in template, so it was not touched:" % out_path.name)
        for problem in e.problems:
            print("  - %s" % problem)
        print("To replace it with a fresh file (a hidden backup is kept), run again with --replace-incompatible.")
        return 2
    except PermissionError:
        print("ERROR: cannot write %s. Is it open in Excel? Close it and try again." % out_path)
        return 2
    print("\n" + base_note)
    for res in results:
        print(res["headline"])
        for line in res["messages"]:
            print(line)
    for r in reserve_results:
        print("  wrote   ימי מילואים        %s = %s%s" % (r["cell"], r["value"], "   (replaced %s)" % r["previous"] if r["replaced"] else ""))
    for r in day_info["nis"]:
        print("  wrote   NI letter          %s = %s%s" % (r["cell"], r["value"], "   (replaced %s)" % r["previous"] if r["replaced"] else ""))
    for r in day_info["written"]:
        print("  wrote   ערך ימ\"מ נוכחי    %s = %s%s" % (r["cell"], r["value"], "   (replaced %s)" % r["previous"] if r["replaced"] else ""))
    if day_info["written"] and day_info["nis"]:
        print('WARNING: both a current day value (%s) and the NI letter were given. In the workbook the current day value wins '
              'for the whole year; clear %s to use the monthly values from the letter.' % ((cfg["workbook"]["current_value_cell"],) * 2))
    print("\nsaved: %s" % out_path.resolve())
    print("NOTE: personal interpretation of the law, not legal/accounting advice; not affiliated with National Insurance. "
          "Consult a qualified professional before acting on the results (see DISCLAIMER.md).")
    if day_info["missing_years"]:
        print('WARNING: no current day value and no NI letter for: %s. Add the NI annual letter, or type the value in cell %s '
              '(or row %d) by hand (the difference formulas depend on it).'
              % (", ".join(day_info["missing_years"]), cfg["workbook"]["current_value_cell"], cfg["workbook"]["rows"]["nis_paid_value"]))
    if not reserve:
        print("REMINDER: row %d (ימי מילואים ע\"פ 3010) is empty unless you fill it - add the reserve-service certificate (form 3010), "
              "or enter the days by hand in Excel. Without it the difference is 0." % cfg["workbook"]["rows"]["reserve_days"])
    return 1 if (failures or unreadable or reserve_issues) else 0


if __name__ == "__main__":
    sys.exit(main())
