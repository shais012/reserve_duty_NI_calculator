"""Local web interface for the pay-slip filler.

Starts a small web server that listens ONLY on this computer (127.0.0.1), opens the page in the default
browser and shuts itself down when the tab is closed. Nothing is sent anywhere: uploaded PDFs are copied to a
private temporary folder, which is deleted when the program ends.

Safety measures: loopback only, random port, secret per-session token on every request, Host / Origin checks
(no DNS-rebinding or cross-site calls), strict Content-Security-Policy (no external resources), uploads
validated (.pdf, %PDF header, size caps), result files refused inside the program folder.

Options (mostly for testing):  --port N   --no-browser   --print-url   --settings-file FILE
"""
import argparse
import hmac
import http.server
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import fill_reserve_duty as core  # noqa: E402

VERSION = "1.0"
MAX_FILE_BYTES = 200 * 1024 * 1024
MAX_TOTAL_BYTES = 2 * 1024 * 1024 * 1024
MAX_FILES = 300
DENIED_BODY_LIMIT = 4 * 1024 * 1024        # most a rejected request may make us read before answering
IDLE_SECONDS = 90
TEMP_PREFIX = "reserve_duty_web_"
UPLOAD_KINDS = ("slips", "form", "letter")   # pay-slips / reserve-service certificate (form 3010) / NI annual letter
RESULT_SUFFIXES = (".xlsx",)


class UserError(Exception):
    """A problem the user can fix; the message (Hebrew) is shown in the page."""


SETTINGS_FILE = Path.home() / ".reserve_duty_web.json"      # remembered choices: in the user's profile, never next to the code


def load_settings():
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(**changes):
    data = load_settings()
    data.update(changes)
    try:
        SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass                                                  # remembering a choice is a convenience, never an error


def default_result_path():
    """Where the result goes unless the user chose otherwise: the file chosen last time, else Documents."""
    saved = load_settings().get("result")
    if saved:
        p = Path(saved)
        if p.parent.is_dir() and not core.inside_program_folder(p):
            return p
    docs = Path.home() / "Documents"
    return (docs if docs.is_dir() else Path.home()) / "filled_reserve_duty.xlsx"


DIALOG_SCRIPT = """
import sys, tkinter
from tkinter import filedialog
sys.stdout.reconfigure(encoding="utf-8")
root = tkinter.Tk()
root.withdraw()
root.attributes("-topmost", True)
path = filedialog.asksaveasfilename(title=sys.argv[3], initialdir=sys.argv[1], initialfile=sys.argv[2],
                                    defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], confirmoverwrite=False)
sys.stdout.write(path or "")
"""


def run_save_dialog(start_path):
    """Show the Windows 'save as' dialog in a helper process (a browser cannot give a real path).
    Picking an existing file selects it for updating. Returns the chosen path, or None if cancelled."""
    start = Path(start_path)
    args = [core.console_python(), "-c", DIALOG_SCRIPT, str(start.parent if start.parent.is_dir() else Path.home()),
            start.name, "בחר איפה לשמור את קובץ התוצאה, או קובץ קיים לעדכון"]
    try:
        done = subprocess.run(args, capture_output=True, timeout=600, creationflags=0x08000000 if os.name == "nt" else 0)  # no console window
    except (OSError, subprocess.TimeoutExpired):
        raise UserError("לא ניתן לפתוח את חלון בחירת הקובץ. אפשר להקליד את הנתיב ידנית.")
    return done.stdout.decode("utf-8", "replace").strip() or None


class NeedsConfirmation(Exception):
    """The chosen result file does not match the template; the user must decide whether to replace it."""

    def __init__(self, problems):
        super().__init__("incompatible result file")
        self.problems = problems


def hebrew_pdf_error(message):
    """The engine's messages are English (command line); the page shows Hebrew."""
    if "install the OCR add-on" in message:
        return ("ה-PDF לא מכיל טקסט (רק תמונות של העמודים), ולכן צריך את תוסף זיהוי התמונה (OCR) כדי לקרוא אותו. "
                "בטרמינל (PowerShell) הרץ: %s  ואז הפעל את התוכנה מחדש." % core.ocr_install_command())
    if "found no reserve-service table" in message:
        return "ה-PDF מכיל רק תמונות של העמודים, וזיהוי התמונה (OCR) לא מצא בו טבלה של אישור שירות מילואים. בדוק שזה הקובץ הנכון ושהסריקה ברורה."
    if message.startswith("The PDF is password-protected"):
        return ("ה-PDF מוגן בסיסמה. פתח אותו, ושמור או הדפס אותו ל-PDF חדש בלי סיסמה (העותק נשאר במחשב שלך), "
                "ואז הוסף את העותק.")
    if message.startswith("No text could be read"):
        return ("לא נמצא טקסט ב-PDF. תלוש שכר שה-PDF שלו הוא תמונה סרוקה (בלי טקסט) לא נתמך. "
                "אם זה אישור שירות מילואים, הוסף אותו בתיבת האישור.")
    return message


def safe_display_name(raw):
    name = os.path.basename(urllib.parse.unquote(raw or "").replace("\\", "/"))
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip()
    return name[:120] or "slip.pdf"


class App:
    """Session state: token, private upload folder, uploaded files, last result."""

    def __init__(self, session_dir=None):
        self.token = secrets.token_urlsafe(24)
        self.session_dir = Path(session_dir) if session_dir else Path(tempfile.mkdtemp(prefix=TEMP_PREFIX))
        self.cfg = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
        self.uploads = []            # {"id","name","path","size"}
        self.parsed = {}             # id -> {"groups": [...]} or {"error": str}
        self.cache_lock = threading.Lock()    # only guards parse_locks; never held while reading a file
        self.parse_locks = {}        # id -> lock held while that file is being read
        self.last_result = None
        self.last_seen = time.time()
        self.lock = threading.RLock()
        self._next_id = 1
        self.port = None
        self.httpd = None

    # ---- uploads
    def total_bytes(self):
        return sum(u["size"] for u in self.uploads)

    def new_upload(self, name, kind="slips"):
        with self.lock:
            uid = self._next_id
            self._next_id += 1
        return {"id": uid, "name": name, "kind": kind, "path": self.session_dir / ("u%d.pdf" % uid), "size": 0}

    def add_upload(self, up):
        with self.lock:
            self.uploads.append(up)
        threading.Thread(target=self.parse_upload, args=(up,), name="parse-upload", daemon=True).start()   # read it while the user keeps working

    def remove_upload(self, uid):
        with self.lock:
            for u in list(self.uploads):
                if u["id"] == uid:
                    self.uploads.remove(u)
                    self.parsed.pop(uid, None)
                    with self.cache_lock:
                        lock = self.parse_locks.pop(uid, None)
                    try:
                        u["path"].unlink()
                    except OSError:                     # still being read in the background: delete when it is done
                        threading.Thread(target=self._delete_when_free, args=(u["path"], lock), daemon=True).start()

    @staticmethod
    def _delete_when_free(path, lock):
        with lock or threading.Lock():
            try:
                path.unlink()
            except OSError:
                pass

    def clear_uploads(self, kind=None):
        for u in list(self.uploads):
            if kind is None or u["kind"] == kind:
                self.remove_upload(u["id"])

    # ---- paths chosen by the user
    def resolve_result(self, text):
        text = (text or "").strip()
        p = Path(text) if text else default_result_path()
        if not p.is_absolute() and p.parent == Path("."):
            p = default_result_path().parent / p.name
        if p.suffix.lower() not in RESULT_SUFFIXES:
            p = p.with_name(p.name + ".xlsx")
        p = p.resolve()
        if core.inside_program_folder(p):
            raise UserError("קובץ התוצאה מכיל נתוני שכר אמיתיים ולכן אסור לשמור אותו בתיקיית התוכנה. בחר תיקייה אחרת, למשל מסמכים.")
        if self.session_dir in p.parents:
            raise UserError("אי אפשר לשמור את התוצאה בתיקייה הזמנית של התוכנה.")
        if not p.parent.is_dir():
            raise UserError("התיקייה של קובץ התוצאה לא קיימת: %s" % p.parent)
        return p

    # ---- analysis
    def parse_upload(self, up):
        """Read one uploaded PDF once; later calls (or a call that arrives while it is being read) share the result."""
        with self.cache_lock:
            if up not in self.uploads:
                return None
            lock = self.parse_locks.setdefault(up["id"], threading.Lock())
        with lock:
            cached = self.parsed.get(up["id"])
            if cached is not None:
                return cached
            try:
                pages = core.load_pages(str(up["path"]), self.cfg, ocr_allowed=(up["kind"] in ("form", "letter")))
                groups, up_forms, up_letters = [], [], []
                for kind, segment in core.split_by_kind(pages, self.cfg):
                    if kind == "form":
                        up_forms.append(core.parse_reserve_form(segment, self.cfg, up["name"]))
                    elif kind == "letter":
                        up_letters.append(core.parse_nis_letter(segment, self.cfg, up["name"]))
                    else:
                        for label, grp, note in core.split_into_slips(segment, self.cfg):
                            groups.append((label, note, core.parse_slip(grp, self.cfg)))
                cached = {"groups": groups, "forms": up_forms, "letters": up_letters}
            except core.PdfReadError as e:
                cached = {"error": hebrew_pdf_error(str(e))}
            except Exception:                                       # damaged or unusual PDF
                cached = {"error": "לא ניתן לקרוא את הקובץ. ייתכן שהוא פגום או מוגן."}
            if up in self.uploads:                  # (no self.lock here: a run holding it may be waiting for this read)
                self.parsed[up["id"]] = cached
            return cached

    def analyze(self):
        """Parse every uploaded PDF. Returns (entries, file_errors, forms, letters): entries are dicts incl. the raw slip,
        forms the parsed reserve-service certificates (form 3010), letters the parsed NI annual letters."""
        entries, file_errors, names, forms, letters = [], [], {}, [], []
        year_sheets = set(self.cfg["workbook"]["year_sheets"])
        with self.lock:
            uploads = list(self.uploads)
        for up in uploads:
            cached = self.parse_upload(up)
            if cached is None:                                      # removed while waiting
                continue
            if "error" in cached:
                file_errors.append({"file": up["name"], "error": cached["error"]})
                continue
            if up["kind"] == "letter":
                if not cached["letters"]:
                    file_errors.append({"file": up["name"], "error": "לא זוהה בקובץ האישור השנתי של ביטוח לאומי (\"אישור שנתי בגין תשלומים שהועברו למעסיק\"). "
                                                                     "תלוש שכר מוסיפים בתיבת התלושים, וטופס 3010 בתיבת האישור."})
                    continue
                if cached["groups"] or cached["forms"]:
                    file_errors.append({"file": up["name"], "included": True,
                                        "error": "בקובץ יש גם עמודים אחרים (תלוש או טופס 3010). הם לא נכללו, כי הקובץ הועלה כאישור של ביטוח לאומי. האישור עצמו נקרא."})
                letters += cached["letters"]
                continue
            if cached["letters"] and not cached["groups"] and not cached["forms"]:
                file_errors.append({"file": up["name"], "error": "הקובץ נראה כאישור השנתי של ביטוח לאומי. הוסף אותו בתיבה של אישור ביטוח לאומי."})
                continue
            if up["kind"] == "form":
                if not cached["forms"]:
                    file_errors.append({"file": up["name"], "error": "לא זוהה בקובץ אישור שירות מילואים (טופס 3010). אם זה תלוש שכר, הוסף אותו בתיבת התלושים."})
                    continue
                if cached["groups"]:
                    file_errors.append({"file": up["name"], "included": True,
                                        "error": "בקובץ יש גם עמודים שנראים כתלוש שכר. הם לא נכללו, כי הקובץ הועלה כאישור מילואים. האישור עצמו נקרא."})
                forms += cached["forms"]
                continue
            if cached["forms"] and not cached["groups"]:
                file_errors.append({"file": up["name"], "error": "הקובץ נראה כאישור שירות מילואים (טופס 3010). הוסף אותו בתיבת האישור, לא בתיבת התלושים."})
                continue
            if cached["forms"]:
                file_errors.append({"file": up["name"], "included": True,
                                    "error": "בקובץ יש גם עמודים של אישור (3010 או ביטוח לאומי). הם לא נכללו, כי הקובץ הועלה כתלושים. התלושים עצמם נקראו."})
            for label, note, slip in cached["groups"]:
                display = up["name"] if len(cached["groups"]) == 1 else "%s [%s]" % (up["name"], label)
                names[display] = names.get(display, 0) + 1
                if names[display] > 1:
                    display = "%s (%d)" % (display, names[display])
                entries.append({"display": display, "slip": slip, "note": note})

        seen = {}
        for e in entries:
            s = e["slip"]
            issues, level = [], "ok"
            for w in s.get("warnings", []):
                issues.append(w)
                level = "warn"
            if e["note"]:
                issues.append(e["note"])
                level = "warn"
            if "month" not in s or "year" not in s:
                issues.append("לא נמצאו חודש ושנה בתלוש")
                level = "error"
            elif str(s["year"]) not in year_sheets:
                issues.append("אין גיליון לשנה %s בקובץ" % s["year"])
                level = "error"
            if s.get("btl_gross") is None:
                if "ברוטו לב.לאומי לא נמצא" not in " ".join(issues) and "ברוטו לב.לאומי not found" not in " ".join(issues):
                    issues.append("ברוטו לב.לאומי לא נמצא - שאר הערכים ייכתבו, ואת הברוטו יש להשלים ידנית באקסל")
                if level != "error":
                    level = "warn"
            if "month" in s and "year" in s:
                key = (s["year"], s["month"])
                if key in seen:
                    issues.append("אותו חודש כמו %s - הכפילות תידלג" % seen[key])
                    level = "error"
                else:
                    seen[key] = e["display"]
            e["issues"], e["level"] = issues, level
        return entries, file_errors, forms, letters

    def reserve_json(self, forms):
        """The service-certificate result for the page: per-month days with the periods behind them."""
        if not forms:
            return None
        years = {int(y) for y in self.cfg["workbook"]["year_sheets"]}
        res = core.combine_reserve(forms, years)
        iso = lambda d: d.strftime("%d/%m/%Y")
        return {
            "months": [{"year": y, "month": m, "days": d,
                        "parts": [{"start": iso(p["start"]), "end": iso(p["end"]), "days": p["days"]} for p in res["parts"].get(k, [])]}
                       for k in sorted(res["months"]) for (y, m), d in [(k, res["months"][k])]],
            "ignored": [{"year": y, "days": d} for y, d in sorted(res["ignored"].items())],
            "periods": res["rows"], "duplicates": res["duplicates"],
            "ocr": bool(res["ocr_files"]),
            "warnings": (
                ["חסרים עמודים %s מתוך %d. העלה את כל עמודי האישור בהרצה אחת, אחרת חלק מהחודשים יהיו חסרים או חלקיים."
                 % (", ".join(map(str, m["missing"])), m["total"]) for m in res["missing_pages"]]
                + (["נמצאו %d תקופות חופפות. כל יום נספר פעם אחת." % len(res["overlaps"])] if res["overlaps"] else [])),
            "mismatches": [{"start": iso(m["start"]), "end": iso(m["end"]), "stated": m["days"], "span": m["span"]} for m in res["mismatches"]],
            "_months": res["months"], "_days": res["all_days"],
        }

    def letter_json(self, letters):
        """The NI letter result for the page: per month the days paid, the amount and the value per day."""
        if not letters:
            return None
        years = {int(y) for y in self.cfg["workbook"]["year_sheets"]}
        res = core.combine_letters(letters, years)
        iso = lambda d: d.strftime("%d/%m/%Y")
        rows = self.cfg["workbook"]["rows"]
        warnings = ["חסרים עמודים %s מתוך %d של האישור. העלה את כל העמודים, אחרת חלק מהתשלומים חסרים."
                    % (", ".join(map(str, m["missing"])), m["total"]) for m in res["missing_pages"]]
        warnings += ["%s: סכום השורות שנקראו (%s) שונה מהסה\"כ באישור (%s). בדוק שכל העמודים נקראו."
                     % (t["file"], "{:,.2f}".format(t["read"]), "{:,.2f}".format(t["stated"])) for t in res["total_checks"]]
        warnings += ["%s – %s: כתוב %d ימים, ובלוח השנה %d. השורה נכללה כפי שהיא (זה מה ששולם), בדוק אותה."
                     % (iso(m["start"]), iso(m["end"]), m["days"], m["span"]) for m in res["mismatches"]]
        return {
            "months": [{"year": y, "month": m, "days": v["days"], "amount": v["amount"],
                        "value": round(v["amount"] / v["days"], 2) if v["days"] else None}
                       for (y, m), v in sorted(res["months"].items())],
            "ignored": [{"year": y, "amount": round(a, 2)} for y, a in sorted(res["ignored"].items())],
            "rows": res["rows"], "duplicates": res["duplicates"], "ocr": bool(res["ocr_files"]),
            "warnings": warnings, "value_row": rows["nis_paid_value"], "days_row": rows["nis_paid_days"],
            "_months": res["months"], "_days": res["all_days"],
        }

    @staticmethod
    def periods_of(reserve, letter):
        """Continuous service periods: from the certificate, else from the NI letter."""
        days = (reserve or {}).get("_days") or (letter or {}).get("_days") or []
        return core.service_periods(days)

    @staticmethod
    def bonus_json(bonus):
        if not bonus:
            return None
        iso = lambda d: d.strftime("%d/%m/%Y")
        return {
            "periods": [{"start": iso(r["start"]), "end": iso(r["end"]), "days": r["days"], "remainder": r["remainder"],
                         "factor": r["factor"], "year": r["year"], "month": r["month"],
                         "daily": round(r["daily"], 2) if r["daily"] is not None else None,
                         "amount": round(r["amount"], 2) if r["amount"] is not None else None} for r in bonus["rows"]],
            "months": [{"year": y, "month": m, "amount": round(a, 2)} for (y, m), a in sorted(bonus["months"].items())],
            "years": [{"year": y, "amount": round(a, 2)} for y, a in sorted(bonus["years"].items())],
            "total": bonus["total"], "no_value": len(bonus["no_value"]),
        }

    @staticmethod
    def entry_json(e):
        s = e["slip"]
        return {
            "name": e["display"], "month": s.get("month"), "year": s.get("year"),
            "gross": s.get("btl_gross"), "variable": s.get("btl_variable"),
            "payment": s.get("current_payment"), "payment_rows": s.get("current_payment_rows"),
            "topup": s.get("completion"), "topup_rows": s.get("completion_rows"),
            "issues": e["issues"], "level": e["level"],
        }

    def owed_json(self, entries, reserve, payload, day_value, letter=None, out=None):
        """The amount the result file will show as "סה"כ ביטוח לאומי חייב לי", worked out in advance (nothing is written)."""
        months = reserve["_months"] if reserve else {}
        nis_months = letter["_months"] if letter else {}
        if not entries and not months and not nis_months:
            return None
        try:
            base = self.resolve_result(payload.get("result"))
        except UserError:
            base = None                                         # a check does not need a valid result path
        try:
            res = core.preview_total_owed([(e["display"], e["slip"]) for e in entries], base, bool(payload.get("overwrite")),
                                          self.cfg, months, day_value, nis_months, self.periods_of(reserve, letter))
        except Exception:                                       # unusual workbook: the check itself must still work
            return None
        if not res:
            return None
        if out is not None:
            out["bonus"] = self.bonus_json(res.get("bonus"))                    # separate from the total owed
        return {"total": round(res["total"], 2), "years": {y: round(v, 2) for y, v in res["years"].items()},
                "day_value_entered": day_value > 0,
                "missing_months": [{"year": y, "month": m} for y, m in res["missing_months"]],
                "unpaid_months": [{"year": y, "month": m} for y, m in res["unpaid_months"]],
                "overridden_years": res["overridden_years"]}

    def run(self, payload):
        mode = payload.get("mode")
        if mode not in ("check", "fill"):
            raise UserError("פעולה לא מוכרת")
        if not self.uploads and mode == "check":
            raise UserError("לא נוספו קבצים. הוסף תלושי שכר, אישור מילואים, או את שניהם.")
        missing = core.missing_modules(need_excel=(mode == "fill"))
        if missing:
            raise UserError("חסרות ספריות בפייתון הזה: %s. להתקנה: %s" % (", ".join(missing), core.install_command()))
        result_path = None
        if mode == "fill":
            result_path = self.resolve_result(payload.get("result"))
        try:
            day_value = core.parse_day_value(payload.get("current_day_value"))
        except ValueError as e:
            raise UserError(str(e))
        entries, file_errors, forms, letters = self.analyze()
        reserve = self.reserve_json(forms)
        letter = self.letter_json(letters)
        out = {"mode": mode, "file_errors": file_errors, "slips": [self.entry_json(e) for e in entries],
               "day_value": {"entered": day_value > 0, "value": day_value, "written": [], "missing_years": [],
                             "range_warning": core.day_value_range_warning(day_value, self.cfg)}}
        if reserve:
            out["reserve"] = {k: v for k, v in reserve.items() if not k.startswith("_")}
        if letter:
            out["letter"] = {k: v for k, v in letter.items() if not k.startswith("_")}
            if not reserve:
                out["letter"]["no_certificate"] = True                          # row 27 stays empty: upload the 3010 or type it
            if day_value > 0:
                out["letter"]["overridden_by"] = self.cfg["workbook"]["current_value_cell"]
        out["owed"] = self.owed_json(entries, reserve, payload, day_value, letter, out)   # before a fill: from the file as it is now
        if mode == "fill":
            reserve_months = reserve["_months"] if reserve else {}
            try:
                results, base_note, failures, reserve_results, day_info = core.fill_workbook(
                    [(e["display"], e["slip"]) for e in entries], result_path, bool(payload.get("overwrite")), self.cfg,
                    reserve_months, bool(payload.get("replace_incompatible")), day_value,
                    letter["_months"] if letter else None, self.periods_of(reserve, letter))
            except core.IncompatibleResult as e:
                raise NeedsConfirmation(e.problems)
            except PermissionError:
                raise UserError("אי אפשר לכתוב לקובץ התוצאה. ייתכן שהוא פתוח באקסל: סגור אותו ונסה שוב.")
            if letter:
                out["letter"]["written"] = [{"cell": r["cell"], "value": r["value"], "previous": r["previous"], "replaced": r["replaced"]}
                                            for r in day_info.pop("nis", [])]
            else:
                day_info.pop("nis", None)
            if out.get("bonus") is not None:
                out["bonus"]["sheet"] = day_info.get("bonus_sheet")
            day_info.pop("bonus_sheet", None)
            out["day_value"] = day_info
            save_settings(result=str(result_path), overwrite=bool(payload.get("overwrite")),
                          **({"current_day_value": day_value} if day_value > 0 else {}))
            if reserve:
                out["reserve"]["written"] = [{"cell": r["cell"], "value": r["value"], "previous": r["previous"], "replaced": r["replaced"]}
                                             for r in reserve_results]
                if reserve["mismatches"] or reserve["warnings"]:
                    failures += 1
            by_name = {r["name"]: r for r in results}
            for slip_json, e in zip(out["slips"], entries):
                r = by_name.get(Path(e["display"]).name)
                if r:
                    slip_json["fill_status"], slip_json["messages"] = r["status"], r["messages"]
                    slip_json["written"], slip_json["skipped"] = r.get("written", 0), r.get("skipped", 0)
            out.update(saved=str(result_path), base_note=base_note, failures=failures,
                       existed=("existing result file" in base_note), migrated=("older layout" in base_note),
                       written=sum(r.get("written", 0) for r in results), skipped=sum(r.get("skipped", 0) for r in results))
            self.last_result = result_path
        return out


class Handler(http.server.BaseHTTPRequestHandler):
    server_version = "ReserveDuty"
    sys_version = ""

    def log_message(self, *args):
        pass

    @property
    def app(self):
        return self.server.app

    # ---- plumbing
    def _common_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Cross-Origin-Resource-Policy", "same-origin")
        self.send_header("Cross-Origin-Opener-Policy", "same-origin")

    def _json(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self._common_headers()
        self.end_headers()
        self.wfile.write(data)

    def _deny(self, status=403, message="forbidden"):
        self._json({"error": message}, status)

    def _host_ok(self):
        return self.headers.get("Host", "") in self.server.allowed_hosts

    def _origin_ok(self):
        origin = self.headers.get("Origin")
        if origin is not None and origin not in self.server.allowed_origins:
            return False
        site = self.headers.get("Sec-Fetch-Site")
        return site in (None, "same-origin", "none")

    def _token_ok(self, supplied):
        return supplied is not None and hmac.compare_digest(supplied.encode("utf-8"), self.app.token.encode("utf-8"))

    def _body_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > 1_000_000:
            self._drain_body(DENIED_BODY_LIMIT)
            raise UserError("בקשה גדולה מדי")
        raw = self.rfile.read(length) if length else b"{}"
        try:
            data = json.loads(raw.decode("utf-8"))
        except ValueError:
            raise UserError("בקשה לא תקינה")
        if not isinstance(data, dict):
            raise UserError("בקשה לא תקינה")
        return data

    # ---- GET
    def do_GET(self):
        self.app.last_seen = time.time()
        if not self._host_ok():
            return self._deny(403, "bad host")
        url = urllib.parse.urlsplit(self.path)
        if url.path == "/":
            query = urllib.parse.parse_qs(url.query)
            if not self._token_ok((query.get("t") or [None])[0]):
                return self._deny(403, "missing or wrong token")
            return self._page()
        if not url.path.startswith("/api/"):
            return self._deny(404, "not found")
        if not self._token_ok(self.headers.get("X-Token")) or not self._origin_ok():
            return self._deny(403, "forbidden")
        if url.path == "/api/ping":
            return self._json({"ok": True})
        if url.path == "/api/env":
            return self._json({
                "missing": core.missing_modules(need_excel=True), "install": core.install_command(),
                "default_result": str(default_result_path()), "overwrite": bool(load_settings().get("overwrite", False)),
                "current_day_value": load_settings().get("current_day_value", 0),
                "day_value_limits": self.app.cfg.get("day_value_limits"),
                "version": VERSION,
                "files": [{"id": u["id"], "name": u["name"], "size": u["size"], "kind": u["kind"]} for u in self.app.uploads]})
        if url.path == "/api/download":
            return self._download()
        return self._deny(404, "not found")

    def _page(self):
        nonce = secrets.token_urlsafe(16)
        html = (HERE / "web_ui.html").read_text(encoding="utf-8").replace("__NONCE__", nonce).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html)))
        self.send_header("Content-Security-Policy",
                         "default-src 'none'; script-src 'nonce-%s'; style-src 'nonce-%s'; connect-src 'self'; "
                         "img-src 'self' data:; base-uri 'none'; form-action 'none'; frame-ancestors 'none'" % (nonce, nonce))
        self.send_header("X-Frame-Options", "DENY")
        self._common_headers()
        self.end_headers()
        self.wfile.write(html)

    def _download(self):
        path = self.app.last_result
        if not path or not Path(path).is_file():
            return self._deny(404, "no result yet")
        data = Path(path).read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", 'attachment; filename="%s"' % re.sub(r'[^A-Za-z0-9._-]', "_", Path(path).name))
        self._common_headers()
        self.end_headers()
        self.wfile.write(data)

    # ---- POST
    def do_POST(self):
        self.app.last_seen = time.time()
        if not self._host_ok():
            self._drain_body(DENIED_BODY_LIMIT)          # answer cleanly instead of resetting the connection
            return self._deny(403, "bad host")
        url = urllib.parse.urlsplit(self.path)
        if not url.path.startswith("/api/") or not self._token_ok(self.headers.get("X-Token")) or not self._origin_ok():
            self._drain_body(DENIED_BODY_LIMIT)
            return self._deny(403, "forbidden")
        try:
            if url.path == "/api/upload":
                return self._upload(url)
            data = self._body_json()
            if url.path == "/api/run":
                with self.app.lock:
                    return self._json(self.app.run(data))
            if url.path == "/api/result-info":
                try:
                    p = self.app.resolve_result(data.get("result"))
                except UserError as e:
                    return self._json({"error": str(e)})
                info = {"path": str(p), "exists": p.is_file()}
                if info["exists"]:
                    problems, more = core.check_result_compatibility(p, self.app.cfg)
                    old_layout = core.old_layout_of(p, self.app.cfg) if problems else None
                    info.update(compatible=not problems or bool(old_layout), problems=[] if old_layout else problems,
                                more=0 if old_layout else more, old_layout=bool(old_layout))
                has_backup, when = core.backup_info(p)
                info.update(has_backup=has_backup, backup_time=when.strftime("%d/%m/%Y %H:%M") if when else None)
                return self._json(info)
            if url.path == "/api/pick-result":
                try:
                    start = self.app.resolve_result(data.get("current"))
                except UserError:
                    start = default_result_path()
                chosen = run_save_dialog(start)
                if not chosen:
                    return self._json({"cancelled": True})
                try:
                    p = self.app.resolve_result(chosen)
                except UserError as e:
                    return self._json({"error": str(e)})
                save_settings(result=str(p))
                return self._json({"path": str(p)})
            if url.path == "/api/restore-backup":
                p = self.app.resolve_result(data.get("result"))
                try:
                    restored = core.restore_backup(p)
                except PermissionError:
                    raise UserError("אי אפשר לשחזר: הקובץ פתוח באקסל. סגור אותו ונסה שוב.")
                return self._json({"restored": restored})
            if url.path == "/api/remove":
                self.app.remove_upload(int(data.get("id", -1)))
                return self._json({"ok": True})
            if url.path == "/api/clear":
                kind = data.get("kind")
                if kind is not None and kind not in UPLOAD_KINDS:
                    raise UserError("סוג לא מוכר")
                self.app.clear_uploads(kind)
                return self._json({"ok": True})
            if url.path == "/api/open-folder":
                path = self.app.last_result
                if path and hasattr(os, "startfile") and Path(path).parent.is_dir():
                    os.startfile(str(Path(path).parent))        # noqa: S606 - opens the user's own results folder
                return self._json({"ok": True})
            if url.path == "/api/shutdown":
                self._json({"ok": True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
                return None
            return self._deny(404, "not found")
        except NeedsConfirmation as e:
            return self._json({"error": "הקובץ שנבחר לא תואם לתבנית המובנית.", "incompatible": True,
                               "problems": e.problems}, 409)
        except UserError as e:
            return self._json({"error": str(e)}, 400)
        except (ValueError, TypeError):
            return self._json({"error": "בקשה לא תקינה"}, 400)
        except Exception:
            return self._json({"error": "אירעה שגיאה בלתי צפויה. נסה שוב."}, 500)

    def _drain_body(self, limit=MAX_FILE_BYTES):
        """Read and discard the request body, so a rejected request still gets a clean answer in the browser
        (closing with unread data makes Windows reset the connection before the answer arrives)."""
        try:
            remaining = min(int(self.headers.get("Content-Length") or 0), limit)
        except ValueError:
            return
        while remaining > 0:
            chunk = self.rfile.read(min(1 << 20, remaining))
            if not chunk:
                break
            remaining -= len(chunk)

    def _upload(self, url):
        query = urllib.parse.parse_qs(url.query)
        name = safe_display_name((query.get("name") or [""])[0])
        kind = (query.get("kind") or ["slips"])[0]
        try:
            if kind not in UPLOAD_KINDS:
                raise UserError("סוג קובץ לא מוכר")
            if not name.lower().endswith(".pdf"):
                raise UserError("רק קובצי PDF נתמכים: %s" % name)
            try:
                length = int(self.headers.get("Content-Length") or -1)
            except ValueError:
                length = -1
            if length <= 0:
                raise UserError("הקובץ ריק או חסר גודל")
            if length > MAX_FILE_BYTES:
                raise UserError("הקובץ גדול מדי (מעל 200MB): %s" % name)
            if len(self.app.uploads) >= MAX_FILES or self.app.total_bytes() + length > MAX_TOTAL_BYTES:
                raise UserError("הגעת למגבלת הקבצים לסשן. נקה את הרשימה והמשך.")
        except UserError:
            self._drain_body()
            raise
        up = self.app.new_upload(name, kind)
        remaining, first = length, True
        with open(up["path"], "wb") as f:
            while remaining > 0:
                chunk = self.rfile.read(min(1 << 20, remaining))
                if not chunk:
                    break
                if first:
                    first = False
                    if not chunk.startswith(b"%PDF-"):
                        f.close()
                        up["path"].unlink(missing_ok=True)
                        raise UserError("הקובץ אינו PDF תקין: %s" % name)
                f.write(chunk)
                remaining -= len(chunk)
                up["size"] += len(chunk)
        if remaining:
            up["path"].unlink(missing_ok=True)
            raise UserError("ההעלאה נקטעה: %s" % name)
        self.app.add_upload(up)
        return self._json({"id": up["id"], "name": up["name"], "size": up["size"], "kind": up["kind"]})


class Server(http.server.ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, app, port=0):
        super().__init__(("127.0.0.1", port), Handler)           # loopback ONLY
        self.app = app
        app.httpd = self
        app.port = self.server_address[1]
        self.allowed_hosts = {"127.0.0.1:%d" % app.port, "localhost:%d" % app.port}
        self.allowed_origins = {"http://127.0.0.1:%d" % app.port, "http://localhost:%d" % app.port}


def cleanup_old_temp_dirs(max_age_hours=12):
    base = Path(tempfile.gettempdir())
    cutoff = time.time() - max_age_hours * 3600
    for d in base.glob(TEMP_PREFIX + "*"):
        try:
            if d.is_dir() and d.stat().st_mtime < cutoff:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


def idle_watchdog(app):
    while True:
        time.sleep(5)
        if time.time() - app.last_seen > IDLE_SECONDS:
            app.httpd.shutdown()
            return


def start(port=0, session_dir=None):
    app = App(session_dir)
    server = Server(app, port)
    return app, server


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--port", type=int, default=0)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--print-url", action="store_true")
    ap.add_argument("--settings-file", help="remember choices in this file instead of ~/.reserve_duty_web.json (for testing)")
    args = ap.parse_args(argv)
    if args.settings_file:
        globals()["SETTINGS_FILE"] = Path(args.settings_file)
    cleanup_old_temp_dirs()
    app, server = start(args.port)
    url = "http://127.0.0.1:%d/?t=%s" % (app.port, app.token)
    if args.print_url and sys.stdout:
        print(url, flush=True)
    threading.Thread(target=idle_watchdog, args=(app,), daemon=True).start()
    if not args.no_browser and not webbrowser.open(url):
        try:
            import tkinter
            from tkinter import messagebox
            root = tkinter.Tk()
            root.withdraw()
            messagebox.showinfo("מילוי ערך ימ״מ", "פתח את הכתובת הבאה בדפדפן:\n" + url)
        except Exception:
            pass
    try:
        server.serve_forever()
    finally:
        server.server_close()
        shutil.rmtree(app.session_dir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    # opened with a Python that lacks the libraries (e.g. by double-clicking this file): use the private environment
    if core.relaunch_with_local_env(Path(__file__).resolve(), sys.argv[1:], windowed=True):
        sys.exit(0)
    sys.exit(main())
