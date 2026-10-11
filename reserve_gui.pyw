"""חלון להפעלת הסקריפט: בחירת תלושים, בדיקה, ומילוי קובץ האקסל."""
import contextlib
import json
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import fill_reserve_duty as core  # noqa: E402

SETTINGS = Path.home() / ".reserve_duty_gui_settings.json"      # in the user's profile, never next to the code
FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
TITLE = "מילוי ערך ימ״מ מתלושי שכר"


def R(s):
    """Force right-to-left direction (Tk defaults to left-to-right, which scrambles mixed Hebrew/English)."""
    return chr(0x202B) + s + chr(0x202C)     # RLE ... PDF (right-to-left embedding)


def default_output_path():
    """Results contain real pay figures, so default to the user's Documents folder, never the code folder."""
    docs = Path.home() / "Documents"
    return str((docs if docs.is_dir() else Path.home()) / "filled_reserve_duty.xlsx")


class QueueWriter:
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("text", s))
        return len(s)

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.running = False
        self.settings = self._load_settings()
        self.items = []

        root.title(TITLE)
        root.geometry("860x700")
        root.minsize(720, 560)
        style = ttk.Style()
        style.configure("TLabel", font=FONT)
        style.configure("TButton", font=FONT)
        style.configure("TCheckbutton", font=FONT)
        style.configure("Head.TLabel", font=FONT_BOLD)

        outer = ttk.Frame(root, padding=12)
        outer.pack(fill="both", expand=True)

        # 1. slips
        ttk.Label(outer, text=R("1. בחר תלושים (קובצי PDF או תיקייה)"), style="Head.TLabel", anchor="e", justify="right").pack(fill="x")
        row = ttk.Frame(outer)
        row.pack(fill="x", pady=(4, 4))
        ttk.Button(row, text="הוסף קובצי PDF…", command=self.add_files).pack(side="right", padx=(6, 0))
        ttk.Button(row, text="הוסף תיקייה…", command=self.add_folder).pack(side="right", padx=(6, 0))
        ttk.Button(row, text="הסר מסומן", command=self.remove_selected).pack(side="right", padx=(6, 0))
        ttk.Button(row, text="נקה רשימה", command=self.clear_items).pack(side="right")
        box = ttk.Frame(outer)
        box.pack(fill="x")
        self.listbox = tk.Listbox(box, height=5, font=FONT, selectmode="extended", activestyle="none")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.listbox.pack(side="left", fill="x", expand=True)

        # 2. where to save (the template is always the built-in one)
        self.output = tk.StringVar(value=self.settings.get("output") or default_output_path())
        self._path_row(outer, "2. איפה לשמור את קובץ האקסל (אפשר לבחור קובץ קיים כדי לעדכן אותו)", self.output, self.pick_output)

        self.overwrite = tk.BooleanVar(value=bool(self.settings.get("overwrite", False)))
        ttk.Checkbutton(outer, text=R("לדרוס ערכים קיימים שונים בתאים (בלי זה, תא עם ערך אחר מדולג)"), variable=self.overwrite).pack(anchor="e", pady=(8, 0))

        # actions
        act = ttk.Frame(outer)
        act.pack(fill="x", pady=(10, 6))
        self.btn_fill = ttk.Button(act, text="מילוי האקסל", command=lambda: self.start("fill"))
        self.btn_check = ttk.Button(act, text="בדיקה (בלי לכתוב)", command=lambda: self.start("check"))
        self.btn_open = ttk.Button(act, text="פתח את קובץ התוצאה", command=self.open_output, state="disabled")
        self.btn_restore = ttk.Button(act, text="שחזר גיבוי", command=self.restore)
        self.btn_fill.pack(side="right", padx=(8, 0))
        self.btn_check.pack(side="right", padx=(8, 0))
        self.btn_open.pack(side="right")
        self.btn_restore.pack(side="right", padx=(0, 8))
        self.status = ttk.Label(act, text="", anchor="w")
        self.status.pack(side="left")

        ttk.Label(outer, text=R("תוצאות"), style="Head.TLabel", anchor="e", justify="right").pack(fill="x")
        txt = ttk.Frame(outer)
        txt.pack(fill="both", expand=True, pady=(4, 0))
        self.text = tk.Text(txt, wrap="word", font=("Segoe UI", 10), height=12, state="disabled")
        tsb = ttk.Scrollbar(txt, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=tsb.set)
        tsb.pack(side="left", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        self.text.tag_configure("bad", foreground="#B00020")
        self.text.tag_configure("good", foreground="#1B6E2E")
        self.text.tag_configure("head", font=("Segoe UI", 10, "bold"))

        ttk.Label(outer, text=R("ימי מילואים (שורה 27) וערך ימ״מ נוכחי (F36) או תשלומי ביטוח לאומי (שורות 22–23) ממלאים ידנית באקסל, או בממשק הדפדפן."), anchor="e", justify="right").pack(fill="x", pady=(6, 0))
        ttk.Label(outer, text=R("הכלי מבוסס על פרשנות אישית לחוק ואינו ייעוץ משפטי, חשבונאי או מקצועי. אין לי קשר לביטוח לאומי. לפני שפועלים לפי התוצאות – יש להתייעץ עם איש מקצוע מוסמך."), anchor="e", justify="right", wraplength=640, foreground="#9C5700").pack(fill="x", pady=(6, 0))

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.poll)
        self.root.after(400, self.warn_missing_libraries)

    # ---------- layout helpers
    def _path_row(self, parent, label, var, command, clear_label=None):
        ttk.Label(parent, text=R(label), style="Head.TLabel", anchor="e", justify="right").pack(fill="x", pady=(10, 0))
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(4, 0))
        ttk.Button(row, text="עיון…", command=command).pack(side="right", padx=(6, 0))
        if clear_label:
            ttk.Button(row, text=clear_label, command=lambda: var.set("")).pack(side="right", padx=(6, 0))
        ttk.Entry(row, textvariable=var, font=FONT).pack(side="left", fill="x", expand=True)

    def warn_missing_libraries(self):
        """Tell the user right away if this Python lacks a needed library, and how to install it."""
        missing = core.missing_modules()
        if not missing:
            return
        text = ("חסרות ספריות בפייתון הזה: %s\nהפייתון שרץ עכשיו: %s\n\n"
                "להתקנה פתח טרמינל (PowerShell) והרץ:\n%s\n\nאחר כך הפעל את התוכנה מחדש."
                % (", ".join(missing), sys.executable, core.install_command()))
        self.log(text + "\n", "bad")
        messagebox.showwarning(TITLE, text)

    # ---------- settings
    def _load_settings(self):
        try:
            return json.loads(SETTINGS.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_settings(self):
        self.settings.update(output=self.output.get(), overwrite=self.overwrite.get())
        try:
            SETTINGS.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass        # settings are a convenience; a read-only folder must not break the app

    def on_close(self):
        self._save_settings()
        self.root.destroy()

    # ---------- list handling
    def _refresh_list(self):
        self.listbox.delete(0, "end")
        for p in self.items:
            self.listbox.insert("end", p)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="בחר קובצי תלוש", initialdir=self.settings.get("pdf_dir") or str(Path.home()),
            filetypes=[("PDF", "*.pdf"), ("כל הקבצים", "*.*")])
        if paths:
            self.settings["pdf_dir"] = str(Path(paths[0]).parent)
            for p in paths:
                p = os.path.normpath(p)
                if p not in self.items:
                    self.items.append(p)
            self._refresh_list()

    def add_folder(self):
        d = filedialog.askdirectory(title="בחר תיקייה עם תלושים", initialdir=self.settings.get("pdf_dir") or str(Path.home()))
        if d:
            d = os.path.normpath(d)
            self.settings["pdf_dir"] = d
            if d not in self.items:
                self.items.append(d)
            self._refresh_list()

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            del self.items[i]
        self._refresh_list()

    def clear_items(self):
        self.items = []
        self._refresh_list()

    def pick_output(self):
        p = filedialog.asksaveasfilename(title="שמור קובץ תוצאה", defaultextension=".xlsx",
                                         initialdir=str(Path(self.output.get()).parent),
                                         initialfile=Path(self.output.get()).name,
                                         filetypes=[("Excel", "*.xlsx")], confirmoverwrite=False)
        if p:
            self.output.set(os.path.normpath(p))

    # ---------- running
    def log(self, s, tag=None):
        self.text.configure(state="normal")
        for line in s.splitlines(True):
            t = tag
            up = line.upper()
            if t is None:
                if "ERROR" in up or "SKIPPED" in up or "WARNING" in up or "NOT FOUND" in up:
                    t = "bad"
                elif line.lstrip().startswith("wrote") or line.startswith("saved"):
                    t = "good"
                elif line.startswith("=="):
                    t = "head"
            self.text.insert("end", line, t)
        self.text.see("end")
        self.text.configure(state="disabled")

    def start(self, mode):
        if self.running:
            return
        if not self.items:
            messagebox.showwarning(TITLE, "לא נבחרו תלושים. הוסף קובצי PDF או תיקייה.")
            return
        argv = list(self.items) + ["--ocr"]               # image-only certificates are read by OCR if the add-on is installed
        if mode == "check":
            argv.append("--check")
        else:
            out = self.output.get().strip()
            if not out:
                messagebox.showerror(TITLE, "לא נבחר קובץ תוצאה.")
                return
            if core.inside_program_folder(out):
                messagebox.showerror(TITLE, "קובץ התוצאה מכיל נתוני שכר אמיתיים ולכן אסור לשמור אותו בתיקיית התוכנה "
                                            "(היא עלולה להיארז או להתפרסם).\nבחר תיקייה אחרת, למשל מסמכים.")
                return
            argv += ["--out", out]
            if Path(out).is_file():                    # an existing file must match the built-in template
                problems, _more = core.check_result_compatibility(out, json.loads((HERE / "config.json").read_text(encoding="utf-8")))
                if problems:
                    shown = "\n".join("• " + x for x in problems[:6])
                    if not messagebox.askyesno(TITLE, "הקובץ שנבחר לא תואם לתבנית המובנית:\n\n" + shown +
                                               "\n\nלהחליף אותו בקובץ חדש? הקובץ הישן יישמר בגיבוי נסתר ואפשר לשחזר אותו.", icon="warning"):
                        return
                    argv.append("--replace-incompatible")
            if self.overwrite.get():
                argv.append("--overwrite")
        self._save_settings()
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
        self.running = True
        for b in (self.btn_fill, self.btn_check):
            b.configure(state="disabled")
        self.status.configure(text="מעבד…")
        threading.Thread(target=self._worker, args=(argv, mode), daemon=True).start()

    def _worker(self, argv, mode):
        w = QueueWriter(self.q)
        rc = 0
        try:
            with contextlib.redirect_stdout(w), contextlib.redirect_stderr(w):
                try:
                    rc = core.main(argv)
                except SystemExit as e:
                    rc = e.code if isinstance(e.code, int) else 1
        except Exception:
            self.q.put(("text", "ERROR:\n" + traceback.format_exc()))
            rc = 3
        self.q.put(("done", (rc, mode)))

    def poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "text":
                    self.log(payload)
                else:
                    rc, mode = payload
                    self._finish(rc, mode)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def _finish(self, rc, mode):
        self.running = False
        for b in (self.btn_fill, self.btn_check):
            b.configure(state="normal")
        if mode == "fill" and Path(self.output.get()).is_file():
            self.btn_open.configure(state="normal")
        if rc == 0:
            self.status.configure(text="הסתיים")
            self.log("\n— הסתיים —\n", "good")
        elif rc == 1:
            self.status.configure(text="הסתיים עם דילוגים")
            self.log("\n— הסתיים, אבל חלק מהתלושים או התאים דולגו (ראה שורות באדום) —\n", "bad")
        else:
            self.status.configure(text="שגיאה")
            self.log("\n— שגיאה, ראה פירוט למעלה —\n", "bad")

    def restore(self):
        out = self.output.get().strip()
        exists, when = core.backup_info(out) if out else (False, None)
        if not exists:
            messagebox.showinfo(TITLE, "לא נמצא גיבוי לקובץ הזה.")
            return
        if not messagebox.askyesno(TITLE, "לשחזר את הגרסה הקודמת של הקובץ (גיבוי מ-%s)?\nהגרסה הנוכחית תישמר כגיבוי, ולחיצה נוספת תחזיר אותה." % when.strftime("%d/%m/%Y %H:%M")):
            return
        try:
            core.restore_backup(out)
        except PermissionError:
            messagebox.showerror(TITLE, "אי אפשר לשחזר: הקובץ פתוח באקסל. סגור אותו ונסה שוב.")
            return
        self.log("\nשוחזרה הגרסה הקודמת של הקובץ.\n", "good")

    def open_output(self):
        p = Path(self.output.get())
        if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm"):
            os.startfile(str(p))  # noqa: S606 - opens the user's own Excel result file


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    import fill_reserve_duty as _core

    if not _core.relaunch_with_local_env(Path(__file__).resolve(), sys.argv[1:], windowed=True):
        main()
