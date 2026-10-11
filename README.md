Reserve-duty (ימ"מ) pay-slip filler
===================================

> **Disclaimer.** I am not an accountant, a lawyer or a tax advisor, and I have no connection to the National Insurance
> Institute (Bituach Leumi), the IDF or any other official body. This tool reflects my own personal interpretation of the
> law; it may contain errors and is NOT legal, accounting or professional advice. Use it at your own risk and consult a
> qualified professional before acting on its results. Provided "AS IS", without warranty. Full text: DISCLAIMER.md.
>
> **הבהרה.** אינני רואה חשבון, עורך דין או יועץ מס, ואין לי כל קשר לביטוח לאומי, לצה"ל או לגוף רשמי אחר. הכלי מבוסס על
> הפרשנות האישית שלי לחוק, עלול לטעות, ואינו ייעוץ משפטי, חשבונאי או מקצועי. השימוש על אחריותך בלבד; לפני שפועלים לפי
> התוצאות יש להתייעץ עם איש מקצוע מוסמך. מסופק "כמות שהוא", ללא אחריות. הנוסח המלא: DISCLAIMER.md.

Quick start (Windows)
---------------------
1. Install Python 3.11 or newer from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
2. Download this repository (green "Code" button -> "Download ZIP") and extract it.
3. Double-click install.bat once. It creates a private environment (.venv) in this folder and downloads the libraries
   from requirements.txt - nothing is installed system-wide. It asks whether to add the optional OCR add-on.
4. Double-click start.bat. The program opens in your browser; it runs only on your computer.

What it does
------------
Reads Hilan pay-slip PDFs and fills the reserve-duty value workbook (the Excel template is built in).
For each slip it writes, into the sheet of the slip's year and the column of its month:
  - gross for National Insurance      ("ברוטו לב.לאומי")
  - variable gross                    ("ברוטו לב.ל ב.ק.", 0 if absent)
  - reserve payments: codes 198 + 197 (summed)
  - top-ups: code 193                 (summed)
What National Insurance actually paid comes from its annual letter ("אישור שנתי בגין תשלומים שהועברו למעסיק"; on the
NI website btl.gov.il: מילואים -> אישורים -> "אישור שנתי בגין תשלומים שהועברו למעסיקך", one letter per year): for each month the value paid per day (total paid / days, row 22) and the days it paid for
(row 23). Instead, a single current reserve-day value ("ערך ימ"מ נוכחי") can be typed into the web interface (or
--current-day-value); it is written to F36 of every year sheet and, as in the template's formula, it OVERRIDES the
letter's monthly values for the whole year (you are warned when both are given). A value outside 328.76 .. 1730.33
(day_value_limits in config.json) gets a warning to check it is the value of ONE day; it is still written.
Reserve days (row 27, "ע"פ 3010") are filled from the reserve-service certificate (form 3010) if you upload it; otherwise
enter them by hand - without them the difference is 0. Row 28 shows the days NI did not pay for (row 27 - row 23).
The helper sheet's block "סה"כ ביטוח לאומי חייב לי" adds up the annual difference (F38) of every year sheet.
A month with reserve days but no value paid (row 22 empty and no F36) is counted as if nothing was paid; you are warned.
It works for years 2025-2027 (the template's sheets). Everything runs locally; no data leaves your computer.

The fill button works with whatever you have: values missing from a slip are left empty for you to type in (and listed),
and with no files at all it creates a ready-to-use workbook. In every workbook the program writes, the formulas are
hidden and locked (sheet protection, not encryption); only the cells you type into are editable. To be able to unprotect
your own files, set the environment variable RESERVE_DUTY_SHEET_PASSWORD before running; otherwise a random password is
used and not kept.

Credit and license
------------------
Made by Shai Shallev. Free to use, share and change, as long as the credit stays:
  - program code: MIT License (LICENSE) - keep the copyright notice
  - Excel template and the workbooks it produces: CC BY 4.0 (LICENSE-TEMPLATE.md) - credit "Shai Shallev"
Every workbook carries the author in its properties and "© Shai Shallev · CC BY 4.0" in the printed footer.
This is a personal tool, not affiliated with National Insurance or the IDF. Check the results against your documents;
it comes with no warranty.

One-time setup (Windows)
------------------------
The easy way is install.bat (see Quick start). By hand, the same thing is:
       python -m venv .venv
       .venv\Scripts\python -m pip install -r requirements.txt
       .venv\Scripts\python -m pip install -r requirements-ocr.txt      (optional OCR add-on)

Three ways to use it
--------------------
1. Web interface (recommended)  -  double-click  start.bat  (or reserve_web.pyw)
   Opens a page in your browser (Hebrew, drag and drop, results table, download button).
   It runs only on your computer: the server listens on 127.0.0.1, every request needs a secret token,
   uploaded PDFs are copied to a private temporary folder that is deleted when you press "סגור את התוכנה"
   (or close the tab). Nothing is loaded from the internet.
2. Desktop window  -  double-click  reserve_gui.pyw
   A simple Hebrew window with the same functions.
3. Command line
       python fill_reserve_duty.py slips_folder --check
       python fill_reserve_duty.py slip1.pdf slip2.pdf --out "%USERPROFILE%\Documents\result.xlsx" [--overwrite]

How to use the windows
----------------------
  1. Add the files. The web interface has two boxes side by side: pay-slips (PDFs or a whole folder; a single PDF may
     contain many slips) and the reserve-service certificate (form 3010). A file in the wrong box is reported, not guessed.
  2. Choose where to save the Excel file: type a path, or press the "choose location" button to open the Windows file
     dialog. The built-in template is always used. To update a file you made earlier, pick that file.
  3. Press the check button first to preview the extracted values, then the fill button to write them.
     The check also shows the bottom line "total National Insurance owes me", worked out in advance with the same
     formulas as the Excel file (from all checked data, plus the chosen result file if it already exists). It is
     final only after the current day value is entered. Nothing is written by a check.
Close the result Excel file before filling. Always compare the preview with a slip by eye.

Reserve-service certificate (form 3010)
----------------------------------------
Upload the certificate PDF together with (or without) the pay-slips. Every period in its table is split by calendar month
(a period such as 18/01 - 05/04 gives days to January, February, March and April; days of years without a sheet are ignored),
each calendar day is counted once (identical or overlapping rows are merged), and the monthly totals are written to row 27
of each year sheet, REPLACING the value that is there (never added to it). Months without service are left untouched.
Upload ALL pages of the certificate in one run; missing pages are reported. A row whose stated days differ from its
calendar days is left out and reported.

Certificates that are only a picture (OCR)
------------------------------------------
Some certificate PDFs contain no text, only a picture of each page. The certificate box (and the command line with --ocr)
reads such a file by image recognition, if the optional add-on is installed:
    pip install -r requirements-ocr.txt          (rapidocr + onnxruntime, about 45 MB, works offline afterwards)
It takes about 8 seconds per page, and is used ONLY for the certificate, never for pay-slips (Hebrew pay-slip scans are
not supported). Every row is checked against the calendar (stated days must equal the days between the dates), and the
result is marked "read by OCR" so you compare the totals with the certificate. Page numbers cannot be read from a picture,
so make sure the whole certificate is in the file.

Where results go
----------------
Result workbooks contain real pay figures. The programs refuse to write them inside the program folder, so they cannot
end up in a zip or a repository by accident. The location you choose is remembered (default: your Documents folder).

Service-period bonus (separate sheet)
-------------------------------------
From the service days (the 3010 certificate, or the NI letter if there is no certificate) the program builds the
continuous service periods: periods that follow each other without a gap are one period, also across months and years
(20-21.09 + 22-24.09 = 20-24.09). For each period, the remainder of its length divided by 7 gives a bonus:
remainder 1-5 -> 0.4 x remainder day values (12 days -> 2 day values), remainder 6 -> one day value, 0 -> nothing.
The whole period is valued by the correct daily value (row 18) of the month it STARTS in, and its bonus is counted in
that month. The sheet "תוספת לתקופות שירות" lists every period, then a summary per month and per year and the grand
total. It is kept separate: it is NOT part of "סה"כ ביטוח לאומי חייב לי". The start/end dates in it can be edited.

Existing result files, backup and restore
-----------------------------------------
- Picking an existing result file updates it: slips are added and cells that already hold a different value are
  skipped (unless you tick "overwrite different values"; rows 22, 23 and 27 are always replaced).
- A result made by an earlier version (days in row 23, day value in F31) is converted to the current layout automatically:
  every input is copied over, and the old file is kept as the hidden backup.
- The file must match the built-in template (same sheets and labels). If it does not, the program tells you why and
  asks whether to REPLACE it with a fresh file. Nothing is changed unless you confirm.
- Before every update, the file as it was is copied to a hidden backup next to it: .<name>.bkup (it has the Windows
  "hidden" attribute, so Explorer does not show it by default). The "restore backup" button puts it back; the version
  being replaced becomes the new backup, so pressing the button twice undoes the restore.
- Command line: --out FILE, --replace-incompatible, --restore-backup FILE.

Checking a result file
----------------------
To see which cells of a result workbook are filled (without showing any figure, so the output is safe to share):
    python fill_reserve_duty.py --inspect result.xlsx

If your slips look different
----------------------------
The program finds values by meaning, not by a fixed layout: the slip title, the codes (198, 197, 193) in a payments
table whose amount column is found by its header ("סמל"/"קוד" ... "סכום"), and the labels of the National Insurance
gross. A deductions table ("ניכויים") is never counted. All wordings are lists in config.json, so a new spelling
is one more entry, not a code change:
    title_keys              slip title wordings ("תלוש שכר לחודש", "תלוש משכורת", ...); the month may be a name or 03/2026
    labels                  wordings of the gross and the unfixed gross ("ברוטו לב.לאומי", "ברוטו לב\"ל", ...)
    label_not_followed_by   words that make a match something else ("ברוטו לב\"ל ב.ק." is not the gross; "מצטבר" is year-to-date)
    codes                   the payment codes that are summed
    parsing.table_header    header and table-title words
Anything uncertain is flagged, never guessed: a missing title, two tables with the same code, a row that mentions
reserve duty (מילואים) under a code that is not configured, two reading methods that disagree.
Not supported: pay-slips that are only a picture (the image recognition cannot read Hebrew), and password-protected
PDFs (save a copy without the password first).
If a PDF is not read properly, run (both show the structure only, no figures):
    python fill_reserve_duty.py file.pdf --diagnose
    python fill_reserve_duty.py file.pdf --debug-text

Development
-----------
    python -m unittest discover -s tests -v
embed_template.py rebuilds template_data.py from an Excel template (use --share to strip origin data).

---------------------------------------------------------------------------------------------

מילוי ערך ימ"מ מתלושי שכר
=========================
הכלי קורא תלושי שכר של חילן (PDF) ומשלים את קובץ האקסל לחישוב ערך יום מילואים (התבנית מובנית בתוכנה).
התלוש נקרא מקומית במחשב בלבד. מה שביטוח לאומי שילם בפועל נקרא מ"אישור שנתי בגין תשלומים שהועברו למעסיק" (מאתר ביטוח לאומי, לכל שנה):
(מפיקים באתר הביטוח הלאומי: מילואים ← אישורים ← "אישור שנתי בגין תשלומים שהועברו למעסיקך", אישור נפרד לכל שנה.)
לכל חודש נכתבים הערך ששולם ליום (סכום ÷ ימים, שורה 22) ומספר הימים ששולמו (שורה 23). במקומו אפשר להקליד ערך ימ"מ נוכחי אחד,
שנכתב לתא F36 בכל גיליון שנה וגובר על ערכי האישור לכל השנה (תתקבל אזהרה אם הוזנו שניהם).
ערך מחוץ לטווח 328.76 עד 1,730.33 (day_value_limits בקובץ config.json) מקבל אזהרה לבדוק שזה הסכום של יום אחד, והוא נכתב בכל זאת.
ימי מילואים (שורה 27, ע"פ 3010) מתמלאים מאישור שירות המילואים (טופס 3010), ובלעדיו ממלאים אותם ידנית – בלעדיהם ההפרש הוא 0.
הטבלה "סה"כ ביטוח לאומי חייב לי" בגיליון העזר מסכמת את ההפרש השנתי (F38) של כל השנים.
תוספת לתקופות שירות (גיליון נפרד): תקופות רצופות מאוחדות לתקופה אחת. שארית 1–5 של מספר הימים חלקי 7 מזכה ב-0.4 × השארית
ערכי יום, שארית 6 בערך יום אחד. ערך היום הוא הערך היומי המחושב (שורה 18) של חודש תחילת התקופה. יש סיכום לפי חודש, לפי שנה
וסה"כ, בנפרד מ"סה"כ ביטוח לאומי חייב לי".
קובץ תוצאה שנוצר בגרסה קודמת (ימים בשורה 23, ערך בתא F31) מומר אוטומטית למבנה החדש, כל הנתונים מועתקים, והקובץ הישן נשמר בגיבוי.

התקנה חד-פעמית: התקן Python מ-python.org (סמן Add to PATH), הורד את המאגר (כפתור Code -> Download ZIP) וחלץ אותו,
ואז לחיצה כפולה על install.bat – הוא יוצר סביבה פרטית בתיקייה (.venv) ומוריד את הספריות; שום דבר לא מותקן במחשב עצמו.

אם אישור המילואים הוא קובץ PDF שהוא רק תמונה (בלי טקסט), אפשר להתקין תוסף זיהוי תמונה אופציונלי: pip install -r requirements-ocr.txt
(לוקח כ-8 שניות לעמוד, משמש רק לאישור המילואים, והתוצאה מסומנת "נקרא בזיהוי תמונה" כדי שתשווה את הסכומים לאישור).

הפעלה (מומלץ): לחיצה כפולה על start.bat. נפתח דף בדפדפן: גוררים קובצי PDF או תיקייה, לוחצים "בדיקה בלבד"
לבדיקת הערכים, ואז "מילוי האקסל". הכל רץ במחשב שלך בלבד, והקבצים הזמניים נמחקים כשסוגרים את התוכנה.
בבדיקה מוצג גם הסכום "סה״כ ביטוח לאומי חייב לי", מחושב מראש באותן נוסחאות של האקסל (מכל הנתונים שנבדקו, ומקובץ התוצאה אם הוא כבר קיים). הסכום סופי רק אחרי שהוזן ערך ימ״מ נוכחי.
חלופה: חלון רגיל בלחיצה כפולה על reserve_gui.pyw.

"מילוי האקסל" עובד גם עם מידע חלקי: ערך שלא נמצא בתלוש נשאר ריק להשלמה ידנית (ומופיע ברשימה), ובלי קבצים בכלל נוצר
קובץ מוכן לשימוש. בכל קובץ שהתוכנה כותבת הנוסחאות מוסתרות ונעולות (הגנת גיליון, לא הצפנה), ורק תאי ההזנה פתוחים לעריכה.

קרדיט ורישיון: נוצר על ידי Shai Shallev. חופשי לשימוש, שיתוף ושינוי, בתנאי שהקרדיט נשמר:
הקוד ברישיון MIT, ותבנית האקסל והקבצים שהיא יוצרת ברישיון CC BY 4.0. כלי אישי, לא קשור לביטוח לאומי או לצה"ל, ובלי אחריות.

תלושים של מערכות שכר אחרות: הכלי מחפש את הערכים לפי המשמעות ולא לפי מיקום קבוע: כותרת התלוש, הקודים (198, 197, 193)
בטבלת התשלומים (עמודת הסכום מזוהה לפי הכותרת "סכום"), והשורה "ברוטו לביטוח לאומי" על כל הקיצורים שלה. טבלת ניכויים לא נספרת.
כל הניסוחים הם רשימות בקובץ config.json, וניסוח חדש הוא עוד שורה ברשימה. כל דבר לא ודאי מסומן "לבדיקה" ולא מנוחש,
למשל שורה שכתוב בה "מילואים" עם קוד שלא מוגדר. לא נתמכים: תלוש שהוא רק תמונה, ו-PDF מוגן בסיסמה (שמור קודם עותק בלי סיסמה).

את מקום השמירה בוחרים בלחצן "בחר מיקום" (נפתח חלון בחירת קובץ רגיל), והבחירה נזכרת בפעם הבאה. אפשר לבחור קובץ קיים כדי לעדכן אותו.
הכלי משתמש תמיד בתבנית המובנית. קובץ קיים שלא תואם לתבנית לא ישתנה בלי אישור שלך, ולפני כל עדכון נשמר גיבוי נסתר
(.שם-הקובץ.bkup) שאפשר לשחזר בלחצן "שחזר גיבוי".
קובץ התוצאה מכיל נתוני שכר אמיתיים ולכן אי אפשר לשמור אותו בתיקיית התוכנה.
יש לסגור את קובץ התוצאה באקסל לפני המילוי, ולהשוות תמיד את הערכים המוצגים לתלוש.
