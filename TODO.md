# TODO

Open items found while making the parser work with other payroll systems (see tests/test_sample_layouts.py).

1. Reserve rows under another system's code are only reported, not added. Payment codes are not standard:
   Hilan uses 3 digits (198), other systems 1-5 digits (01, 40, 0042, 1220, 91001). Option: find reserve rows by
   wording ("מילואים" + "השלמה" -> completion, row 12; other "מילואים" rows -> payment, row 11), flagged for review.
   Needs a decision (it changes money).
2. Some slips do not print the National Insurance gross at all (e.g. Ministry of Education). Row 9 cannot be filled
   from such a slip; it is flagged.
3. Retro-differences column ("הפרשים") on some slips: should it be added to the month's amount?
4. Two employers in the same month: the second slip is skipped as "same month". NI may count income from all
   employers - verify, then sum instead of skip.
5. Template limits: the 2027 maximum/minimum are copies of 2026, and the template ends at 2027.
6. DONE: the value paid can now change from month to month (rows 22/23 from the NI annual letter); a typed current
   day value (F36) still applies to the whole year and overrides the letter.
7. Verify the averaging rule: the template uses the 3 months before each calendar month; NI may use the 3 months
   before the service started for a period that crosses months.
8. Picture-only pay-slips are not supported (the OCR engine cannot read Hebrew).
9. "Payment difference" instead of a day-value difference: compare what NI should have paid in total with what it
   actually paid (the NI annual letter's amounts), not only the difference per day. Today row 29 is days (row 27) x
   (day value - value paid per day), so days NI did not pay for at all (row 28 > 0) count only the difference, not the
   full day value. The user will define this later.
10. DONE (first version): the sheet "תוספת לתקופות שירות" lists every continuous service period with its remainder
    of days / 7 and the bonus (remainder 1-5 -> 0.4 x remainder day values, 6 -> one day value), valued by the start
    month's daily value (row 18), with a summary per month, per year and in total. It is kept separate from the total
    owed. Still open: linking it to the totals (together with item 9), and any further "Mod7" calculations.
