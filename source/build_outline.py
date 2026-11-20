"""
build_outline.py
----------------
Builds the full course outline workbook:
    2026-11-19_course-outline.xlsx

Sheets:
  - Outline : one row per lesson (section, title, what it teaches, type,
              estimated video minutes, estimated student exercise minutes,
              example used, free preview flag, production status).
  - Summary : per-section totals computed with SUMIFS / COUNTIFS formulas,
              course totals, and a check against Udemy's minimum length.
  - Notes   : legend, assumptions and sources of the estimates.

Usage:
    python build_outline.py <output.xlsx>

All minutes are Claude's estimates (blue input cells), to be replaced with
real numbers after the pilot recording of lesson 1.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

OUT = sys.argv[1] if len(sys.argv) > 1 else "outline.xlsx"
FONT = "Arial"

SECTIONS = {
    1: "The Problem",
    2: "The Language",
    3: "The Goal",
    4: "Situation Assessment A: Time and Space",
    5: "Situation Assessment B: Influencing Factors",
    6: "Who Does What",
    7: "Checkpoints and Plan B",
    8: "Cascading",
    9: "Final Exercise and Your Own Case",
}

# (section, lesson #, title EN, what it teaches HE, type, video min, exercise min, example, preview)
L = [
    (1, "1.1", "Busy Is Not Progress", "הבעיה: עסוקים ולא מתקדמים. משימה היא לא מטרה. מה תדע לעשות בסוף", "Video", 4, 0, "ארוחת צהריים", "Yes"),
    (1, "1.2", "How This Course Works", "מבנה הקורס, התרגילים, חוברת העבודה ואיך להפיק ממנו את המרב", "Video", 2, 0, "", "Yes"),
    (1, "1.3", "Exercise: Task or Goal?", "תרגיל 1: סיווג משפטים למשימה או מטרה, וגזירה הפוכה מרשימת המטלות", "Exercise", 0, 10, "עבודה", "No"),

    (2, "2.1", "Goal, Task, Constraint", "שלושת המונחים והגדרה חדה לכל אחד", "Video", 5, 0, "ארוחת צהריים", "Yes"),
    (2, "2.2", "Why \"Mission\" Confuses Everyone", "למה המילה mission מבלבלת בין משימה למטרה, ואיך נדבר בקורס", "Video", 3, 0, "", "No"),
    (2, "2.3", "Quiz: Name It Right", "בוחן קצר על המונחים", "Quiz", 0, 5, "", "No"),

    (3, "3.1", "The Goal Is an End State", "מטרה היא מצב סיום, במילים הכי פשוטות", "Video", 5, 0, "ארוחת צהריים", "No"),
    (3, "3.2", "Keep the Goal Clean", "הפרדת המטרה מזמנים, איכות ודרך ביצוע", "Video", 5, 0, "ארוחת צהריים, עבודה", "No"),
    (3, "3.3", "Exercise: Clean Up the Goal", "ניקוי מטרות עמוסות עד שנשאר רק מצב הסיום", "Exercise", 0, 10, "עבודה", "No"),

    (4, "4.1", "Assess Before You Plan", "למה הערכת מצב באה לפני המשימות", "Video", 4, 0, "ארוחת צהריים", "No"),
    (4, "4.2", "Time", "מתי, כמה זמן פנוי, כמה זמן לוקחת כל אפשרות", "Video", 5, 0, "ארוחת צהריים, עבודה", "No"),
    (4, "4.3", "Space", "איפה כל אחד נמצא, איפה מתבצעת העבודה, מה יש בסביבה", "Video", 4, 0, "ארוחת צהריים, עבודה", "No"),
    (4, "4.4", "Exercise: Time and Space", "הערכת זמן ומרחב על מקרה מהעבודה", "Exercise", 0, 10, "עבודה", "No"),

    (5, "5.1", "What Influencing Factors Are", "גורמים משפיעים: ארבעה סוגים", "Video", 4, 0, "ארוחת צהריים", "No"),
    (5, "5.2", "Conditions: Plan Around What You Can Expect", "תנאים צפויים שלא תלויים בך: מזג אוויר, חגים, תנועה", "Video", 5, 0, "ארוחת צהריים, עבודה", "No"),
    (5, "5.3", "Our Resources and Others Involved", "כוחותינו (אנשים, כסף, ציוד, ידע) וגורמים אחרים", "Video", 5, 0, "ארוחת צהריים, עבודה", "No"),
    (5, "5.4", "What Could Happen: Most Likely, Most Dangerous, Unlikely", "דפ\"א סבירה, מסוכנת ולא סבירה, וכמה להשקיע בכל אחת", "Video", 7, 0, "ארוחת צהריים, עבודה", "No"),
    (5, "5.5", "Exercise: Conditions or Risks?", "סיווג גורמים לתנאי צפוי או לדבר שעלול לקרות", "Exercise", 0, 10, "עבודה", "No"),
    (5, "5.6", "Exercise: Sort the Contingencies", "מיון אירועים לסבירה, מסוכנת ולא סבירה", "Exercise", 0, 10, "עבודה", "No"),
    (5, "5.7", "Exercise: Same Goal, Different Situation", "שינוי נתון אחד בהערכת המצב, ומה קורה למשימות", "Exercise", 0, 10, "עבודה", "No"),

    (6, "6.1", "Tasks in Light of the Goal", "משימה היא פעולה שמקרבת למטרה. מבחן \"כדי ש...\"", "Video", 6, 0, "ארוחת צהריים, עבודה", "No"),
    (6, "6.2", "Breaking Tasks into Subtasks", "פירוק משימות לתת-משימות", "Video", 5, 0, "עבודה", "No"),
    (6, "6.3", "Every Task Has an Owner", "כוחות ומשימות: הצמדת אחראי לכל משימה", "Video", 4, 0, "ארוחת צהריים, עבודה", "No"),
    (6, "6.4", "Exercise: The \"So That\" Test", "מציאת משימות שלא משרתות את המטרה", "Exercise", 0, 10, "עבודה", "No"),
    (6, "6.5", "Exercise: Find the Gaps", "משימות בלי אחראי ואחראי בלי משימה", "Exercise", 0, 10, "עבודה", "No"),

    (7, "7.1", "Check After Every Task", "בקרה: התקרבנו למטרה?", "Video", 5, 0, "ארוחת צהריים, עבודה", "No"),
    (7, "7.2", "When to Switch to Plan B", "המטרה קבועה, המשימות גמישות. מעבר לחלופה", "Video", 5, 0, "עבודה", "No"),
    (7, "7.3", "When the Road Looks Blocked", "לא משנים מטרה, מרחיבים את החיפוש אחרי משימות", "Video", 5, 0, "ארוחת צהריים (בית תמחוי)", "No"),
    (7, "7.4", "Exercise: Checkpoint Mid-Way", "תרחיש שמשתנה תוך כדי ביצוע: להמשיך או לעבור לחלופה", "Exercise", 0, 10, "עבודה", "No"),
    (7, "7.5", "Exercise: Blocked Path", "לפחות 3 משימות חלופיות לאותה מטרה", "Exercise", 0, 10, "עבודה", "No"),

    (8, "8.1", "Your Task Is Their Goal", "גלגול משימות: המשימה של הדרג הפוקד היא המטרה של הדרג שמתחתיו", "Video", 6, 0, "מתקן ספורט בדיוני", "No"),
    (8, "8.2", "Receiving an Assignment", "איך מקבלים משימה מהמנהל והופכים אותה למטרה", "Video", 5, 0, "עבודה", "No"),
    (8, "8.3", "Handing Tasks Down", "איך מורידים משימה לצוות, יחד עם המטרה שמאחוריה", "Video", 5, 0, "עבודה", "No"),
    (8, "8.4", "The Kitchen Brief", "הצגת תרגיל המטבח: ערב שבת במסעדה", "Video", 4, 0, "מטבח מסעדה", "No"),
    (8, "8.5", "Exercise: Saturday Night in the Kitchen", "גלגול משימות מבעל המסעדה עד עמדת הגריל, כולל אירועים באמצע", "Exercise", 0, 20, "מטבח מסעדה", "No"),
    (8, "8.6", "Kitchen Walkthrough", "פתרון לדוגמה והסבר", "Video", 8, 0, "מטבח מסעדה", "No"),
    (8, "8.7", "Exercise: Cascade It", "קבלת משימה מדרג עליון וגזירת משימות לדרג שמתחת", "Exercise", 0, 10, "עבודה", "No"),

    (9, "9.1", "The Tournament Brief", "הצגת התרגיל המסכם: טורניר פוטבול באצטדיון בדיוני", "Video", 5, 0, "טורניר פוטבול", "No"),
    (9, "9.2", "Final Exercise: Football Tournament", "התהליך המלא, כולל גלגול ושינויים באמצע", "Exercise", 0, 30, "טורניר פוטבול", "No"),
    (9, "9.3", "Walkthrough Part 1: Goal and Assessment", "פתרון לדוגמה: מטרה והערכת מצב", "Video", 8, 0, "טורניר פוטבול", "No"),
    (9, "9.4", "Walkthrough Part 2: Tasks, Cascade, Surprises", "פתרון לדוגמה: כוחות ומשימות, גלגול ושינויים", "Video", 8, 0, "טורניר פוטבול", "No"),
    (9, "9.5", "Exercise: Speed Run", "תהליך מלא ב-5 דקות על תרחיש אקראי", "Exercise", 0, 5, "מגוון", "No"),
    (9, "9.6", "Exercise: Your Own Case", "התהליך המלא על אתגר אמיתי מהצוות של התלמיד, עם תבנית", "Exercise", 0, 30, "המקרה של התלמיד", "No"),
    (9, "9.7", "Wrap-Up and Next Steps", "סיכום, דף סיכום להורדה, ואיך ממשיכים ליישם", "Video", 3, 0, "", "No"),
]

# ---------- styles ----------
f_norm = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="17263A")
f_input = Font(name=FONT, size=10, color="0000FF")
fill_head = PatternFill("solid", fgColor="17263A")
fill_sec = PatternFill("solid", fgColor="E9E4D8")
fill_total = PatternFill("solid", fgColor="F4F1EA")
thin = Side(style="thin", color="CFC7B6")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
wrap_r = Alignment(wrap_text=True, vertical="top", horizontal="right", readingOrder=2)
center = Alignment(horizontal="center", vertical="top")

wb = Workbook()

# ---------- Outline ----------
ws = wb.active
ws.title = "Outline"
ws["A1"] = "Goal-First Planning for New Managers: Course Outline (draft v2, 20.11.2026)"
ws["A1"].font = f_title
ws["A2"] = "Blue cells are estimates you can edit. Totals are in the Summary sheet."
ws["A2"].font = Font(name=FONT, size=10, italic=True, color="4A5568")

headers = ["Section #", "Section", "Lesson #", "Lesson title (EN)", "מה לומדים",
           "Type", "Video min (est.)", "Student exercise min (est.)", "Example", "Free preview", "Status"]
HR = 4
for c, t in enumerate(headers, 1):
    cell = ws.cell(row=HR, column=c, value=t)
    cell.font = f_head; cell.fill = fill_head; cell.border = border
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

r = HR + 1
first = r
for (s, num, title, he, typ, vmin, emin, ex, prev) in L:
    vals = [s, SECTIONS[s], num, title, he, typ, vmin, emin, ex, prev, "Not started"]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = f_norm; cell.border = border; cell.alignment = wrap
    for c in (7, 8):
        ws.cell(row=r, column=c).font = f_input
        ws.cell(row=r, column=c).alignment = center
    ws.cell(row=r, column=1).alignment = center
    ws.cell(row=r, column=3).alignment = center
    ws.cell(row=r, column=5).alignment = wrap_r
    ws.cell(row=r, column=9).alignment = wrap_r
    if prev == "Yes":
        ws.cell(row=r, column=10).font = f_bold
    r += 1
last = r - 1

# totals row
ws.cell(row=r, column=4, value="Total").font = f_bold
ws.cell(row=r, column=7, value=f"=SUM(G{first}:G{last})")
ws.cell(row=r, column=8, value=f"=SUM(H{first}:H{last})")
for c in range(1, 12):
    ws.cell(row=r, column=c).fill = fill_total
    ws.cell(row=r, column=c).border = border
for c in (7, 8):
    ws.cell(row=r, column=c).font = f_bold
    ws.cell(row=r, column=c).alignment = center
    ws.cell(row=r, column=c).number_format = "0"

# dropdowns
dv_status = DataValidation(type="list", formula1='"Not started,Script,Slides,Recorded,Edited,Uploaded"', allow_blank=True)
dv_type = DataValidation(type="list", formula1='"Video,Exercise,Quiz,Resource"', allow_blank=True)
dv_prev = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
for dv, col in ((dv_status, "K"), (dv_type, "F"), (dv_prev, "J")):
    ws.add_data_validation(dv)
    dv.add(f"{col}{first}:{col}{last}")

ws.cell(row=HR, column=7).comment = Comment("Claude's estimate per lesson. Replace with real numbers after recording the pilot (lesson 1.1).", "Claude")
ws.cell(row=HR, column=10).comment = Comment("Udemy lets non-buyers watch preview lessons. Rule: previews give the 'why' and the outcomes, not the method steps.", "Claude")

widths = [9, 26, 8, 40, 46, 10, 11, 13, 20, 10, 13]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[chr(64 + i)].width = w
ws.freeze_panes = ws.cell(row=HR + 1, column=4)

# ---------- Summary ----------
sm = wb.create_sheet("Summary")
sm["A1"] = "Summary by section"
sm["A1"].font = f_title
sh = ["Section #", "Section", "Lessons", "Exercises and quizzes", "Video min", "Student exercise min"]
for c, t in enumerate(sh, 1):
    cell = sm.cell(row=3, column=c, value=t)
    cell.font = f_head; cell.fill = fill_head; cell.border = border
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
rng = lambda col: f"Outline!${col}${first}:${col}${last}"
row = 4
for s, name in SECTIONS.items():
    sm.cell(row=row, column=1, value=s)
    sm.cell(row=row, column=2, value=name)
    sm.cell(row=row, column=3, value=f"=COUNTIFS({rng('A')},A{row})")
    sm.cell(row=row, column=4, value=f'=COUNTIFS({rng("A")},A{row},{rng("F")},"Exercise")+COUNTIFS({rng("A")},A{row},{rng("F")},"Quiz")')
    sm.cell(row=row, column=5, value=f"=SUMIFS({rng('G')},{rng('A')},A{row})")
    sm.cell(row=row, column=6, value=f"=SUMIFS({rng('H')},{rng('A')},A{row})")
    for c in range(1, 7):
        sm.cell(row=row, column=c).font = f_norm
        sm.cell(row=row, column=c).border = border
        if c != 2:
            sm.cell(row=row, column=c).alignment = center
    row += 1
s_first, s_last = 4, row - 1
sm.cell(row=row, column=2, value="Total")
for c, col in zip((3, 4, 5, 6), "CDEF"):
    sm.cell(row=row, column=c, value=f"=SUM({col}{s_first}:{col}{s_last})")
for c in range(1, 7):
    sm.cell(row=row, column=c).font = f_bold
    sm.cell(row=row, column=c).fill = fill_total
    sm.cell(row=row, column=c).border = border
    if c != 2:
        sm.cell(row=row, column=c).alignment = center
tot = row

row += 2
facts = [
    ("Total video, hours", f"=E{tot}/60", "0.0"),
    ("Total student time incl. exercises, hours", f"=(E{tot}+F{tot})/60", "0.0"),
    ("Target video length, hours (from plan)", 2.5, "0.0"),
    ("Video vs target", f'=IF(B{row}>=B{row+2}*0.8,"OK","Below target")', "@"),
    ("Udemy minimum video, min (per guides, unverified)", 30, "0"),
    ("Meets Udemy minimum?", f'=IF(E{tot}>=B{row+4},"Yes","No")', "@"),
    ("Free preview lessons", f'=COUNTIFS({rng("J")},"Yes")', "0"),
]
for i, (label, val, fmt) in enumerate(facts):
    rr = row + i
    sm.cell(row=rr, column=1, value=label).font = f_bold
    cell = sm.cell(row=rr, column=2, value=val)
    cell.number_format = fmt
    cell.font = f_input if not (isinstance(val, str) and val.startswith("=")) else f_norm
    cell.alignment = center
sm.cell(row=row + 2, column=3, value="From the planning document: target of 2 to 3 hours of video.").font = Font(name=FONT, size=9, italic=True)
sm.cell(row=row + 4, column=3, value="Third-party guides: 30 minutes and 5 lectures. Not checked against Udemy's official page.").font = Font(name=FONT, size=9, italic=True)
sm.column_dimensions["A"].width = 44
sm.column_dimensions["B"].width = 40
for col in "CDEF":
    sm.column_dimensions[col].width = 14

# ---------- Notes ----------
nt = wb.create_sheet("Notes")
notes = [
    ("Legend", ""),
    ("Blue text", "Estimates you can edit (minutes, targets)."),
    ("Black text", "Formulas or fixed text. Do not edit formulas in Summary."),
    ("Status column", "Pick from the list: Not started, Script, Slides, Recorded, Edited, Uploaded."),
    ("", ""),
    ("Assumptions", ""),
    ("Minutes", "All minutes are Claude's estimates, not measured. Replace them after the pilot recording."),
    ("Lesson length", "Video lessons 2 to 8 minutes, one idea per lesson (rule from the planning document)."),
    ("Previews", "Only lessons 1.1, 1.2 and 2.1 are free previews. They give the 'why' and outcomes, not the method steps."),
    ("Method structure", "Goal, Situation Assessment (Time, Space, Influencing Factors), Who Does What, Checkpoints and Plan B, Cascading."),
    ("Examples", "All examples are fictional. No real venues, teams or employers."),
    ("", ""),
    ("Example row", "Row 5 in Outline (lesson 1.1) shows the expected format for a new lesson."),
]
for i, (a, b) in enumerate(notes, 1):
    nt.cell(row=i, column=1, value=a).font = f_bold
    nt.cell(row=i, column=2, value=b).font = f_norm
nt.column_dimensions["A"].width = 20
nt.column_dimensions["B"].width = 110

wb.save(OUT)
print("written:", OUT)
