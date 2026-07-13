"""
Csapat nyilvántartó Excel sablon generálása.
Létrehoz egy xlsx fájlt:
  - Legenda            - kódok, színek, használati útmutató
  - Napi jelenlét YYYY  - évenként külön munkalap, napi szintű szabadság/beteg/ünnep (16 fő)
  - Heti beosztás       - egy folytonos munkalap az összes évre, heti váltott műszak
                          (Éjjel / Délután / Délelőtt 3-hetes rotáció, mindenkire egyszerre)

Ez a fájl a "forrás" adat, amit a csoportvezető szerkeszt. A generate_html.py
script ebből készíti a csapat számára megtekinthető weboldalt.
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

YEARS = [2025, 2026, 2027, 2028]
TODAY = datetime.date(2026, 7, 13)
N_EMPLOYEES = 16
EMPLOYEES = [f"Munkatárs {i:02d}" for i in range(1, N_EMPLOYEES + 1)]

DAILY_CODES = {
    "":   ("Munkanap", "FFFFFF"),
    "SZ": ("Szabadság", "FFC000"),
    "B":  ("Betegszabadság", "FF6B6B"),
    "OH": ("Home office", "9DC3E6"),
    "UN": ("Ünnepnap", "C39BD3"),
    "P":  ("Pihenőnap", "A9D18E"),
}
SUMMARY_CODES = [("SZ", "Szabadság"), ("B", "Betegszabadság"), ("OH", "Home office"), ("UN", "Ünnepnap"), ("P", "Pihenőnap")]

# Váltott műszakrend: 3 hetes rotáció, az egész csapatra egyszerre érvényes.
# A TODAY-t tartalmazó hét = Éjjel, utána Délután, utána Délelőtt, majd újra Éjjel...
SHIFT_CYCLE = ["Éjjel", "Délután", "Délelőtt"]
WEEKLY_EXTRA_CODES = ["Szabadság", "Home office", "Pihenő", "Egyéb"]
WEEKLY_CODES = SHIFT_CYCLE + WEEKLY_EXTRA_CODES
WEEKLY_COLORS = {
    "Éjjel": "5B4B8A", "Délután": "F4B183", "Délelőtt": "FFE699",
    "Szabadság": "FFC000", "Home office": "9DC3E6", "Pihenő": "70AD47", "Egyéb": "D9D9D9",
}

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
WEEKEND_FILL = PatternFill("solid", fgColor="E7E6E6")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HU_MONTHS = ["jan", "feb", "márc", "ápr", "máj", "jún", "júl", "aug", "szep", "okt", "nov", "dec"]

wb = Workbook()

# ---------------------------------------------------------------- Legenda
ws_legend = wb.active
ws_legend.title = "Legenda"
ws_legend.sheet_view.showGridLines = False
ws_legend["B2"] = "Csapat nyilvántartó — használati útmutató"
ws_legend["B2"].font = Font(name=FONT_NAME, bold=True, size=14)

ws_legend["B4"] = "1) Napi jelenlét (évenkénti lapok): minden nap / minden dolgozó cellájába írd be a kódot (legördülő listából is választható)."
ws_legend["B5"] = "2) Heti beosztás: a váltott műszakrend (Éjjel/Délután/Délelőtt) automatikusan ki van töltve 3 hetes rotációban, mindenkire egyszerre. Ahol valaki kivétel (szabin van, pihenőn, home office-ban stb.), azt felülírhatod az adott cellában — ez automatikusan bekerül a weboldal 'Kivételek' listájába."
ws_legend["B6"] = "3) Éves összesítő: automatikusan számolja évenként/dolgozónként a szabadság, betegszabadság, home office, ünnepnap és pihenőnap napok számát (a napi jelenlét lapok alapján)."
ws_legend["B7"] = "4) Mentsd el a fájlt, majd futtasd a generate_html.py scriptet — ez elkészíti a csapat számára a megtekinthető weboldalt (csapat_naptar.html)."
ws_legend["B8"] = "5) Az elkészült html fájlt oszd meg a csapattal (email, közös meghajtó, intranet) — ők csak megnézik, nem szerkesztik."
for r in range(4, 9):
    ws_legend[f"B{r}"].font = Font(name=FONT_NAME, size=10)
    ws_legend[f"B{r}"].alignment = Alignment(wrap_text=True)

ws_legend["B9"] = "Napi kódok:"
ws_legend["B9"].font = Font(name=FONT_NAME, bold=True, size=11)
row = 10
for code, (label, color) in DAILY_CODES.items():
    ws_legend[f"B{row}"] = code if code else "(üres)"
    ws_legend[f"C{row}"] = label
    ws_legend[f"B{row}"].fill = PatternFill("solid", fgColor=color)
    ws_legend[f"B{row}"].font = Font(name=FONT_NAME, size=10)
    ws_legend[f"C{row}"].font = Font(name=FONT_NAME, size=10)
    row += 1

row += 1
ws_legend[f"B{row}"] = "Heti műszakrend (3 hetes rotáció, automatikus):"
ws_legend[f"B{row}"].font = Font(name=FONT_NAME, bold=True, size=11)
row += 1
for shift in SHIFT_CYCLE:
    ws_legend[f"B{row}"] = shift
    ws_legend[f"B{row}"].fill = PatternFill("solid", fgColor=WEEKLY_COLORS[shift])
    ws_legend[f"B{row}"].font = Font(name=FONT_NAME, size=10)
    row += 1

row += 1
ws_legend[f"B{row}"] = "Kivétel esetén választható (felülírja a rotációt, bekerül a kivétel listába):"
ws_legend[f"B{row}"].font = Font(name=FONT_NAME, bold=True, size=11)
row += 1
for code in WEEKLY_EXTRA_CODES:
    ws_legend[f"B{row}"] = code
    ws_legend[f"B{row}"].fill = PatternFill("solid", fgColor=WEEKLY_COLORS[code])
    ws_legend[f"B{row}"].font = Font(name=FONT_NAME, size=10)
    row += 1

ws_legend.column_dimensions["A"].width = 2
ws_legend.column_dimensions["B"].width = 60
ws_legend.column_dimensions["C"].width = 22

# ---------------------------------------------------------------- Napi jelenlét (évenként)
for YEAR in YEARS:
    ws = wb.create_sheet(f"Napi jelenlét {YEAR}")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B2"

    start = datetime.date(YEAR, 1, 1)
    end = datetime.date(YEAR, 12, 31)
    n_days = (end - start).days + 1

    ws["A1"] = "Munkatárs"
    ws["A1"].font = HEADER_FONT
    ws["A1"].fill = HEADER_FILL
    ws.column_dimensions["A"].width = 16

    for d in range(n_days):
        day = start + datetime.timedelta(days=d)
        col = d + 2
        letter = get_column_letter(col)
        cell = ws.cell(row=1, column=col, value=f"{HU_MONTHS[day.month-1]}.{day.day}")
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(text_rotation=90, horizontal="center")
        ws.column_dimensions[letter].width = 3.2
        if day.weekday() >= 5:  # szombat, vasárnap
            for r in range(2, N_EMPLOYEES + 2):
                ws.cell(row=r, column=col).fill = WEEKEND_FILL

    for i, name in enumerate(EMPLOYEES):
        r = i + 2
        ws.cell(row=r, column=1, value=name).font = Font(name=FONT_NAME, size=10)

    data_range = f"B2:{get_column_letter(n_days + 1)}{N_EMPLOYEES + 1}"
    dv = DataValidation(type="list", formula1='"" ,SZ,B,OH,UN,P', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(data_range)

    for code, (label, color) in DAILY_CODES.items():
        if not code:
            continue
        rule = CellIsRule(operator="equal", formula=[f'"{code}"'], fill=PatternFill("solid", fgColor=color))
        ws.conditional_formatting.add(data_range, rule)

    # minta adat csak a 2026-os (jelenlegi) lapon, 2026.07 hónap körül
    if YEAR == 2026:
        sample_start_col = 2 + (datetime.date(YEAR, 7, 1) - start).days
        ws.cell(row=2, column=sample_start_col + 5, value="SZ")
        ws.cell(row=2, column=sample_start_col + 6, value="SZ")
        ws.cell(row=2, column=sample_start_col + 7, value="SZ")
        ws.cell(row=3, column=sample_start_col + 1, value="B")
        ws.cell(row=3, column=sample_start_col + 2, value="B")
        ws.cell(row=4, column=sample_start_col + 10, value="OH")
        ws.cell(row=4, column=sample_start_col + 11, value="OH")
        ws.cell(row=5, column=sample_start_col - 20, value="SZ")  # júniusi minta

# ---------------------------------------------------------------- Heti beosztás (folytonos, minden évre)
ws2 = wb.create_sheet("Heti beosztás")
ws2.sheet_view.showGridLines = False
ws2.freeze_panes = "B2"
ws2["A1"] = "Munkatárs"
ws2["A1"].font = HEADER_FONT
ws2["A1"].fill = HEADER_FILL
ws2.column_dimensions["A"].width = 16

overall_start = datetime.date(YEARS[0], 1, 1)
overall_end = datetime.date(YEARS[-1], 12, 31)

# aktuális hét hétfője = a rotáció "nulladik" hete (Éjjel)
today_monday = TODAY - datetime.timedelta(days=TODAY.weekday())

week_starts = []
d = overall_start - datetime.timedelta(days=overall_start.weekday())
while d <= overall_end:
    week_starts.append(d)
    d += datetime.timedelta(days=7)

for w, wk in enumerate(week_starts):
    col = w + 2
    letter = get_column_letter(col)
    cell = ws2.cell(row=1, column=col, value=f"{wk.year}.{wk.month}.{wk.day}")
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(text_rotation=90, horizontal="center")
    ws2.column_dimensions[letter].width = 4.5

for i, name in enumerate(EMPLOYEES):
    r = i + 2
    ws2.cell(row=r, column=1, value=name).font = Font(name=FONT_NAME, size=10)

dv2 = DataValidation(type="list", formula1='"' + ",".join(WEEKLY_CODES) + '"', allow_blank=True)
ws2.add_data_validation(dv2)
weekly_range = f"B2:{get_column_letter(len(week_starts)+1)}{N_EMPLOYEES+1}"
dv2.add(weekly_range)

for shift, color in WEEKLY_COLORS.items():
    rule = CellIsRule(operator="equal", formula=[f'"{shift}"'], fill=PatternFill("solid", fgColor=color))
    ws2.conditional_formatting.add(weekly_range, rule)

# automatikus rotáció kitöltése mindenkinek minden hétre
week_offset_to_shift = {}
for w, wk in enumerate(week_starts):
    offset = (wk - today_monday).days // 7
    shift = SHIFT_CYCLE[offset % 3]
    week_offset_to_shift[w] = shift

for i in range(N_EMPLOYEES):
    r = i + 2
    for w in range(len(week_starts)):
        ws2.cell(row=r, column=w + 2, value=week_offset_to_shift[w])

# minta kivételek: néhány felülírt hét, hogy a "Kivételek" lista lásson valamit
cur_week_col = week_starts.index(today_monday) + 2
ws2.cell(row=2, column=cur_week_col, value="Szabadság")       # Munkatárs 01: e héten szabin, nem Éjjel
ws2.cell(row=3, column=cur_week_col + 1, value="Pihenő")      # Munkatárs 02: jövő héten pihenő
ws2.cell(row=4, column=cur_week_col + 2, value="Home office") # Munkatárs 03: az azutáni héten home office

# ---------------------------------------------------------------- Éves összesítő
ws3 = wb.create_sheet("Éves összesítő")
ws3.sheet_view.showGridLines = False
ws3.freeze_panes = "B2"
ws3["A1"] = "Munkatárs"
ws3["A1"].font = HEADER_FONT
ws3["A1"].fill = HEADER_FILL
ws3.column_dimensions["A"].width = 16

col = 2
year_last_col = {}
for YEAR in YEARS:
    n_days = (datetime.date(YEAR, 12, 31) - datetime.date(YEAR, 1, 1)).days + 1
    year_last_col[YEAR] = get_column_letter(n_days + 1)
    for code, label in SUMMARY_CODES:
        letter = get_column_letter(col)
        cell = ws3.cell(row=1, column=col, value=f"{YEAR} {label}")
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(text_rotation=90, horizontal="center", wrap_text=True)
        ws3.column_dimensions[letter].width = 4.5
        col += 1

for i, name in enumerate(EMPLOYEES):
    r = i + 2
    ws3.cell(row=r, column=1, value=name).font = Font(name=FONT_NAME, size=10)
    col = 2
    for YEAR in YEARS:
        sheet_name = f"Napi jelenlét {YEAR}"
        last_col = year_last_col[YEAR]
        rng = f"'{sheet_name}'!B{r}:{last_col}{r}"
        for code, label in SUMMARY_CODES:
            c = ws3.cell(row=r, column=col, value=f'=COUNTIF({rng},"{code}")')
            c.number_format = "0"
            c.font = Font(name=FONT_NAME, size=10)
            c.alignment = Alignment(horizontal="center")
            col += 1

wb.save("csapat_nyilvantartas.xlsx")
print("OK: csapat_nyilvantartas.xlsx elkészült")
print("Aktuális hét (", today_monday, ") műszakja:", week_offset_to_shift[week_starts.index(today_monday)])
