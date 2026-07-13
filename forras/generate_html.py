"""
Beolvassa a csapat_nyilvantartas.xlsx fájlt és elkészíti a csapat_naptar.html
weboldalt — ezt lehet megosztani a csapattal (csak megtekintésre).

Futtatás:  python3 generate_html.py
Bemenet:   csapat_nyilvantartas.xlsx  (a csoportvezető szerkeszti)
Kimenet:   csapat_naptar.html        (ezt kapja a csapat)
"""
import json
import re
import datetime
from openpyxl import load_workbook

SRC = "csapat_nyilvantartas.xlsx"
OUT = "csapat_naptar.html"
TODAY = datetime.date(2026, 7, 13)

DAILY_CODES = {
    "":   {"label": "Munkanap", "color": "#ffffff"},
    "SZ": {"label": "Szabadság", "color": "#f5a623"},
    "B":  {"label": "Betegszabadság", "color": "#ef5350"},
    "OH": {"label": "Home office", "color": "#5b9bd5"},
    "UN": {"label": "Ünnepnap", "color": "#ab7fd1"},
    "P":  {"label": "Pihenőnap", "color": "#7cb872"},
}
SUMMARY_CODES = ["Szabadság", "Betegszabadság", "Home office", "Ünnepnap", "Pihenőnap"]
QUOTA_LABELS = ["Szabadságkeret", "Hátralévő szabadság"]
SHIFT_CYCLE = ["Éjjel", "Délután", "Délelőtt"]
SHIFT_COLORS = {
    "Éjjel": "#5b4b8a",
    "Délután": "#ef8f4e",
    "Délelőtt": "#f0c93b",
    "Szabadság": "#f5a623",
    "Home office": "#5b9bd5",
    "Pihenő": "#4d9a4d",
    "Egyéb": "#9aa1ab",
}

wb = load_workbook(SRC, data_only=True)

# --- napi jelenlét: minden "Napi jelenlét YYYY" lap beolvasása ---
years = []
daily_by_year = {}
for sheet_name in wb.sheetnames:
    m = re.match(r"Napi jelenlét (\d{4})", sheet_name)
    if not m:
        continue
    year = int(m.group(1))
    years.append(year)
    ws = wb[sheet_name]
    max_col = ws.max_column
    dates = [ws.cell(row=1, column=c).value for c in range(2, max_col + 1)]

    employees = []
    daily = {}
    r = 2
    while True:
        name = ws.cell(row=r, column=1).value
        if not name or name in ("Jelenlévők száma", "Minimum létszám:"):
            break
        employees.append(name)
        row_codes = []
        for c in range(2, max_col + 1):
            v = ws.cell(row=r, column=c).value
            row_codes.append((v or "").strip().upper())
        daily[name] = row_codes
        r += 1

    staff_row = r
    threshold_row = r + 1
    staff_counts = [ws.cell(row=staff_row, column=c).value for c in range(2, max_col + 1)]
    min_staffing = ws.cell(row=threshold_row, column=2).value or 0

    daily_by_year[year] = {
        "dates": dates,
        "employees": employees,
        "daily": daily,
        "staff_counts": staff_counts,
        "min_staffing": min_staffing,
    }
years.sort()

all_employees = daily_by_year[years[0]]["employees"]

# --- heti beosztás beolvasása (folytonos, évekre bontva utólag a JS-ben) ---
ws2 = wb["Heti beosztás"]
max_col2 = ws2.max_column
week_labels = []  # "YYYY.M.D"
week_years = []
for c in range(2, max_col2 + 1):
    label = ws2.cell(row=1, column=c).value
    week_labels.append(label)
    week_years.append(int(str(label).split(".")[0]))

weekly = {}
r = 2
while ws2.cell(row=r, column=1).value:
    name = ws2.cell(row=r, column=1).value
    row_vals = []
    for c in range(2, max_col2 + 1):
        v = ws2.cell(row=r, column=c).value
        row_vals.append(v or "")
    weekly[name] = row_vals
    r += 1

# --- kivétel lista: ahol a tényleges heti bejegyzés eltér az elvárt rotációtól ---
today_monday = TODAY - datetime.timedelta(days=TODAY.weekday())

exceptions = []
for name in all_employees:
    vals = weekly.get(name, [])
    for i, label in enumerate(week_labels):
        y, mo, da = [int(x) for x in str(label).split(".")]
        wk = datetime.date(y, mo, da)
        offset = (wk - today_monday).days // 7
        expected = SHIFT_CYCLE[offset % 3]
        actual = vals[i] if i < len(vals) else ""
        if actual and actual != expected:
            exceptions.append({
                "employee": name,
                "week": f"{y}.{mo}.{da}",
                "year": y,
                "expected": expected,
                "actual": actual,
            })

# --- éves összesítő beolvasása (COUNTIF + szabadságkeret formulák eredménye) ---
ws3 = wb["Éves összesítő"]
max_col3 = ws3.max_column
summary_headers = [ws3.cell(row=1, column=c).value for c in range(2, max_col3 + 1)]
summary = {}
r = 2
while ws3.cell(row=r, column=1).value:
    name = ws3.cell(row=r, column=1).value
    row_by_year = {}
    for idx, header in enumerate(summary_headers):
        y_str, label = str(header).split(" ", 1)
        y = int(y_str)
        v = ws3.cell(row=r, column=idx + 2).value or 0
        row_by_year.setdefault(y, {})[label] = v
    summary[name] = row_by_year
    r += 1

# a mai nap "híd" adatok a JS-hez (napi + heti kiemeléshez)
today_huabbr = ["jan", "feb", "márc", "ápr", "máj", "jún", "júl", "aug", "szep", "okt", "nov", "dec"][TODAY.month - 1]
today_label = f"{today_huabbr}.{TODAY.day}"
today_week_label = f"{today_monday.year}.{today_monday.month}.{today_monday.day}"

data = {
    "generated": datetime.datetime.now().strftime("%Y.%m.%d %H:%M"),
    "years": years,
    "employees": all_employees,
    "daily_by_year": daily_by_year,
    "week_labels": week_labels,
    "week_years": week_years,
    "weekly": weekly,
    "codes": DAILY_CODES,
    "shift_colors": SHIFT_COLORS,
    "exceptions": exceptions,
    "summary": summary,
    "summary_codes": SUMMARY_CODES,
    "quota_labels": QUOTA_LABELS,
    "today_year": TODAY.year,
    "today_label": today_label,
    "today_week_label": today_week_label,
    "today_display": f"{TODAY.year}. {['jan.','febr.','márc.','ápr.','máj.','jún.','júl.','aug.','szept.','okt.','nov.','dec.'][TODAY.month-1]} {TODAY.day}.",
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="hu">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Csapat naptár</title>
<style>
  :root {
    --border: #e5e7eb;
    --text: #1f2430;
    --muted: #6b7280;
    --accent: #4f46e5;
    --accent2: #7c3aed;
    --bg: #f4f5f9;
    --card: #ffffff;
    --radius: 14px;
    --shadow: 0 1px 3px rgba(17, 24, 39, 0.06), 0 1px 2px rgba(17, 24, 39, 0.08);
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    margin: 0;
    background: var(--bg);
    color: var(--text);
  }
  header {
    background: linear-gradient(120deg, var(--accent) 0%, var(--accent2) 100%);
    color: white;
    padding: 26px 32px 30px;
    display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 14px;
  }
  header .titles h1 { margin: 0 0 6px 0; font-size: 23px; font-weight: 700; letter-spacing: -0.01em; }
  header .titles p { margin: 0; opacity: 0.88; font-size: 13px; }
  header .today-badge {
    background: rgba(255,255,255,0.16); border: 1px solid rgba(255,255,255,0.3);
    border-radius: 999px; padding: 7px 16px; font-size: 13px; font-weight: 600; backdrop-filter: blur(2px);
  }
  .wrap { max-width: 1440px; margin: -14px auto 0; padding: 0 22px 30px; }
  .tabs {
    display: flex; gap: 6px; margin-bottom: 16px; background: var(--card); border-radius: 999px;
    padding: 6px; box-shadow: var(--shadow); width: fit-content;
  }
  .tab-btn {
    display: flex; align-items: center; gap: 7px;
    padding: 9px 18px; border-radius: 999px; border: none;
    background: transparent; cursor: pointer; font-size: 13.5px; font-weight: 600;
    color: var(--muted); transition: all .15s ease;
  }
  .tab-btn:hover { background: #f1f1f6; }
  .tab-btn.active { background: linear-gradient(120deg, var(--accent), var(--accent2)); color: white; box-shadow: 0 2px 6px rgba(79,70,229,.35); }
  .tab-btn svg { width: 15px; height: 15px; flex-shrink: 0; }
  .panel { display: none; background: var(--card); border-radius: var(--radius); box-shadow: var(--shadow); padding: 20px; }
  .panel.active { display: block; }
  .controls { display: flex; gap: 14px; align-items: center; margin-bottom: 16px; flex-wrap: wrap; }
  .controls label { font-size: 13px; font-weight: 600; color: var(--muted); }
  select, input[type=text] {
    padding: 8px 12px; border-radius: 8px; border: 1px solid var(--border); font-size: 13.5px;
    background: #fafafd; outline: none;
  }
  select:focus, input[type=text]:focus { border-color: var(--accent); background: white; }
  .btn-icon {
    display: inline-flex; align-items: center; gap: 6px; margin-left: auto;
    padding: 8px 14px; border-radius: 8px; border: 1px solid var(--border); background: white;
    cursor: pointer; font-size: 13px; font-weight: 600; color: var(--text);
  }
  .btn-icon:hover { background: #f4f4fb; border-color: var(--accent); }
  .btn-icon svg { width: 14px; height: 14px; }
  table { border-collapse: collapse; font-size: 12px; width: 100%; }
  th, td {
    border: 1px solid var(--border); padding: 4px 5px; text-align: center; white-space: nowrap;
  }
  th.name-col, td.name-col {
    position: sticky; left: 0; background: white; text-align: left; z-index: 2;
    min-width: 150px; font-weight: 600; font-size: 12.5px;
  }
  .avatar {
    display: inline-flex; align-items: center; justify-content: center;
    width: 22px; height: 22px; border-radius: 50%; background: #eef0fb; color: var(--accent);
    font-size: 10px; font-weight: 700; margin-right: 7px; vertical-align: middle;
  }
  thead th { position: sticky; top: 0; background: #383e73; color: white; z-index: 3; font-weight: 600; font-size: 11px; }
  thead th.name-col { z-index: 4; background: #383e73; }
  tbody tr:nth-child(even) td:not(.name-col):not([style]) { background: #fafafc; }
  tbody tr:hover td:not(.name-col) { filter: brightness(0.96); }
  tbody tr:hover td.name-col { background: #f4f4fb; }
  .table-scroll { overflow: auto; max-height: 65vh; border: 1px solid var(--border); border-radius: 10px; }
  .staff-row td { font-weight: 700; background: #eef0f7; }
  .staff-row td.alert { background: #ef5350 !important; color: white; }
  .legend { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 16px; }
  .legend span.chip {
    display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; padding: 5px 12px;
    border-radius: 999px; background: #f4f4fb; border: 1px solid var(--border);
  }
  .legend span.swatch { display: inline-block; width: 11px; height: 11px; border-radius: 50%; border: 1px solid rgba(0,0,0,0.1); }
  footer { text-align: center; color: var(--muted); font-size: 12px; padding: 20px 24px 30px; }
  .year-bar {
    display: flex; justify-content: center; gap: 8px; margin: 26px auto 4px; flex-wrap: wrap;
  }
  .year-btn {
    padding: 9px 22px; border-radius: 999px; border: 1px solid var(--border);
    background: white; cursor: pointer; font-size: 13.5px; font-weight: 700; color: var(--muted);
    box-shadow: var(--shadow); transition: all .15s ease;
  }
  .year-btn:hover { color: var(--accent); }
  .year-btn.active { background: linear-gradient(120deg, var(--accent), var(--accent2)); color: white; border-color: transparent; }
  .summary-cards { display: flex; flex-wrap: wrap; gap: 14px; }
  .summary-card {
    flex: 1 1 210px; border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px;
    background: linear-gradient(180deg, #fafbff 0%, #ffffff 100%);
  }
  .summary-card h3 { margin: 0 0 12px 0; font-size: 14px; display: flex; align-items: center; }
  .summary-card table { width: 100%; font-size: 12.5px; }
  .summary-card td { border: none; padding: 3px 0; text-align: left; }
  .summary-card td.num { text-align: right; font-weight: 700; }
  .summary-card .quota-box {
    margin-top: 10px; padding-top: 10px; border-top: 1px dashed var(--border);
    display: flex; justify-content: space-between; align-items: baseline;
  }
  .summary-card .quota-remaining { font-size: 20px; font-weight: 800; color: var(--accent); }
  .summary-card .quota-caption { font-size: 11px; color: var(--muted); }
  .empty-note { color: var(--muted); font-size: 13px; padding: 10px 0 16px; }
  td.today-col { box-shadow: inset 0 0 0 2px var(--accent); font-weight: 700; }
  th.today-col { background: #2f2b6b !important; }
  @media print {
    header .today-badge, .tabs, .controls, .legend, .year-bar, footer, .btn-icon { display: none !important; }
    body { background: white; }
    .wrap { margin: 0; padding: 0; max-width: 100%; }
    .panel { display: none !important; box-shadow: none; border-radius: 0; padding: 0; }
    .panel.active { display: block !important; }
    .table-scroll { max-height: none; overflow: visible; border: none; }
    header { background: white !important; color: black; padding: 10px 0; }
    thead th { background: #383e73 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
</style>
</head>
<body>
<header>
  <div class="titles">
    <h1>Csapat naptár — szabadság, betegség, munkarend</h1>
    <p>Utolsó frissítés: __GENERATED__ &nbsp;·&nbsp; csak megtekintésre</p>
  </div>
  <div class="today-badge">Ma: __TODAY_DISPLAY__</div>
</header>
<div class="wrap">
  <div class="tabs">
    <button class="tab-btn active" data-tab="daily">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>
      Napi jelenlét
    </button>
    <button class="tab-btn" data-tab="weekly">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>
      Heti beosztás
    </button>
    <button class="tab-btn" data-tab="exceptions">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/><path d="M12 9v4M12 17h.01"/></svg>
      Kivételek
    </button>
    <button class="tab-btn" data-tab="summary">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 3v18h18M8 17V10M13 17V6M18 17v-4"/></svg>
      Éves összesítő
    </button>
  </div>

  <div id="daily" class="panel active">
    <div class="controls">
      <label>Hónap</label>
      <select id="monthSelect"></select>
      <label>Keresés</label>
      <input type="text" id="searchDaily" placeholder="Munkatárs neve...">
      <button class="btn-icon" onclick="window.print()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 14h12v8H6z"/></svg>
        Nyomtatás / PDF
      </button>
    </div>
    <div class="table-scroll"><table id="dailyTable"></table></div>
    <div class="legend" id="legend"></div>
  </div>

  <div id="weekly" class="panel">
    <div class="controls">
      <label>Keresés</label>
      <input type="text" id="searchWeekly" placeholder="Munkatárs neve...">
      <button class="btn-icon" onclick="window.print()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 14h12v8H6z"/></svg>
        Nyomtatás / PDF
      </button>
    </div>
    <div class="table-scroll"><table id="weeklyTable"></table></div>
    <div class="legend" id="legendWeekly"></div>
  </div>

  <div id="exceptions" class="panel">
    <div class="controls">
      <label>Keresés</label>
      <input type="text" id="searchExc" placeholder="Munkatárs neve...">
    </div>
    <p class="empty-note">Azok a hetek, ahol a tényleges beosztás eltér az automatikus Éjjel→Délután→Délelőtt rotációtól (pl. szabadság, pihenő, home office, egyéb).</p>
    <div class="table-scroll"><table id="excTable"></table></div>
  </div>

  <div id="summary" class="panel">
    <div class="summary-cards" id="summaryCards"></div>
  </div>

  <div class="year-bar" id="yearBar"></div>
</div>
<footer>Generálva a csapat_nyilvantartas.xlsx fájlból &middot; a szerkesztést a csoportvezető végzi</footer>

<script>
const DATA = __DATA_JSON__;
let currentYear = DATA.years.includes(DATA.today_year) ? DATA.today_year : DATA.years[0];

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
});

// --- year bar (alul, minden nézetet vezérel) ---
const yearBar = document.getElementById('yearBar');
DATA.years.forEach(y => {
  const btn = document.createElement('button');
  btn.className = 'year-btn' + (y === currentYear ? ' active' : '');
  btn.textContent = y;
  btn.dataset.year = y;
  btn.addEventListener('click', () => {
    currentYear = y;
    document.querySelectorAll('.year-btn').forEach(b => b.classList.toggle('active', parseInt(b.dataset.year) === y));
    renderDaily(); renderWeekly(); renderExceptions(); renderSummary();
  });
  yearBar.appendChild(btn);
});

// --- legend (napi) ---
const legendEl = document.getElementById('legend');
Object.entries(DATA.codes).forEach(([code, info]) => {
  if (!code) return;
  const el = document.createElement('span');
  el.className = 'chip';
  el.innerHTML = `<span class="swatch" style="background:${info.color}"></span>${info.label} (${code})`;
  legendEl.appendChild(el);
});

// --- legend (heti / műszak) ---
const legendWeeklyEl = document.getElementById('legendWeekly');
Object.entries(DATA.shift_colors).forEach(([label, color]) => {
  const el = document.createElement('span');
  el.className = 'chip';
  el.innerHTML = `<span class="swatch" style="background:${color}"></span>${label}`;
  legendWeeklyEl.appendChild(el);
});

// --- month select ---
const monthNames = ['Január','Február','Március','Április','Május','Június','Július','Augusztus','Szeptember','Október','November','December'];
const huMonthAbbr = ['jan','feb','márc','ápr','máj','jún','júl','aug','szep','okt','nov','dec'];
const monthSelect = document.getElementById('monthSelect');
monthNames.forEach((name, idx) => {
  const opt = document.createElement('option');
  opt.value = idx;
  opt.textContent = name;
  monthSelect.appendChild(opt);
});
monthSelect.value = new Date(2000, huMonthAbbr.indexOf(DATA.today_label.split('.')[0])).getMonth();

function monthOfIndex(dates, i) {
  return huMonthAbbr.indexOf(dates[i].split('.')[0]);
}

function initials(name) {
  return name.replace('Munkatárs ', '').trim();
}

function renderDaily() {
  const yearData = DATA.daily_by_year[currentYear];
  if (!yearData) return;
  const monthIdx = parseInt(monthSelect.value, 10);
  const search = document.getElementById('searchDaily').value.trim().toLowerCase();
  const colIndexes = [];
  yearData.dates.forEach((d, i) => { if (monthOfIndex(yearData.dates, i) === monthIdx) colIndexes.push(i); });

  const isToday = (i) => currentYear === DATA.today_year && yearData.dates[i] === DATA.today_label;

  let html = '<thead><tr><th class="name-col">Munkatárs</th>';
  colIndexes.forEach(i => {
    const label = yearData.dates[i].split('.')[1];
    html += `<th class="${isToday(i) ? 'today-col' : ''}">${label}</th>`;
  });
  html += '</tr></thead><tbody>';

  yearData.employees.forEach(name => {
    if (search && !name.toLowerCase().includes(search)) return;
    html += `<tr><td class="name-col"><span class="avatar">${initials(name)}</span>${name}</td>`;
    colIndexes.forEach(i => {
      const code = (yearData.daily[name] && yearData.daily[name][i]) || '';
      const info = DATA.codes[code];
      const bg = (info && code) ? `background:${info.color}${code==='B'?';color:white':''};` : '';
      const cls = isToday(i) ? 'today-col' : '';
      html += `<td class="${cls}" style="${bg}" title="${info ? info.label : ''}">${code}</td>`;
    });
    html += '</tr>';
  });

  // jelenlévők száma sor
  html += `<tr class="staff-row"><td class="name-col">Jelenlévők száma</td>`;
  colIndexes.forEach(i => {
    const count = yearData.staff_counts[i];
    const low = count !== null && count < yearData.min_staffing;
    html += `<td class="${low ? 'alert' : ''}" title="Minimum létszám: ${yearData.min_staffing}">${count ?? ''}</td>`;
  });
  html += '</tr>';

  html += '</tbody>';
  document.getElementById('dailyTable').innerHTML = html;
}

function renderWeekly() {
  const search = document.getElementById('searchWeekly').value.trim().toLowerCase();
  const colIndexes = [];
  DATA.week_labels.forEach((w, i) => { if (DATA.week_years[i] === currentYear) colIndexes.push(i); });

  let html = '<thead><tr><th class="name-col">Munkatárs</th>';
  colIndexes.forEach(i => {
    const parts = DATA.week_labels[i].split('.');
    const isToday = DATA.week_labels[i] === DATA.today_week_label;
    html += `<th class="${isToday ? 'today-col' : ''}">${parts[1]}.${parts[2]}</th>`;
  });
  html += '</tr></thead><tbody>';
  DATA.employees.forEach(name => {
    if (search && !name.toLowerCase().includes(search)) return;
    const vals = DATA.weekly[name] || [];
    html += `<tr><td class="name-col"><span class="avatar">${initials(name)}</span>${name}</td>`;
    colIndexes.forEach(i => {
      const v = vals[i] || '';
      const color = DATA.shift_colors[v];
      const isToday = DATA.week_labels[i] === DATA.today_week_label;
      const style = color ? `background:${color}${v === 'Éjjel' ? ';color:white' : ''};` : '';
      html += `<td class="${isToday ? 'today-col' : ''}" style="${style}">${v}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody>';
  document.getElementById('weeklyTable').innerHTML = html;
}

function renderExceptions() {
  const search = document.getElementById('searchExc').value.trim().toLowerCase();
  const rows = DATA.exceptions.filter(e => e.year === currentYear && (!search || e.employee.toLowerCase().includes(search)));
  let html = '<thead><tr><th class="name-col">Munkatárs</th><th>Hét</th><th>Elvárt (rotáció)</th><th>Tényleges</th></tr></thead><tbody>';
  if (rows.length === 0) {
    html += '<tr><td colspan="4" style="text-align:center;color:#6b7280;">Nincs kivétel ebben az évben.</td></tr>';
  }
  rows.forEach(e => {
    const parts = e.week.split('.');
    const color = DATA.shift_colors[e.actual];
    const style = color ? `background:${color}${e.actual === 'Éjjel' ? ';color:white' : ''};` : '';
    html += `<tr><td class="name-col">${e.employee}</td><td>${parts[1]}.${parts[2]}</td><td>${e.expected}</td><td style="${style}">${e.actual}</td></tr>`;
  });
  html += '</tbody>';
  document.getElementById('excTable').innerHTML = html;
}

function renderSummary() {
  const container = document.getElementById('summaryCards');
  container.innerHTML = '';
  DATA.employees.forEach(name => {
    const yearData = (DATA.summary[name] || {})[currentYear] || {};
    const card = document.createElement('div');
    card.className = 'summary-card';
    let rowsHtml = '';
    DATA.summary_codes.forEach(code => {
      rowsHtml += `<tr><td>${code}</td><td class="num">${yearData[code] ?? 0}</td></tr>`;
    });
    const keret = yearData[DATA.quota_labels[0]] ?? 0;
    const hatra = yearData[DATA.quota_labels[1]] ?? 0;
    card.innerHTML = `<h3><span class="avatar">${initials(name)}</span>${name}</h3><table>${rowsHtml}</table>
      <div class="quota-box">
        <span class="quota-caption">Keret: ${keret} nap</span>
        <span class="quota-remaining">${hatra}<span class="quota-caption"> nap hátra</span></span>
      </div>`;
    container.appendChild(card);
  });
}

monthSelect.addEventListener('change', renderDaily);
document.getElementById('searchDaily').addEventListener('input', renderDaily);
document.getElementById('searchWeekly').addEventListener('input', renderWeekly);
document.getElementById('searchExc').addEventListener('input', renderExceptions);

renderDaily();
renderWeekly();
renderExceptions();
renderSummary();
</script>
</body>
</html>
"""

html = (HTML_TEMPLATE
        .replace("__GENERATED__", data["generated"])
        .replace("__TODAY_DISPLAY__", data["today_display"])
        .replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False)))
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print("OK:", OUT)
