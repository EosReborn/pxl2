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

DAILY_CODES = {
    "":   {"label": "Munkanap", "color": "#ffffff"},
    "SZ": {"label": "Szabadság", "color": "#ffc000"},
    "B":  {"label": "Betegszabadság", "color": "#ff6b6b"},
    "OH": {"label": "Home office", "color": "#9dc3e6"},
    "UN": {"label": "Ünnepnap", "color": "#c39bd3"},
}
SHIFT_COLORS = {
    "Éjjel": "#5b4b8a",
    "Délután": "#f4b183",
    "Délelőtt": "#ffe699",
    "Szabadság": "#ffc000",
    "Home office": "#9dc3e6",
    "Egyéb": "#d9d9d9",
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
    while ws.cell(row=r, column=1).value:
        name = ws.cell(row=r, column=1).value
        employees.append(name)
        row_codes = []
        for c in range(2, max_col + 1):
            v = ws.cell(row=r, column=c).value
            row_codes.append((v or "").strip().upper())
        daily[name] = row_codes
        r += 1
    daily_by_year[year] = {"dates": dates, "employees": employees, "daily": daily}
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
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="hu">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Csapat naptár</title>
<style>
  :root {
    --border: #e2e2e2;
    --text: #23272f;
    --muted: #6b7280;
    --accent: #2f5496;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, "Segoe UI", Arial, sans-serif;
    margin: 0;
    background: #f7f8fa;
    color: var(--text);
  }
  header {
    background: var(--accent);
    color: white;
    padding: 20px 28px;
  }
  header h1 { margin: 0 0 4px 0; font-size: 22px; }
  header p { margin: 0; opacity: 0.85; font-size: 13px; }
  .wrap { max-width: 1400px; margin: 0 auto; padding: 20px; }
  .tabs { display: flex; gap: 8px; margin-bottom: 16px; }
  .tab-btn {
    padding: 8px 16px; border-radius: 8px; border: 1px solid var(--border);
    background: white; cursor: pointer; font-size: 14px; font-weight: 600;
    color: var(--muted);
  }
  .tab-btn.active { background: var(--accent); color: white; border-color: var(--accent); }
  .panel { display: none; background: white; border-radius: 10px; border: 1px solid var(--border); padding: 18px; }
  .panel.active { display: block; }
  .controls { display: flex; gap: 12px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }
  select, input[type=text] {
    padding: 7px 10px; border-radius: 6px; border: 1px solid var(--border); font-size: 14px;
  }
  table { border-collapse: collapse; font-size: 12px; }
  th, td {
    border: 1px solid var(--border); padding: 3px 4px; text-align: center; white-space: nowrap;
  }
  th.name-col, td.name-col {
    position: sticky; left: 0; background: white; text-align: left; z-index: 2;
    min-width: 130px; font-weight: 600; font-size: 13px;
  }
  thead th { position: sticky; top: 0; background: var(--accent); color: white; z-index: 3; font-weight: 500; }
  thead th.name-col { z-index: 4; }
  .table-scroll { overflow: auto; max-height: 65vh; border: 1px solid var(--border); border-radius: 8px; }
  .weekend { background: #f1f1f1; }
  .legend { display: flex; gap: 16px; flex-wrap: wrap; margin-top: 16px; font-size: 13px; }
  .legend span.swatch { display: inline-block; width: 14px; height: 14px; border-radius: 3px; margin-right: 6px; vertical-align: middle; border: 1px solid #d0d0d0; }
  footer { text-align: center; color: var(--muted); font-size: 12px; padding: 16px 24px 28px; }
  td.code-SZ { background: #ffc000; }
  td.code-B { background: #ff6b6b; color: white; }
  td.code-OH { background: #9dc3e6; }
  td.code-UN { background: #c39bd3; }
  .year-bar {
    display: flex; justify-content: center; gap: 8px; margin: 28px auto 8px; flex-wrap: wrap;
  }
  .year-btn {
    padding: 8px 20px; border-radius: 20px; border: 1px solid var(--border);
    background: white; cursor: pointer; font-size: 14px; font-weight: 600; color: var(--muted);
  }
  .year-btn.active { background: var(--accent); color: white; border-color: var(--accent); }
</style>
</head>
<body>
<header>
  <h1>Csapat naptár — szabadság, betegség, munkarend</h1>
  <p>Utolsó frissítés: __GENERATED__ &nbsp;·&nbsp; csak megtekintésre</p>
</header>
<div class="wrap">
  <div class="tabs">
    <button class="tab-btn active" data-tab="daily">Napi jelenlét</button>
    <button class="tab-btn" data-tab="weekly">Heti beosztás (műszakrend)</button>
  </div>

  <div id="daily" class="panel active">
    <div class="controls">
      <label>Hónap:</label>
      <select id="monthSelect"></select>
      <label>Keresés:</label>
      <input type="text" id="searchDaily" placeholder="Munkatárs neve...">
    </div>
    <div class="table-scroll"><table id="dailyTable"></table></div>
    <div class="legend" id="legend"></div>
  </div>

  <div id="weekly" class="panel">
    <div class="controls">
      <label>Keresés:</label>
      <input type="text" id="searchWeekly" placeholder="Munkatárs neve...">
    </div>
    <div class="table-scroll"><table id="weeklyTable"></table></div>
    <div class="legend" id="legendWeekly"></div>
  </div>

  <div class="year-bar" id="yearBar"></div>
</div>
<footer>Generálva a csapat_nyilvantartas.xlsx fájlból &middot; a szerkesztést a csoportvezető végzi</footer>

<script>
const DATA = __DATA_JSON__;
let currentYear = DATA.years.includes(2026) ? 2026 : DATA.years[0];

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
});

// --- year bar (alul, mindkét nézetet vezérli) ---
const yearBar = document.getElementById('yearBar');
DATA.years.forEach(y => {
  const btn = document.createElement('button');
  btn.className = 'year-btn' + (y === currentYear ? ' active' : '');
  btn.textContent = y;
  btn.dataset.year = y;
  btn.addEventListener('click', () => {
    currentYear = y;
    document.querySelectorAll('.year-btn').forEach(b => b.classList.toggle('active', parseInt(b.dataset.year) === y));
    renderDaily();
    renderWeekly();
  });
  yearBar.appendChild(btn);
});

// --- legend (napi) ---
const legendEl = document.getElementById('legend');
Object.entries(DATA.codes).forEach(([code, info]) => {
  if (!code) return;
  const el = document.createElement('span');
  el.innerHTML = `<span class="swatch" style="background:${info.color}"></span>${info.label} (${code})`;
  legendEl.appendChild(el);
});

// --- legend (heti / műszak) ---
const legendWeeklyEl = document.getElementById('legendWeekly');
Object.entries(DATA.shift_colors).forEach(([label, color]) => {
  const el = document.createElement('span');
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
monthSelect.value = 6; // alapértelmezett: július

function monthOfIndex(dates, i) {
  const label = dates[i];
  const prefix = label.split('.')[0];
  return huMonthAbbr.indexOf(prefix);
}

function renderDaily() {
  const yearData = DATA.daily_by_year[currentYear];
  if (!yearData) return;
  const monthIdx = parseInt(monthSelect.value, 10);
  const search = document.getElementById('searchDaily').value.trim().toLowerCase();
  const colIndexes = [];
  yearData.dates.forEach((d, i) => { if (monthOfIndex(yearData.dates, i) === monthIdx) colIndexes.push(i); });

  let html = '<thead><tr><th class="name-col">Munkatárs</th>';
  colIndexes.forEach(i => {
    const label = yearData.dates[i].split('.')[1];
    html += `<th>${label}</th>`;
  });
  html += '</tr></thead><tbody>';

  yearData.employees.forEach(name => {
    if (search && !name.toLowerCase().includes(search)) return;
    html += `<tr><td class="name-col">${name}</td>`;
    colIndexes.forEach(i => {
      const code = (yearData.daily[name] && yearData.daily[name][i]) || '';
      const cls = code ? `code-${code}` : '';
      html += `<td class="${cls}" title="${DATA.codes[code] ? DATA.codes[code].label : ''}">${code}</td>`;
    });
    html += '</tr>';
  });
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
    html += `<th>${parts[1]}.${parts[2]}</th>`;
  });
  html += '</tr></thead><tbody>';
  DATA.employees.forEach(name => {
    if (search && !name.toLowerCase().includes(search)) return;
    const vals = DATA.weekly[name] || [];
    html += `<tr><td class="name-col">${name}</td>`;
    colIndexes.forEach(i => {
      const v = vals[i] || '';
      const color = DATA.shift_colors[v];
      const style = color ? `style="background:${color}${v === 'Éjjel' ? ';color:white' : ''}"` : '';
      html += `<td ${style}>${v}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody>';
  document.getElementById('weeklyTable').innerHTML = html;
}

monthSelect.addEventListener('change', renderDaily);
document.getElementById('searchDaily').addEventListener('input', renderDaily);
document.getElementById('searchWeekly').addEventListener('input', renderWeekly);

renderDaily();
renderWeekly();
</script>
</body>
</html>
"""

html = HTML_TEMPLATE.replace("__GENERATED__", data["generated"]).replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print("OK:", OUT)
