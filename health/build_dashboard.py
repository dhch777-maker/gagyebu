"""건강검진 대시보드 빌더

data/structured/*.json 파일들을 읽어 trends 데이터를 추출하고,
Chart.js 기반 정적 HTML 대시보드를 생성한다.

사용법: python build_dashboard.py
출력:   dashboard.html (브라우저에서 바로 열기)
"""
import os
import json
import glob


def load_yearly_data(structured_dir):
    """연도별 JSON 파일들을 로드하여 연도순 정렬된 리스트로 반환"""
    pattern = os.path.join(structured_dir, "*.json")
    files = sorted(glob.glob(pattern))
    data = []
    for f in files:
        basename = os.path.splitext(os.path.basename(f))[0]
        # 연도 파일만 (2024.json, 2025.json 등)
        if not basename.isdigit():
            continue
        with open(f, "r", encoding="utf-8") as fh:
            data.append(json.load(fh))
    return sorted(data, key=lambda d: d["year"])


def safe_get(data, *keys):
    """중첩 딕셔너리에서 안전하게 value 추출"""
    obj = data
    for k in keys:
        if not isinstance(obj, dict):
            return None
        obj = obj.get(k)
    if isinstance(obj, dict):
        return obj.get("value")
    return obj


def extract_trends(yearly_data):
    """연도별 데이터에서 트렌드 추적용 구조 생성"""
    years = [d["year"] for d in yearly_data]
    meta = []
    for d in yearly_data:
        meta.append({
            "year": d["year"],
            "date": d.get("date"),
            "institution": d.get("institution"),
            "age": d.get("age"),
        })

    def series(label, unit, ref, *keys):
        values = [safe_get(d, *keys) for d in yearly_data]
        statuses = []
        for d in yearly_data:
            obj = d
            for k in keys[:-1]:
                obj = obj.get(k, {}) if isinstance(obj, dict) else {}
            item = obj.get(keys[-1], {}) if isinstance(obj, dict) else {}
            statuses.append(item.get("status") if isinstance(item, dict) else None)
        return {"label": label, "unit": unit, "ref": ref, "values": values, "statuses": statuses}

    trends = {
        "years": years,
        "meta": meta,
        "sections": {
            "basic": {
                "title": "신체계측",
                "metrics": [
                    series("체중", "kg", None, "basic", "weight"),
                    series("BMI", "kg/m²", "18.5-24.9", "basic", "bmi"),
                    series("허리둘레", "cm", "남<90", "basic", "waist"),
                    series("수축기 혈압", "mmHg", "<120", "basic", "bp_systolic"),
                    series("이완기 혈압", "mmHg", "<80", "basic", "bp_diastolic"),
                ],
            },
            "body_composition": {
                "title": "체성분 (InBody)",
                "metrics": [
                    series("체지방률", "%", "10-20(남)", "body_composition", "body_fat_pct"),
                    series("골격근량", "kg", None, "body_composition", "skeletal_muscle"),
                    series("내장지방", "레벨", "1-9", "body_composition", "visceral_fat_level"),
                    series("InBody 점수", "점", None, "body_composition", "inbody_score"),
                    series("기초대사량", "kcal", None, "body_composition", "bmr"),
                ],
            },
            "lipid": {
                "title": "지질 검사",
                "metrics": [
                    series("총콜레스테롤", "mg/dL", "<200", "blood", "lipid", "total_cholesterol"),
                    series("HDL", "mg/dL", "≥60", "blood", "lipid", "hdl"),
                    series("LDL", "mg/dL", "<130", "blood", "lipid", "ldl"),
                    series("중성지방", "mg/dL", "<150", "blood", "lipid", "triglyceride"),
                ],
            },
            "liver": {
                "title": "간기능",
                "metrics": [
                    series("AST (SGOT)", "U/L", "0-40", "blood", "liver", "sgot_ast"),
                    series("ALT (SGPT)", "U/L", "0-41", "blood", "liver", "sgpt_alt"),
                    series("GGT", "U/L", "11-63", "blood", "liver", "gamma_gtp"),
                    series("ALP", "U/L", "40-129", "blood", "liver", "alp"),
                    series("LDH", "U/L", "0-250", "blood", "liver", "ldh"),
                ],
            },
            "kidney": {
                "title": "신장기능",
                "metrics": [
                    series("BUN", "mg/dL", "6-20", "blood", "kidney", "bun"),
                    series("크레아티닌", "mg/dL", "0.6-1.2", "blood", "kidney", "creatinine"),
                    series("eGFR", "mL/min", "≥90", "blood", "kidney", "egfr"),
                    series("요산", "mg/dL", "3.4-7.0", "blood", "kidney", "uric_acid"),
                ],
            },
            "diabetes": {
                "title": "당뇨",
                "metrics": [
                    series("공복혈당", "mg/dL", "<100", "blood", "diabetes", "fasting_glucose"),
                    series("HbA1c", "%", "<5.7", "blood", "diabetes", "hba1c"),
                ],
            },
            "inflammation": {
                "title": "염증",
                "metrics": [
                    series("hs-CRP", "mg/L", "<3.0", "blood", "inflammation", "hs_crp"),
                ],
            },
            "cbc": {
                "title": "일반혈액 (CBC)",
                "metrics": [
                    series("WBC", "10³/μL", "4.0-10.0", "blood", "cbc", "wbc"),
                    series("헤모글로빈", "g/dL", "13.0-17.0", "blood", "cbc", "hemoglobin"),
                    series("헤마토크릿", "%", "39-52", "blood", "cbc", "hematocrit"),
                    series("혈소판", "10³/μL", "150-400", "blood", "cbc", "platelet"),
                ],
            },
        },
    }

    # 영상/내시경 소견 (텍스트 기반)
    imaging_data = []
    for d in yearly_data:
        year_imaging = {"year": d["year"], "items": []}
        for section_key in ["imaging", "endoscopy"]:
            section = d.get(section_key, {})
            for item_key, item_val in section.items():
                if isinstance(item_val, dict):
                    year_imaging["items"].append({
                        "name": item_key,
                        "status": item_val.get("status"),
                        "finding": item_val.get("finding"),
                    })
        imaging_data.append(year_imaging)
    trends["imaging"] = imaging_data

    # flags (이상항목)
    flags_data = []
    for d in yearly_data:
        flags_data.append({
            "year": d["year"],
            "flags": d.get("flags", []),
        })
    trends["flags"] = flags_data

    # 특수검사
    special_data = []
    for d in yearly_data:
        sp = d.get("special", {})
        special_data.append({
            "year": d["year"],
            "items": sp,
        })
    trends["special"] = special_data

    # 유전체
    genetics_data = []
    for d in yearly_data:
        gen = d.get("genetics")
        if gen:
            genetics_data.append({"year": d["year"], "data": gen})
    trends["genetics"] = genetics_data

    return trends


def build_html(trends):
    """트렌드 데이터를 포함한 대시보드 HTML 생성"""
    trends_json = json.dumps(trends, ensure_ascii=False, indent=2)
    return HTML_TEMPLATE.replace("__TRENDS_DATA__", trends_json)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>건강검진 대시보드</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
  :root {
    --bg: #0f1117;
    --surface: #1a1d27;
    --surface2: #232733;
    --border: #2e3344;
    --text: #e4e6ef;
    --text2: #8b8fa3;
    --accent: #6c8cff;
    --green: #34d399;
    --yellow: #fbbf24;
    --red: #f87171;
    --orange: #fb923c;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.6;
  }

  /* Header */
  .header {
    background: linear-gradient(135deg, #1e2235 0%, #141722 100%);
    border-bottom: 1px solid var(--border);
    padding: 2rem 2rem 1.5rem;
  }
  .header h1 {
    font-size: 1.75rem;
    font-weight: 700;
    margin-bottom: 0.5rem;
  }
  .header .subtitle { color: var(--text2); font-size: 0.9rem; }
  .meta-cards {
    display: flex;
    gap: 1rem;
    margin-top: 1rem;
    flex-wrap: wrap;
  }
  .meta-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.75rem 1.25rem;
    font-size: 0.85rem;
  }
  .meta-card .year { font-weight: 700; color: var(--accent); }

  /* Layout */
  .container { max-width: 1400px; margin: 0 auto; padding: 1.5rem; }
  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(580px, 1fr));
    gap: 1.5rem;
  }

  /* Section cards */
  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
  }
  .card h2 {
    font-size: 1.1rem;
    font-weight: 600;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
  }
  .card h2 .icon { font-size: 1.2rem; }
  .chart-wrap { position: relative; height: 280px; }

  /* Data table inside card */
  .data-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-top: 0.75rem;
  }
  .data-table th {
    text-align: left;
    padding: 0.5rem 0.75rem;
    border-bottom: 1px solid var(--border);
    color: var(--text2);
    font-weight: 500;
  }
  .data-table td {
    padding: 0.5rem 0.75rem;
    border-bottom: 1px solid rgba(46,51,68,0.5);
  }
  .data-table tr:last-child td { border-bottom: none; }

  /* Status badges */
  .badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 0.75rem;
    font-weight: 600;
  }
  .badge-normal { background: rgba(52,211,153,0.15); color: var(--green); }
  .badge-high   { background: rgba(248,113,113,0.15); color: var(--red); }
  .badge-low    { background: rgba(251,191,36,0.15); color: var(--yellow); }
  .badge-warn   { background: rgba(251,146,60,0.15); color: var(--orange); }

  /* Value display */
  .metric-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.4rem 0;
    border-bottom: 1px solid rgba(46,51,68,0.3);
  }
  .metric-row:last-child { border-bottom: none; }
  .metric-name { color: var(--text2); font-size: 0.85rem; }
  .metric-values { display: flex; gap: 1rem; font-size: 0.85rem; }
  .metric-val { min-width: 70px; text-align: right; }
  .metric-val.null { color: var(--text2); font-style: italic; }

  /* Trend arrows */
  .trend-up   { color: var(--red); }
  .trend-down { color: var(--green); }
  .trend-flat { color: var(--text2); }

  /* Full-width sections */
  .full-width { grid-column: 1 / -1; }

  /* Flags table */
  .flags-section { margin-top: 0.5rem; }
  .flag-year {
    font-weight: 600;
    color: var(--accent);
    padding: 0.5rem 0;
    border-bottom: 1px solid var(--border);
    margin-top: 0.75rem;
  }

  /* Responsive */
  @media (max-width: 640px) {
    .grid { grid-template-columns: 1fr; }
    .header { padding: 1.5rem 1rem 1rem; }
    .container { padding: 1rem; }
    .chart-wrap { height: 220px; }
  }

  /* Nav */
  .nav {
    position: sticky;
    top: 0;
    z-index: 100;
    background: rgba(15,17,23,0.95);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--border);
    padding: 0.5rem 1.5rem;
    display: flex;
    gap: 0.25rem;
    flex-wrap: wrap;
  }
  .nav a {
    color: var(--text2);
    text-decoration: none;
    font-size: 0.8rem;
    padding: 0.4rem 0.75rem;
    border-radius: 6px;
    transition: all 0.15s;
  }
  .nav a:hover {
    background: var(--surface2);
    color: var(--text);
  }
</style>
</head>
<body>

<script>
const TRENDS = __TRENDS_DATA__;
</script>

<div class="header">
  <h1>Health Dashboard</h1>
  <p class="subtitle" id="subtitle"></p>
  <div class="meta-cards" id="meta-cards"></div>
</div>

<nav class="nav" id="nav"></nav>

<div class="container">
  <div class="grid" id="grid"></div>
</div>

<script>
// ─── Helpers ────────────────────────────────────────────────
const SECTION_ICONS = {
  basic: '📏', body_composition: '🏋️', lipid: '🩸', liver: '🫁',
  kidney: '💧', diabetes: '🍬', inflammation: '🔥', cbc: '🔬',
};

function statusBadge(status) {
  if (!status) return '';
  const cls = {
    '정상': 'badge-normal', '높음': 'badge-high', '낮음': 'badge-low',
    '주의': 'badge-warn', '부족': 'badge-warn', '비대상': 'badge-normal',
  }[status] || 'badge-warn';
  return `<span class="badge ${cls}">${status}</span>`;
}

function fmtVal(v) {
  if (v === null || v === undefined) return '<span class="null">-</span>';
  return String(v);
}

function trendArrow(values) {
  const nums = values.filter(v => v !== null && v !== undefined);
  if (nums.length < 2) return '';
  const diff = nums[nums.length - 1] - nums[nums.length - 2];
  if (Math.abs(diff) < 0.01) return '<span class="trend-flat">→</span>';
  return diff > 0
    ? `<span class="trend-up">↑${Math.abs(diff).toFixed(1)}</span>`
    : `<span class="trend-down">↓${Math.abs(diff).toFixed(1)}</span>`;
}

const CHART_COLORS = [
  '#6c8cff', '#34d399', '#fbbf24', '#f87171', '#a78bfa',
  '#fb923c', '#38bdf8', '#f472b6',
];

function parseRef(ref) {
  if (!ref) return null;
  // "<200" or "≥60"
  let m = ref.match(/^[<≤]\s*([\d.]+)$/);
  if (m) return { max: parseFloat(m[1]) };
  m = ref.match(/^[>≥]\s*([\d.]+)$/);
  if (m) return { min: parseFloat(m[1]) };
  // "0.6-1.2" or "18.5-24.9"
  m = ref.match(/^([\d.]+)\s*[-~]\s*([\d.]+)$/);
  if (m) return { min: parseFloat(m[1]), max: parseFloat(m[2]) };
  return null;
}

// ─── Render meta ────────────────────────────────────────────
function renderMeta() {
  const sub = TRENDS.meta.map(m => `${m.year}년`).join(' → ') + ' 추이';
  document.getElementById('subtitle').textContent = sub;

  const cards = TRENDS.meta.map(m => `
    <div class="meta-card">
      <span class="year">${m.year}</span>
      ${m.date || ''} · ${m.institution || ''} · ${m.age || ''}세
    </div>
  `).join('');
  document.getElementById('meta-cards').innerHTML = cards;
}

// ─── Render chart sections ──────────────────────────────────
function createChart(canvas, metrics, years) {
  const datasets = metrics
    .filter(m => m.values.some(v => v !== null))
    .map((m, i) => ({
      label: `${m.label} (${m.unit})`,
      data: m.values,
      borderColor: CHART_COLORS[i % CHART_COLORS.length],
      backgroundColor: CHART_COLORS[i % CHART_COLORS.length] + '22',
      borderWidth: 2.5,
      pointRadius: 5,
      pointHoverRadius: 7,
      tension: 0.3,
      spanGaps: true,
    }));

  // Reference line annotations
  const annotations = [];
  metrics.forEach((m, i) => {
    const parsed = parseRef(m.ref);
    if (!parsed) return;
    const color = CHART_COLORS[i % CHART_COLORS.length];
    if (parsed.max !== undefined) {
      annotations.push({
        type: 'line', yMin: parsed.max, yMax: parsed.max,
        borderColor: color + '55', borderWidth: 1, borderDash: [4, 4],
        label: { display: false },
      });
    }
    if (parsed.min !== undefined) {
      annotations.push({
        type: 'line', yMin: parsed.min, yMax: parsed.min,
        borderColor: color + '55', borderWidth: 1, borderDash: [4, 4],
        label: { display: false },
      });
    }
  });

  new Chart(canvas, {
    type: 'line',
    data: {
      labels: years.map(String),
      datasets,
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          labels: { color: '#8b8fa3', font: { size: 11 }, usePointStyle: true, pointStyle: 'circle' },
        },
        tooltip: {
          backgroundColor: '#232733',
          titleColor: '#e4e6ef',
          bodyColor: '#e4e6ef',
          borderColor: '#2e3344',
          borderWidth: 1,
          padding: 10,
        },
        annotation: annotations.length > 0 ? { annotations } : undefined,
      },
      scales: {
        x: {
          ticks: { color: '#8b8fa3' },
          grid: { color: 'rgba(46,51,68,0.5)' },
        },
        y: {
          ticks: { color: '#8b8fa3' },
          grid: { color: 'rgba(46,51,68,0.5)' },
        },
      },
    },
  });
}

function renderChartSections() {
  const grid = document.getElementById('grid');
  const nav = document.getElementById('nav');
  const sections = TRENDS.sections;

  for (const [key, section] of Object.entries(sections)) {
    const hasData = section.metrics.some(m => m.values.some(v => v !== null));
    if (!hasData) continue;

    const id = `section-${key}`;
    const icon = SECTION_ICONS[key] || '📊';

    // Nav link
    nav.innerHTML += `<a href="#${id}">${icon} ${section.title}</a>`;

    // Metric table rows
    const tableRows = section.metrics.map(m => {
      const cells = TRENDS.years.map((y, i) => {
        const v = fmtVal(m.values[i]);
        const badge = statusBadge(m.statuses[i]);
        return `<td>${v} ${badge}</td>`;
      }).join('');
      return `<tr>
        <td>${m.label} <span style="color:var(--text2);font-size:0.75rem">${m.unit}</span></td>
        ${cells}
        <td>${trendArrow(m.values)}</td>
        <td style="color:var(--text2);font-size:0.75rem">${m.ref || '-'}</td>
      </tr>`;
    }).join('');

    const yearHeaders = TRENDS.years.map(y => `<th>${y}</th>`).join('');

    const card = document.createElement('div');
    card.className = 'card';
    card.id = id;
    card.innerHTML = `
      <h2><span class="icon">${icon}</span>${section.title}</h2>
      <div class="chart-wrap"><canvas id="chart-${key}"></canvas></div>
      <table class="data-table">
        <thead><tr><th>항목</th>${yearHeaders}<th>변동</th><th>참고치</th></tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
    `;
    grid.appendChild(card);

    const canvas = card.querySelector(`#chart-${key}`);
    createChart(canvas, section.metrics, TRENDS.years);
  }
}

// ─── Imaging / Endoscopy table ──────────────────────────────
function renderImaging() {
  const grid = document.getElementById('grid');
  const nav = document.getElementById('nav');

  nav.innerHTML += `<a href="#section-imaging">🏥 영상/내시경</a>`;

  const ITEM_LABELS = {
    chest_xray: '흉부 X선', chest_ct: '흉부 CT', abdominal_us: '복부 초음파',
    thyroid_us: '갑상선 초음파', prostate_us: '전립선 초음파',
    gastroscopy: '위내시경', colonoscopy: '대장내시경',
  };

  let rows = '';
  const allItems = new Set();
  TRENDS.imaging.forEach(y => y.items.forEach(i => allItems.add(i.name)));

  allItems.forEach(name => {
    const label = ITEM_LABELS[name] || name;
    const cells = TRENDS.imaging.map(y => {
      const item = y.items.find(i => i.name === name);
      if (!item) return '<td style="color:var(--text2)">-</td>';
      return `<td>${statusBadge(item.status)} <span style="font-size:0.8rem">${item.finding || ''}</span></td>`;
    }).join('');
    rows += `<tr><td>${label}</td>${cells}</tr>`;
  });

  const yearHeaders = TRENDS.years.map(y => `<th>${y}</th>`).join('');
  const card = document.createElement('div');
  card.className = 'card full-width';
  card.id = 'section-imaging';
  card.innerHTML = `
    <h2><span class="icon">🏥</span>영상/내시경 소견</h2>
    <table class="data-table">
      <thead><tr><th>검사</th>${yearHeaders}</tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
  grid.appendChild(card);
}

// ─── Flags (이상항목) ───────────────────────────────────────
function renderFlags() {
  const grid = document.getElementById('grid');
  const nav = document.getElementById('nav');

  nav.innerHTML += `<a href="#section-flags">⚠️ 이상항목</a>`;

  const CATEGORY_LABELS = {
    blood: '혈액', imaging: '영상', endoscopy: '내시경',
    lifestyle: '생활습관', genetics: '유전체',
  };
  const ITEM_LABELS = {
    total_cholesterol: '총콜레스테롤', ldl: 'LDL', hdl: 'HDL',
    sgpt_alt: 'ALT', sgot_ast: 'AST', hs_crp: 'hs-CRP',
    neutrophil: '호중구', fasting_glucose: '공복혈당',
    abdominal_us: '복부초음파', gastroscopy: '위내시경',
    exercise: '운동', myocardial_infarction: '심근경색 위험',
    atrial_fibrillation: '심방세동 위험',
  };

  let content = '';
  TRENDS.flags.forEach(yf => {
    content += `<div class="flag-year">${yf.year}년 (${yf.flags.length}건)</div>`;
    if (yf.flags.length === 0) {
      content += '<p style="color:var(--text2);padding:0.5rem 0;font-size:0.85rem">이상항목 없음</p>';
      return;
    }
    content += '<table class="data-table"><thead><tr><th>분류</th><th>항목</th><th>수치</th><th>상태</th><th>조치</th></tr></thead><tbody>';
    yf.flags.forEach(f => {
      const cat = CATEGORY_LABELS[f.category] || f.category;
      const item = ITEM_LABELS[f.item] || f.item;
      content += `<tr>
        <td>${cat}</td>
        <td>${item}</td>
        <td>${fmtVal(f.value)}</td>
        <td>${statusBadge(f.status)}</td>
        <td style="font-size:0.8rem;color:var(--text2)">${f.action || ''}</td>
      </tr>`;
    });
    content += '</tbody></table>';
  });

  const card = document.createElement('div');
  card.className = 'card full-width';
  card.id = 'section-flags';
  card.innerHTML = `
    <h2><span class="icon">⚠️</span>이상항목 추이</h2>
    <div class="flags-section">${content}</div>
  `;
  grid.appendChild(card);
}

// ─── Special tests ──────────────────────────────────────────
function renderSpecial() {
  const grid = document.getElementById('grid');
  const nav = document.getElementById('nav');

  nav.innerHTML += `<a href="#section-special">🧪 특수검사</a>`;

  const LABELS = {
    ecg: '심전도', bone_density: '골밀도',
    arterial_stiffness: '동맥경화도', nk_cell_activity: 'NK세포활성도',
  };

  let rows = '';
  const allKeys = new Set();
  TRENDS.special.forEach(s => {
    Object.keys(s.items).forEach(k => {
      if (['cardiovascular_risk', 'metabolic_age', 'health_type', 'foot_analysis'].includes(k)) return;
      allKeys.add(k);
    });
  });

  allKeys.forEach(key => {
    const label = LABELS[key] || key;
    const cells = TRENDS.special.map(s => {
      const item = s.items[key];
      if (!item) return '<td style="color:var(--text2)">-</td>';
      const val = item.value !== undefined ? item.value : '';
      return `<td>${statusBadge(item.status)} ${val} <span style="font-size:0.78rem;color:var(--text2)">${item.finding || ''}</span></td>`;
    }).join('');
    rows += `<tr><td>${label}</td>${cells}</tr>`;
  });

  const yearHeaders = TRENDS.years.map(y => `<th>${y}</th>`).join('');
  const card = document.createElement('div');
  card.className = 'card full-width';
  card.id = 'section-special';
  card.innerHTML = `
    <h2><span class="icon">🧪</span>특수검사</h2>
    <table class="data-table">
      <thead><tr><th>검사</th>${yearHeaders}</tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
  grid.appendChild(card);
}

// ─── Genetics ───────────────────────────────────────────────
function renderGenetics() {
  if (!TRENDS.genetics || TRENDS.genetics.length === 0) return;

  const grid = document.getElementById('grid');
  const nav = document.getElementById('nav');
  nav.innerHTML += `<a href="#section-genetics">🧬 유전체</a>`;

  const ITEM_LABELS = {
    myocardial_infarction: '심근경색', atrial_fibrillation: '심방세동',
    parkinsons: '파킨슨병', stroke: '뇌졸중', alzheimers: '알츠하이머',
    colorectal: '대장암', lung: '폐암', prostate: '전립선암',
    liver: '간암', pancreatic: '췌장암',
  };

  let content = '';
  TRENDS.genetics.forEach(g => {
    for (const [catKey, cat] of Object.entries(g.data)) {
      const title = catKey === 'cardiovascular_5' ? '심뇌혈관 5종' : catKey === 'cancer_5_male' ? '남성 암 5종' : catKey;
      content += `<div class="flag-year">${g.year}년 · ${title} (${cat.provider})</div>`;
      content += '<table class="data-table"><thead><tr><th>항목</th><th>위험배율</th><th>수준</th></tr></thead><tbody>';
      for (const [itemKey, item] of Object.entries(cat.items)) {
        const label = ITEM_LABELS[itemKey] || itemKey;
        const levelClass = item.risk_ratio > 1.5 ? 'badge-high' : item.risk_ratio > 1.0 ? 'badge-warn' : 'badge-normal';
        content += `<tr>
          <td>${label}</td>
          <td>${item.risk_ratio}x</td>
          <td><span class="badge ${levelClass}">${item.level}</span></td>
        </tr>`;
      }
      content += '</tbody></table>';
    }
  });

  const card = document.createElement('div');
  card.className = 'card full-width';
  card.id = 'section-genetics';
  card.innerHTML = `
    <h2><span class="icon">🧬</span>유전체 검사</h2>
    ${content}
  `;
  grid.appendChild(card);
}

// ─── Init ───────────────────────────────────────────────────
renderMeta();
renderChartSections();
renderImaging();
renderFlags();
renderSpecial();
renderGenetics();
</script>

</body>
</html>
"""


def main():
    health_dir = os.path.dirname(os.path.abspath(__file__))
    structured_dir = os.path.join(health_dir, "data", "structured")
    output_path = os.path.join(health_dir, "dashboard.html")

    yearly_data = load_yearly_data(structured_dir)
    if not yearly_data:
        print("data/structured/ 에 연도별 JSON 파일이 없습니다.")
        return

    print(f"로드된 연도: {[d['year'] for d in yearly_data]}")

    trends = extract_trends(yearly_data)
    html = build_html(trends)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"대시보드 생성 완료: {output_path}")


if __name__ == "__main__":
    main()
