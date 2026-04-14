const XLSX = require('../node_modules/xlsx');
const ExcelJS = require('../node_modules/exceljs');
const { ChartJSNodeCanvas } = require('../node_modules/chartjs-node-canvas');
const path = require('path');

const SRC  = path.join(__dirname, '2025_연간매출_통합.xlsx');
const DEST = path.join(__dirname, '2025_연간매출_대시보드.xlsx');

// ── 데이터 로드 ────────────────────────────────────────────
const rawWb = XLSX.readFile(SRC);
const rows  = XLSX.utils.sheet_to_json(rawWb.Sheets[rawWb.SheetNames[0]]);

// ── 집계 헬퍼 ─────────────────────────────────────────────
function aggregate(rows, key) {
  const map = {};
  rows.forEach(r => {
    const k = r[key];
    if (!map[k]) map[k] = { 매출액: 0, 비용: 0, 이익: 0, 건수: 0 };
    map[k].매출액 += r['매출액'] || 0;
    map[k].비용   += r['비용']   || 0;
    map[k].이익   += r['이익']   || 0;
    map[k].건수   += 1;
  });
  return map;
}

const byMonth  = aggregate(rows, '월');
const byDept   = aggregate(rows, '부서');
const byMgr    = aggregate(rows, '담당자');
const byItem   = aggregate(rows, '항목');
const byClient = aggregate(rows, '거래처');

const months = Object.keys(byMonth).sort();

// 전체 합계
const total = rows.reduce((s, r) => {
  s.매출액 += r['매출액'] || 0;
  s.비용   += r['비용']   || 0;
  s.이익   += r['이익']   || 0;
  return s;
}, { 매출액: 0, 비용: 0, 이익: 0 });

// ── 차트 생성 함수 ─────────────────────────────────────────
const W = 900, H = 480;
function makeCanvas() { return new ChartJSNodeCanvas({ width: W, height: H, backgroundColour: 'white' }); }

const NAVY   = '#1F4E79';
const BLUE   = '#2E75B6';
const LBLUE  = '#9DC3E6';
const GREEN  = '#375623';
const LGREEN = '#70AD47';
const RED    = '#C00000';
const GRAY   = '#595959';

async function chartMonthlyBar() {
  const canvas = makeCanvas();
  return canvas.renderToBuffer({
    type: 'bar',
    data: {
      labels: months.map(m => m.replace('2025-', '') + '월'),
      datasets: [
        {
          label: '매출액',
          data: months.map(m => Math.round(byMonth[m].매출액 / 1e8) / 10),
          backgroundColor: BLUE,
          order: 1,
        },
        {
          label: '비용',
          data: months.map(m => Math.round(byMonth[m].비용 / 1e8) / 10),
          backgroundColor: LBLUE,
          order: 1,
        },
        {
          label: '이익',
          data: months.map(m => Math.round(byMonth[m].이익 / 1e8) / 10),
          backgroundColor: LGREEN,
          order: 1,
        },
      ],
    },
    options: {
      plugins: {
        title: { display: true, text: '월별 매출 / 비용 / 이익 (단위: 억원)', font: { size: 16, weight: 'bold' }, color: NAVY },
        legend: { position: 'bottom' },
      },
      scales: {
        y: { ticks: { callback: v => v + '억' } },
      },
    },
  });
}

async function chartMonthlyLine() {
  const canvas = makeCanvas();
  return canvas.renderToBuffer({
    type: 'line',
    data: {
      labels: months.map(m => m.replace('2025-', '') + '월'),
      datasets: [
        {
          label: '이익률 (%)',
          data: months.map(m => parseFloat((byMonth[m].이익 / byMonth[m].매출액 * 100).toFixed(1))),
          borderColor: RED,
          backgroundColor: 'rgba(192,0,0,0.1)',
          borderWidth: 3,
          pointRadius: 5,
          pointBackgroundColor: RED,
          fill: true,
          tension: 0.3,
          yAxisID: 'y',
        },
        {
          label: '매출액 (억원)',
          data: months.map(m => Math.round(byMonth[m].매출액 / 1e8) / 10),
          borderColor: BLUE,
          backgroundColor: 'transparent',
          borderWidth: 2,
          pointRadius: 4,
          pointBackgroundColor: BLUE,
          tension: 0.3,
          yAxisID: 'y2',
        },
      ],
    },
    options: {
      plugins: {
        title: { display: true, text: '월별 이익률 추이 & 매출액', font: { size: 16, weight: 'bold' }, color: NAVY },
        legend: { position: 'bottom' },
      },
      scales: {
        y:  { position: 'left',  ticks: { callback: v => v + '%' }, min: 30, max: 42 },
        y2: { position: 'right', ticks: { callback: v => v + '억' }, grid: { drawOnChartArea: false } },
      },
    },
  });
}

async function chartDeptHorizontal() {
  const depts = Object.entries(byDept).sort((a, b) => b[1].매출액 - a[1].매출액);
  const canvas = makeCanvas();
  return canvas.renderToBuffer({
    type: 'bar',
    data: {
      labels: depts.map(([d]) => d),
      datasets: [
        {
          label: '매출액 (억원)',
          data: depts.map(([, v]) => Math.round(v.매출액 / 1e8) / 10),
          backgroundColor: [NAVY, BLUE, LBLUE, '#5B9BD5', '#BDD7EE'],
        },
        {
          label: '이익 (억원)',
          data: depts.map(([, v]) => Math.round(v.이익 / 1e8) / 10),
          backgroundColor: [GREEN, LGREEN, '#A9D18E', '#E2EFDA', '#C6E0B4'],
        },
      ],
    },
    options: {
      indexAxis: 'y',
      plugins: {
        title: { display: true, text: '부서별 매출액 vs 이익 (단위: 억원)', font: { size: 16, weight: 'bold' }, color: NAVY },
        legend: { position: 'bottom' },
      },
    },
  });
}

async function chartMgrBar() {
  const mgrs = Object.entries(byMgr).sort((a, b) => b[1].매출액 - a[1].매출액);
  const canvas = new ChartJSNodeCanvas({ width: 900, height: 520, backgroundColour: 'white' });
  return canvas.renderToBuffer({
    type: 'bar',
    data: {
      labels: mgrs.map(([m]) => m),
      datasets: [
        {
          label: '매출액 (억원)',
          data: mgrs.map(([, v]) => Math.round(v.매출액 / 1e8) / 10),
          backgroundColor: BLUE,
          yAxisID: 'y',
        },
        {
          label: '이익률 (%)',
          data: mgrs.map(([, v]) => parseFloat((v.이익 / v.매출액 * 100).toFixed(1))),
          type: 'line',
          borderColor: RED,
          backgroundColor: 'transparent',
          borderWidth: 2,
          pointRadius: 5,
          pointBackgroundColor: RED,
          yAxisID: 'y2',
        },
      ],
    },
    options: {
      plugins: {
        title: { display: true, text: '담당자별 매출액 & 이익률', font: { size: 16, weight: 'bold' }, color: NAVY },
        legend: { position: 'bottom' },
      },
      scales: {
        y:  { position: 'left',  ticks: { callback: v => v + '억' } },
        y2: { position: 'right', ticks: { callback: v => v + '%' }, min: 30, max: 42, grid: { drawOnChartArea: false } },
      },
    },
  });
}

async function chartItemPie() {
  const items = Object.entries(byItem).sort((a, b) => b[1].이익 - a[1].이익).slice(0, 10);
  const canvas = makeCanvas();
  const COLORS = [NAVY, BLUE, LBLUE, '#5B9BD5', '#BDD7EE', LGREEN, '#A9D18E', '#E2EFDA', RED, '#FF6666'];
  return canvas.renderToBuffer({
    type: 'doughnut',
    data: {
      labels: items.map(([it]) => it),
      datasets: [{
        data: items.map(([, v]) => Math.round(v.이익 / 1e4)),
        backgroundColor: COLORS,
        borderWidth: 2,
        borderColor: '#ffffff',
      }],
    },
    options: {
      plugins: {
        title: { display: true, text: '항목별 이익 비중 TOP10 (단위: 만원)', font: { size: 16, weight: 'bold' }, color: NAVY },
        legend: { position: 'right' },
      },
    },
  });
}

async function chartClientTop10() {
  const clients = Object.entries(byClient).sort((a, b) => b[1].매출액 - a[1].매출액).slice(0, 10);
  const canvas = new ChartJSNodeCanvas({ width: 900, height: 520, backgroundColour: 'white' });
  return canvas.renderToBuffer({
    type: 'bar',
    data: {
      labels: clients.map(([c]) => c),
      datasets: [
        {
          label: '매출액 (억원)',
          data: clients.map(([, v]) => Math.round(v.매출액 / 1e8) / 10),
          backgroundColor: clients.map((_, i) => i === 0 ? NAVY : BLUE),
          yAxisID: 'y',
        },
        {
          label: '이익률 (%)',
          data: clients.map(([, v]) => parseFloat((v.이익 / v.매출액 * 100).toFixed(1))),
          type: 'line',
          borderColor: RED,
          backgroundColor: 'transparent',
          borderWidth: 2,
          pointRadius: 5,
          pointBackgroundColor: RED,
          yAxisID: 'y2',
        },
      ],
    },
    options: {
      plugins: {
        title: { display: true, text: '거래처 매출 TOP10 & 이익률', font: { size: 16, weight: 'bold' }, color: NAVY },
        legend: { position: 'bottom' },
      },
      scales: {
        y:  { position: 'left',  ticks: { callback: v => v + '억' } },
        y2: { position: 'right', ticks: { callback: v => v + '%' }, min: 28, max: 42, grid: { drawOnChartArea: false } },
      },
    },
  });
}

// ── ExcelJS 유틸 ────────────────────────────────────────────
function headerStyle(cell, bgArgb = 'FF1F4E79') {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: '맑은 고딕' };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF2E75B6' } },
    bottom: { style: 'thin', color: { argb: 'FF2E75B6' } },
    left:   { style: 'thin', color: { argb: 'FF2E75B6' } },
    right:  { style: 'thin', color: { argb: 'FF2E75B6' } },
  };
}

function dataCell(cell, isEven, numFmt) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFD6E4F0' : 'FFFFFFFF' } };
  cell.font = { size: 10, name: '맑은 고딕' };
  cell.border = {
    top: { style: 'hair', color: { argb: 'FFBDD7EE' } },
    bottom: { style: 'hair', color: { argb: 'FFBDD7EE' } },
    left:   { style: 'hair', color: { argb: 'FFBDD7EE' } },
    right:  { style: 'hair', color: { argb: 'FFBDD7EE' } },
  };
  if (numFmt) {
    cell.numFmt = numFmt;
    cell.alignment = { horizontal: 'right', vertical: 'middle' };
  } else {
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  }
}

function kpiBox(sheet, startRow, startCol, label, value, subLabel, bgArgb) {
  // 배경 병합 셀 (3행 x 3열)
  const r1 = startRow, r2 = startRow + 2;
  const c1 = startCol, c2 = startCol + 2;
  sheet.mergeCells(r1, c1, r2, c2);
  const cell = sheet.getCell(r1, c1);
  cell.value = `${label}\n${value}\n${subLabel}`;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: '맑은 고딕' };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
}

// ── 메인 ───────────────────────────────────────────────────
async function main() {
  console.log('차트 이미지 생성 중...');
  const [imgMonthBar, imgMonthLine, imgDept, imgMgr, imgItemPie, imgClient] = await Promise.all([
    chartMonthlyBar(),
    chartMonthlyLine(),
    chartDeptHorizontal(),
    chartMgrBar(),
    chartItemPie(),
    chartClientTop10(),
  ]);
  console.log('✓ 차트 6종 생성 완료');

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Claude';
  workbook.created = new Date();

  // ═══════════════════════════════════════════════════════════
  // 시트 1: 대시보드 (KPI + 차트)
  // ═══════════════════════════════════════════════════════════
  const dash = workbook.addWorksheet('📊 대시보드');
  dash.views = [{ showGridLines: false }];

  // 컬럼 너비
  for (let c = 1; c <= 16; c++) dash.getColumn(c).width = 9.5;

  // 제목
  dash.mergeCells('A1:P1');
  const titleCell = dash.getCell('A1');
  titleCell.value = '2025년 연간 매출 대시보드';
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 18, name: '맑은 고딕' };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  dash.getRow(1).height = 36;

  // 부제목
  dash.mergeCells('A2:P2');
  const subCell = dash.getCell('A2');
  subCell.value = `분석 기준: 2025년 1월 ~ 12월  |  총 거래 건수: ${rows.length}건  |  담당자: ${Object.keys(byMgr).length}명  |  거래처: ${Object.keys(byClient).length}개`;
  subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
  subCell.font = { color: { argb: 'FFFFFFFF' }, size: 10, name: '맑은 고딕' };
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };
  dash.getRow(2).height = 20;

  // ── KPI 박스 (4개) ──
  dash.getRow(3).height = 8;
  for (let r = 4; r <= 6; r++) dash.getRow(r).height = 26;
  dash.getRow(7).height = 8;

  const profitRate = (total.이익 / total.매출액 * 100).toFixed(1);
  const bestMonth  = months.reduce((best, m) => byMonth[m].매출액 > byMonth[best].매출액 ? m : best, months[0]);
  const bestDept   = Object.entries(byDept).sort((a,b) => b[1].이익/b[1].매출액 - a[1].이익/a[1].매출액)[0];

  const kpis = [
    { label: '총 매출액', value: (total.매출액/1e8).toFixed(1) + '억원', sub: 'Annual Revenue', bg: 'FF1F4E79' },
    { label: '총 이익',   value: (total.이익/1e8).toFixed(1) + '억원',   sub: 'Annual Profit',  bg: 'FF2E75B6' },
    { label: '평균 이익률', value: profitRate + '%',                      sub: 'Profit Margin',  bg: 'FF375623' },
    { label: '최고 매출 월', value: bestMonth.replace('2025-','') + '월', sub: (byMonth[bestMonth].매출액/1e8).toFixed(1)+'억', bg: 'FFC00000' },
  ];

  // KPI를 A4:P6 범위에 4개 배치
  const kpiCols = [1, 5, 9, 13]; // A, E, I, M
  kpis.forEach((kpi, i) => {
    const col = kpiCols[i];
    dash.mergeCells(4, col, 6, col + 3);
    const cell = dash.getCell(4, col);
    cell.value = `${kpi.label}\n${kpi.value}\n${kpi.sub}`;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: '맑은 고딕' };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // ── 차트 삽입 ──
  const addChart = async (sheet, imgBuf, cell, w, h) => {
    const imgId = workbook.addImage({ buffer: imgBuf, extension: 'png' });
    sheet.addImage(imgId, {
      tl: { col: cell.col - 1, row: cell.row - 1 },
      ext: { width: w, height: h },
    });
  };

  // 행 높이 설정 (차트 영역)
  for (let r = 8; r <= 40; r++) dash.getRow(r).height = 16;

  // 차트1: 월별 막대 (A8, 좌상)
  await addChart(dash, imgMonthBar, { col: 1, row: 8 }, 750, 360);
  // 차트2: 이익률 라인 (A28, 좌하)
  await addChart(dash, imgMonthLine, { col: 1, row: 30 }, 750, 360);

  console.log('✓ 대시보드 시트 완료');

  // ═══════════════════════════════════════════════════════════
  // 시트 2: 월별 분석
  // ═══════════════════════════════════════════════════════════
  const sheetMonth = workbook.addWorksheet('📅 월별 분석');
  sheetMonth.views = [{ showGridLines: false }];

  // 제목
  sheetMonth.mergeCells('A1:H1');
  Object.assign(sheetMonth.getCell('A1'), {
    value: '월별 매출 분석',
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } },
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 14, name: '맑은 고딕' },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  sheetMonth.getRow(1).height = 28;

  // 헤더
  const mHeaders = ['월', '매출액', '비용', '이익', '이익률', '전월대비 매출', '전월대비 이익률', '누적 매출'];
  const mCols    = [10, 18, 18, 18, 10, 16, 16, 18];
  sheetMonth.getRow(2).height = 22;
  mHeaders.forEach((h, i) => {
    sheetMonth.getColumn(i + 1).width = mCols[i];
    const cell = sheetMonth.getCell(2, i + 1);
    cell.value = h;
    headerStyle(cell);
  });

  let cumRev = 0;
  months.forEach((m, idx) => {
    const v = byMonth[m];
    const prev = idx > 0 ? byMonth[months[idx - 1]] : null;
    const revDiff  = prev ? ((v.매출액 - prev.매출액) / prev.매출액 * 100).toFixed(1) + '%' : '-';
    const rateDiff = prev
      ? ((v.이익/v.매출액 - prev.이익/prev.매출액) * 100).toFixed(2) + '%p'
      : '-';
    cumRev += v.매출액;

    const row = sheetMonth.getRow(idx + 3);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const values = [
      m.replace('2025-', '') + '월',
      v.매출액, v.비용, v.이익,
      parseFloat((v.이익 / v.매출액 * 100).toFixed(1)),
      revDiff, rateDiff, cumRev,
    ];
    values.forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      const numFmt = [1,2,3,7].includes(ci) ? '#,##0' : (ci === 4 ? '0.0"%"' : null);
      dataCell(cell, isEven, numFmt);
      // 이익률 컬럼 색상
      if (ci === 4 && typeof val === 'number') {
        cell.font = { ...cell.font, color: { argb: val >= 36 ? 'FF375623' : val <= 34 ? 'FFC00000' : 'FF000000' } };
      }
    });
  });

  // 합계 행
  const sumRow = sheetMonth.getRow(months.length + 3);
  sumRow.height = 22;
  ['합계', total.매출액, total.비용, total.이익, parseFloat(profitRate), '', '', total.매출액].forEach((val, ci) => {
    const cell = sumRow.getCell(ci + 1);
    cell.value = val;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: '맑은 고딕' };
    cell.alignment = { horizontal: [0].includes(ci) ? 'center' : 'right', vertical: 'middle' };
    if ([1,2,3,7].includes(ci)) cell.numFmt = '#,##0';
    if (ci === 4) cell.numFmt = '0.0"%"';
  });

  // 차트 삽입
  await addChart(sheetMonth, imgMonthBar,  { col: 10, row: 2 }, 750, 360);
  await addChart(sheetMonth, imgMonthLine, { col: 10, row: 24 }, 750, 360);

  console.log('✓ 월별 분석 시트 완료');

  // ═══════════════════════════════════════════════════════════
  // 시트 3: 부서 / 담당자
  // ═══════════════════════════════════════════════════════════
  const sheetOrg = workbook.addWorksheet('🏢 조직별 분석');
  sheetOrg.views = [{ showGridLines: false }];

  // 부서별 테이블
  sheetOrg.mergeCells('A1:F1');
  Object.assign(sheetOrg.getCell('A1'), {
    value: '부서별 실적',
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } },
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: '맑은 고딕' },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  sheetOrg.getRow(1).height = 26;

  const deptHeaders = ['부서', '매출액', '비용', '이익', '이익률', '건수'];
  const deptWidths  = [16, 20, 20, 20, 10, 8];
  deptHeaders.forEach((h, i) => {
    sheetOrg.getColumn(i + 1).width = deptWidths[i];
    const cell = sheetOrg.getCell(2, i + 1);
    cell.value = h;
    headerStyle(cell);
  });
  sheetOrg.getRow(2).height = 22;

  Object.entries(byDept).sort((a,b) => b[1].매출액 - a[1].매출액).forEach(([dept, v], idx) => {
    const row = sheetOrg.getRow(idx + 3);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const rate = parseFloat((v.이익/v.매출액*100).toFixed(1));
    [dept, v.매출액, v.비용, v.이익, rate, v.건수].forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      dataCell(cell, isEven, [1,2,3].includes(ci) ? '#,##0' : ci===4 ? '0.0"%"' : null);
      if (ci === 0) cell.font = { bold: true, size: 10, name: '맑은 고딕' };
    });
  });

  // 담당자 테이블 (H열부터)
  const mgrStartCol = 8;
  sheetOrg.mergeCells(1, mgrStartCol, 1, mgrStartCol + 5);
  Object.assign(sheetOrg.getCell(1, mgrStartCol), {
    value: '담당자별 실적',
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } },
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: '맑은 고딕' },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });

  const mgrHeaders = ['담당자', '매출액', '이익', '이익률', '건수', '건당 매출'];
  const mgrWidths  = [12, 20, 20, 10, 8, 18];
  mgrHeaders.forEach((h, i) => {
    sheetOrg.getColumn(mgrStartCol + i).width = mgrWidths[i];
    const cell = sheetOrg.getCell(2, mgrStartCol + i);
    cell.value = h;
    headerStyle(cell);
  });

  Object.entries(byMgr).sort((a,b) => b[1].매출액 - a[1].매출액).forEach(([mgr, v], idx) => {
    const row = sheetOrg.getRow(idx + 3);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const rate = parseFloat((v.이익/v.매출액*100).toFixed(1));
    const perDeal = Math.round(v.매출액 / v.건수);
    [mgr, v.매출액, v.이익, rate, v.건수, perDeal].forEach((val, ci) => {
      const cell = row.getCell(mgrStartCol + ci);
      cell.value = val;
      dataCell(cell, isEven, [1,2,5].includes(ci) ? '#,##0' : ci===3 ? '0.0"%"' : null);
      if (ci === 0) cell.font = { bold: true, size: 10, name: '맑은 고딕' };
    });
  });

  // 차트 삽입
  await addChart(sheetOrg, imgDept, { col: 1, row: 12 }, 750, 400);
  await addChart(sheetOrg, imgMgr,  { col: 8, row: 12 }, 750, 400);

  console.log('✓ 조직별 분석 시트 완료');

  // ═══════════════════════════════════════════════════════════
  // 시트 4: 항목 / 거래처
  // ═══════════════════════════════════════════════════════════
  const sheetBiz = workbook.addWorksheet('📦 상품·거래처 분석');
  sheetBiz.views = [{ showGridLines: false }];

  // 항목 테이블
  sheetBiz.mergeCells('A1:F1');
  Object.assign(sheetBiz.getCell('A1'), {
    value: '항목별 실적 (이익 기준 정렬)',
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } },
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: '맑은 고딕' },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  sheetBiz.getRow(1).height = 26;

  const itemHeaders = ['항목', '매출액', '이익', '이익률', '건수', '순위'];
  const itemWidths  = [22, 20, 20, 10, 8, 8];
  itemHeaders.forEach((h, i) => {
    sheetBiz.getColumn(i + 1).width = itemWidths[i];
    const cell = sheetBiz.getCell(2, i + 1);
    cell.value = h;
    headerStyle(cell);
  });
  sheetBiz.getRow(2).height = 22;

  Object.entries(byItem).sort((a,b) => b[1].이익 - a[1].이익).forEach(([item, v], idx) => {
    const row = sheetBiz.getRow(idx + 3);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const rate = parseFloat((v.이익/v.매출액*100).toFixed(1));
    [item, v.매출액, v.이익, rate, v.건수, idx + 1].forEach((val, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = val;
      dataCell(cell, isEven, [1,2].includes(ci) ? '#,##0' : ci===3 ? '0.0"%"' : null);
      if (ci === 3) {
        cell.font = { ...cell.font, color: { argb: val >= 37 ? 'FF375623' : val <= 33 ? 'FFC00000' : 'FF000000' }, bold: val >= 37 };
      }
    });
  });

  // 거래처 테이블 (H열)
  const cStartCol = 8;
  sheetBiz.mergeCells(1, cStartCol, 1, cStartCol + 5);
  Object.assign(sheetBiz.getCell(1, cStartCol), {
    value: '거래처별 실적 TOP15',
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } },
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: '맑은 고딕' },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });

  const cHeaders = ['거래처', '매출액', '이익', '이익률', '건수', '비고'];
  const cWidths  = [20, 20, 20, 10, 8, 10];
  cHeaders.forEach((h, i) => {
    sheetBiz.getColumn(cStartCol + i).width = cWidths[i];
    const cell = sheetBiz.getCell(2, cStartCol + i);
    cell.value = h;
    headerStyle(cell);
  });

  Object.entries(byClient).sort((a,b) => b[1].매출액 - a[1].매출액).slice(0, 15).forEach(([client, v], idx) => {
    const row = sheetBiz.getRow(idx + 3);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const rate = parseFloat((v.이익/v.매출액*100).toFixed(1));
    const note = idx === 0 ? '⭐ 1위' : rate >= 37 ? '고마진' : rate <= 33 ? '저마진' : '';
    [client, v.매출액, v.이익, rate, v.건수, note].forEach((val, ci) => {
      const cell = row.getCell(cStartCol + ci);
      cell.value = val;
      dataCell(cell, isEven, [1,2].includes(ci) ? '#,##0' : ci===3 ? '0.0"%"' : null);
      if (ci === 0 && idx === 0) cell.font = { ...cell.font, bold: true, color: { argb: 'FFC00000' } };
    });
  });

  await addChart(sheetBiz, imgItemPie, { col: 1,  row: 28 }, 750, 400);
  await addChart(sheetBiz, imgClient,  { col: 8,  row: 28 }, 750, 400);

  console.log('✓ 상품·거래처 분석 시트 완료');

  // ═══════════════════════════════════════════════════════════
  // 시트 5: 인사이트 요약
  // ═══════════════════════════════════════════════════════════
  const sheetInsight = workbook.addWorksheet('💡 인사이트');
  sheetInsight.views = [{ showGridLines: false }];
  for (let c = 1; c <= 12; c++) sheetInsight.getColumn(c).width = 14;

  sheetInsight.mergeCells('A1:L1');
  Object.assign(sheetInsight.getCell('A1'), {
    value: '2025년 매출 분석 — 핵심 인사이트',
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } },
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 16, name: '맑은 고딕' },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  sheetInsight.getRow(1).height = 36;

  const insights = [
    {
      emoji: '📈', title: '하반기 성장세 뚜렷',
      body: `상반기 평균 매출 112억 vs 하반기 140억 (▲25%). 특히 10~12월 3개월 연속 최고 기록 경신. 12월이 연간 최고(158억)`,
      bg: 'FF1F4E79',
    },
    {
      emoji: '💰', title: '12월 이익률 최고 (37.3%)',
      body: `연중 이익률이 33.8%~37.3% 범위. 2월·5월이 상대적 저점. 8월(36.9%)·11월(36.3%)은 비용 효율이 좋은 달`,
      bg: 'FF2E75B6',
    },
    {
      emoji: '🏆', title: '브랜딩·일본수출 고마진 항목',
      body: `브랜딩 컨설팅(38.1%), 일본 수출(38.0%), 엔터프라이즈(37.8%), 광고 대행(37.3%) — 집중 투자 권고`,
      bg: 'FF375623',
    },
    {
      emoji: '⚠️', title: '싱가포르·현지기술지원 저마진',
      body: `싱가포르 프로젝트(32.9%), 현지 기술지원(33.2%), 콘텐츠 제작(33.5%) — 원가 구조 재검토 필요`,
      bg: 'FFC00000',
    },
    {
      emoji: '👑', title: '케이에스테크 압도적 1위 거래처',
      body: `케이에스테크 148억으로 전체의 10.5%. 2위 미래테크(115억)와 격차 큼. 동아시스템즈는 거래 건수 대비 이익률 36.6%로 우수`,
      bg: 'FF1F4E79',
    },
    {
      emoji: '🌟', title: '임도현 담당자 1위, 이지현 이익률 1위',
      body: `임도현 166억으로 매출 1위. 이지현(36.4%), 조은비(36.0%), 윤재혁(35.9%)이 이익률 상위. 강민지·최영호는 이익률 개선 여지`,
      bg: 'FF375623',
    },
  ];

  let iRow = 3;
  insights.forEach((ins, i) => {
    sheetInsight.getRow(iRow).height = 10;
    iRow++;

    // 제목 행
    sheetInsight.mergeCells(iRow, 1, iRow, 12);
    const titleC = sheetInsight.getCell(iRow, 1);
    titleC.value = `${ins.emoji}  ${ins.title}`;
    titleC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ins.bg } };
    titleC.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: '맑은 고딕' };
    titleC.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    sheetInsight.getRow(iRow).height = 26;
    iRow++;

    // 내용 행
    sheetInsight.mergeCells(iRow, 1, iRow + 1, 12);
    const bodyC = sheetInsight.getCell(iRow, 1);
    bodyC.value = ins.body;
    bodyC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i%2===0 ? 'FFD6E4F0' : 'FFE2EFDA' } };
    bodyC.font = { size: 11, name: '맑은 고딕' };
    bodyC.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 2 };
    sheetInsight.getRow(iRow).height = 22;
    sheetInsight.getRow(iRow + 1).height = 22;
    iRow += 2;
  });

  // 액션 아이템
  iRow += 2;
  sheetInsight.mergeCells(iRow, 1, iRow, 12);
  const actTitle = sheetInsight.getCell(iRow, 1);
  actTitle.value = '📋  2026년 권고 액션 아이템';
  actTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF595959' } };
  actTitle.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: '맑은 고딕' };
  actTitle.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  sheetInsight.getRow(iRow).height = 28;
  iRow++;

  const actions = [
    '① 브랜딩 컨설팅·일본 수출·엔터프라이즈 라인업 확대 — 고마진 항목 비중 상향 목표',
    '② 싱가포르 프로젝트·현지 기술지원 원가 절감 방안 수립 (아웃소싱·자동화 검토)',
    '③ 케이에스테크 집중 관리 + 동아시스템즈 거래 확대 (이익률·관계 모두 우수)',
    '④ 강민지·최영호 담당자 영업 지원 강화 — 이익률 향상을 위한 협상 교육',
    '⑤ 상반기(1~2월) 비수기 대응 프로모션 기획 — 연간 매출 평탄화 목표',
  ];

  actions.forEach((act, i) => {
    sheetInsight.mergeCells(iRow, 1, iRow, 12);
    const cell = sheetInsight.getCell(iRow, 1);
    cell.value = act;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i%2===0 ? 'FFFFF2CC' : 'FFFFFFCC' } };
    cell.font = { size: 11, name: '맑은 고딕' };
    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 2 };
    sheetInsight.getRow(iRow).height = 24;
    iRow++;
  });

  console.log('✓ 인사이트 시트 완료');

  // ═══════════════════════════════════════════════════════════
  // 원본 데이터 시트 복사 (기존 통합 데이터)
  // ═══════════════════════════════════════════════════════════
  const STANDARD_COLUMNS = ['월', '부서', '항목', '거래처', '담당자', '매출액', '비용', '이익', '비고'];
  const COL_WIDTHS = { '월':12,'부서':13,'항목':20,'거래처':20,'담당자':11,'매출액':18,'비용':18,'이익':18,'비고':10 };
  const sheetRaw = workbook.addWorksheet('📋 원본 데이터');
  sheetRaw.columns = STANDARD_COLUMNS.map(col => ({ header: col, key: col, width: COL_WIDTHS[col] }));

  const rawHdr = sheetRaw.getRow(1);
  rawHdr.height = 22;
  rawHdr.eachCell(cell => headerStyle(cell));

  rows.forEach((rowData, idx) => {
    const row = sheetRaw.addRow(rowData);
    row.height = 18;
    const isEven = idx % 2 === 0;
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const col = STANDARD_COLUMNS[colNum - 1];
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFD6E4F0' : 'FFFFFFFF' } };
      cell.font = { size: 10, name: '맑은 고딕' };
      cell.border = {
        top: { style: 'hair', color: { argb: 'FFBDD7EE' } }, bottom: { style: 'hair', color: { argb: 'FFBDD7EE' } },
        left: { style: 'hair', color: { argb: 'FFBDD7EE' } }, right: { style: 'hair', color: { argb: 'FFBDD7EE' } },
      };
      if (['매출액','비용','이익'].includes(col)) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else if (col === '월') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { ...cell.font, bold: true, color: { argb: 'FF1F4E79' } };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  });
  sheetRaw.views = [{ state: 'frozen', ySplit: 1 }];
  sheetRaw.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: STANDARD_COLUMNS.length } };

  console.log('✓ 원본 데이터 시트 완료');

  // 저장
  await workbook.xlsx.writeFile(DEST);
  console.log(`\n✅ 저장 완료: ${DEST}`);
}

main().catch(console.error);
