const XLSX = require('../node_modules/xlsx');
const ExcelJS = require('../node_modules/exceljs');
const fs = require('fs');
const path = require('path');

const folder = process.env.MONTHLY_DIR || path.join(__dirname, '월간매출');
const outputPath = process.env.UNIFIED_XLSX || path.join(__dirname, '2025_연간매출_통합.xlsx');

// 헤더 매핑: 각 파일의 다양한 헤더명 → 표준 컬럼명
const HEADER_MAP = {
  '부서': '부서', '팀명': '부서', 'Department': '부서',
  '항목': '항목', '매출항목': '항목', 'Category': '항목',
  '거래처': '거래처', '거래처명': '거래처', 'Client': '거래처',
  '담당자': '담당자', '담당': '담당자', 'Manager': '담당자',
  '매출액': '매출액', '매출(원)': '매출액', 'Revenue': '매출액',
  '비용': '비용', '원가(원)': '비용', 'Cost': '비용',
  '이익': '이익', '이익(원)': '이익', 'Profit': '이익', '영업이익': '이익',
  '비고': '비고',
};

const STANDARD_COLUMNS = ['월', '부서', '항목', '거래처', '담당자', '매출액', '비용', '이익', '비고'];

const COL_WIDTHS = {
  '월': 12, '부서': 13, '항목': 20, '거래처': 20,
  '담당자': 11, '매출액': 18, '비용': 18, '이익': 18, '비고': 10
};

function extractMonth(filename) {
  const match = filename.match(/(\d{4})년_(\d{2})월/);
  if (match) return `${match[1]}-${match[2]}`;
  return null;
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return Math.round(value);
  const parsed = Number(String(value).replace(/,/g, ''));
  return isNaN(parsed) ? value : Math.round(parsed);
}

// ── 1단계: xlsx로 데이터 수집 ──────────────────────────────
const allRows = [];

const files = fs.readdirSync(folder)
  .filter(f => f.endsWith('.xlsx'))
  .sort();

console.log(`처리할 파일 수: ${files.length}개\n`);

files.forEach(filename => {
  const monthLabel = extractMonth(filename);
  const wb = XLSX.readFile(path.join(folder, filename));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (rawData.length < 2) { console.log(`[경고] ${filename}: 데이터 없음`); return; }

  const colMapping = {};
  rawData[0].forEach((h, i) => {
    const std = HEADER_MAP[h];
    if (std) colMapping[std] = i;
  });

  let rowCount = 0;
  for (let i = 1; i < rawData.length; i++) {
    const raw = rawData[i];
    if (raw.every(c => c === '' || c === null || c === undefined)) continue;

    const row = { '월': monthLabel };
    STANDARD_COLUMNS.forEach(col => {
      if (col === '월') return;
      const idx = colMapping[col];
      if (idx !== undefined) {
        const val = raw[idx];
        row[col] = ['매출액', '비용', '이익'].includes(col) ? normalizeNumber(val) : (val ?? '');
      } else {
        row[col] = ['매출액', '비용', '이익'].includes(col) ? null : '';
      }
    });

    allRows.push(row);
    rowCount++;
  }
  console.log(`✓ ${filename} (${monthLabel}): ${rowCount}행 처리`);
});

console.log(`\n총 ${allRows.length}행 통합`);

// ── 2단계: ExcelJS로 스타일 적용하며 저장 ────────────────────
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('2025_연간매출');

// 컬럼 정의
sheet.columns = STANDARD_COLUMNS.map(col => ({
  header: col,
  key: col,
  width: COL_WIDTHS[col] || 12,
}));

// 헤더 행 스타일
const headerRow = sheet.getRow(1);
headerRow.height = 22;
headerRow.eachCell(cell => {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F4E79' },   // 진한 네이비
  };
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: '맑은 고딕' };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = {
    top:    { style: 'thin', color: { argb: 'FF2E75B6' } },
    bottom: { style: 'thin', color: { argb: 'FF2E75B6' } },
    left:   { style: 'thin', color: { argb: 'FF2E75B6' } },
    right:  { style: 'thin', color: { argb: 'FF2E75B6' } },
  };
});

// 데이터 행 추가 및 스타일
const NUM_COLS = new Set(['매출액', '비용', '이익']);

allRows.forEach((rowData, idx) => {
  const row = sheet.addRow(rowData);
  row.height = 18;

  const isEven = idx % 2 === 0;

  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const colName = STANDARD_COLUMNS[colNumber - 1];

    // 줄 배경: 짝수/홀수 교대
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: isEven ? 'FFD6E4F0' : 'FFFFFFFF' },  // 연파랑 / 흰색
    };

    // 폰트
    cell.font = { size: 10, name: '맑은 고딕' };

    // 테두리
    cell.border = {
      top:    { style: 'hair', color: { argb: 'FFBDD7EE' } },
      bottom: { style: 'hair', color: { argb: 'FFBDD7EE' } },
      left:   { style: 'hair', color: { argb: 'FFBDD7EE' } },
      right:  { style: 'hair', color: { argb: 'FFBDD7EE' } },
    };

    // 정렬 & 숫자 포맷
    if (NUM_COLS.has(colName)) {
      cell.numFmt = '#,##0';
      cell.alignment = { horizontal: 'right', vertical: 'middle' };

      // 이익 음수면 빨간 글씨
      if (colName === '이익' && typeof cell.value === 'number' && cell.value < 0) {
        cell.font = { ...cell.font, color: { argb: 'FFFF0000' } };
      }
    } else if (colName === '월') {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { ...cell.font, bold: true, color: { argb: 'FF1F4E79' } };
    } else {
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
    }
  });
});

// 틀 고정 (헤더 고정)
sheet.views = [{ state: 'frozen', ySplit: 1 }];

// 자동 필터
sheet.autoFilter = {
  from: { row: 1, column: 1 },
  to:   { row: 1, column: STANDARD_COLUMNS.length },
};

workbook.xlsx.writeFile(outputPath).then(() => {
  console.log(`\n출력 파일: ${outputPath}`);
});
