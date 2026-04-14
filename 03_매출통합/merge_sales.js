const XLSX = require('../node_modules/xlsx');
const fs = require('fs');
const path = require('path');

const folder = path.join(__dirname, '월간매출');
const outputPath = path.join(__dirname, '2025_연간매출_통합.xlsx');

// 헤더 매핑: 각 파일의 다양한 헤더명 → 표준 컬럼명
const HEADER_MAP = {
  // 부서
  '부서': '부서',
  '팀명': '부서',
  'Department': '부서',

  // 항목
  '항목': '항목',
  '매출항목': '항목',
  'Category': '항목',

  // 거래처
  '거래처': '거래처',
  '거래처명': '거래처',
  'Client': '거래처',

  // 담당자
  '담당자': '담당자',
  '담당': '담당자',
  'Manager': '담당자',

  // 매출액
  '매출액': '매출액',
  '매출(원)': '매출액',
  'Revenue': '매출액',

  // 비용
  '비용': '비용',
  '원가(원)': '비용',
  'Cost': '비용',

  // 이익
  '이익': '이익',
  '이익(원)': '이익',
  'Profit': '이익',
  '영업이익': '이익',

  // 비고 (10~12월에만 존재)
  '비고': '비고',
};

// 표준 컬럼 순서
const STANDARD_COLUMNS = ['월', '부서', '항목', '거래처', '담당자', '매출액', '비용', '이익', '비고'];

// 파일명에서 월 추출: "2025년_01월_매출.xlsx" → "2025-01"
function extractMonth(filename) {
  const match = filename.match(/(\d{4})년_(\d{2})월/);
  if (match) return `${match[1]}-${match[2]}`;
  return null;
}

// 숫자 포맷 통일: 정수로 반올림, 숫자가 아니면 그대로
function normalizeNumber(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') return Math.round(value);
  const parsed = Number(String(value).replace(/,/g, ''));
  return isNaN(parsed) ? value : Math.round(parsed);
}

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

  if (rawData.length < 2) {
    console.log(`[경고] ${filename}: 데이터 없음`);
    return;
  }

  const rawHeaders = rawData[0];

  // 원본 헤더 → 표준 컬럼명 매핑 인덱스 생성
  const colMapping = {}; // 표준컬럼명 → 원본 인덱스
  rawHeaders.forEach((h, i) => {
    const standardName = HEADER_MAP[h];
    if (standardName) colMapping[standardName] = i;
  });

  // 매핑 안 된 필수 컬럼 경고
  ['부서', '항목', '거래처', '담당자', '매출액', '비용', '이익'].forEach(col => {
    if (colMapping[col] === undefined) {
      console.log(`[경고] ${filename}: '${col}' 컬럼을 찾지 못했습니다.`);
    }
  });

  let rowCount = 0;
  for (let i = 1; i < rawData.length; i++) {
    const rawRow = rawData[i];

    // 완전히 빈 행 스킵
    if (rawRow.every(cell => cell === '' || cell === null || cell === undefined)) continue;

    const row = {};
    row['월'] = monthLabel;

    STANDARD_COLUMNS.forEach(col => {
      if (col === '월') return;
      const idx = colMapping[col];
      if (idx !== undefined) {
        const val = rawRow[idx];
        // 숫자 컬럼 정규화
        if (['매출액', '비용', '이익'].includes(col)) {
          row[col] = normalizeNumber(val);
        } else {
          row[col] = val !== undefined && val !== null ? val : '';
        }
      } else {
        row[col] = '';
      }
    });

    allRows.push(row);
    rowCount++;
  }

  console.log(`✓ ${filename} (${monthLabel}): ${rowCount}행 처리`);
});

console.log(`\n총 ${allRows.length}행 통합 완료`);

// 출력 워크북 생성
const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.json_to_sheet(allRows, { header: STANDARD_COLUMNS });

// 컬럼 너비 설정
const colWidths = {
  '월': 12, '부서': 12, '항목': 18, '거래처': 18,
  '담당자': 10, '매출액': 16, '비용': 16, '이익': 16, '비고': 10
};
newWs['!cols'] = STANDARD_COLUMNS.map(col => ({ wch: colWidths[col] || 12 }));

// 숫자 셀에 천단위 콤마 포맷 적용
const numCols = ['매출액', '비용', '이익'];
const numColIdxs = STANDARD_COLUMNS
  .map((col, i) => numCols.includes(col) ? i : -1)
  .filter(i => i !== -1);

const range = XLSX.utils.decode_range(newWs['!ref']);
for (let R = range.s.r + 1; R <= range.e.r; R++) {
  numColIdxs.forEach(C => {
    const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
    const cell = newWs[cellAddr];
    if (cell && typeof cell.v === 'number') {
      cell.z = '#,##0';
    }
  });
}

XLSX.utils.book_append_sheet(newWb, newWs, '2025_연간매출');
XLSX.writeFile(newWb, outputPath);

console.log(`\n출력 파일: ${outputPath}`);
