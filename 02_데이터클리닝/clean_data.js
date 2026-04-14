const XLSX = require('xlsx');
const path = require('path');

const inputFile = path.join(__dirname, '거래처_명단_원본.xlsx');
const outputFile = path.join(__dirname, '거래처_명단_정리.xlsx');

// ─────────────────────────────────────────
// 정리 기준
// 1. 중복 제거: 동일 담당자+연락처 기준으로 중복 판별, 두 행의 정보를 합쳐서 가장 완전한 데이터 유지
// 2. 회사명 정규화: (주)를 앞으로 통일 → (주)회사명
// 3. 주소 정규화: 서울→서울특별시, 경기→경기도, 인천시→인천광역시
// 4. 빈 행 / 회사명 없는 행 제거
// 5. N/A, 담당자 미정 → 빈값 처리
// 6. 이메일 충돌 시 더 공식적인 포맷(이니셜.성) 우선 채택, 불일치 표기
// 7. 거래중단 업체: 비고에 "[거래중단]" 통일 표기
// 8. 이메일 도메인 불일치: 비고에 "[이메일 확인필요]" 표기
// ─────────────────────────────────────────

const wb = XLSX.readFile(inputFile);
const ws = wb.Sheets[wb.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

const headers = rawData[0];
const rows = rawData.slice(1);

// 컬럼 인덱스
const COL = {
  회사명: 0, 담당자: 1, 직급: 2, 연락처: 3,
  이메일: 4, 주소: 5, 거래시작일: 6, 거래유형: 7, 비고: 8
};

// ── 유틸 함수 ──────────────────────────────

function normalizeCompanyName(name) {
  if (!name || !name.trim()) return '';
  name = name.trim();
  // (주) 뒤에 있으면 앞으로 이동
  name = name.replace(/^(.+)\(주\)$/, '(주)$1');
  // 이미 (주)로 시작하면 그대로, 없으면 앞에 추가
  // 단, 교육/협회 등 명백히 주식회사가 아닐 수 있는 경우는 판단 불가능하므로
  // 두 버전이 존재할 때만 (주) 추가 (mergeRows에서 처리)
  return name;
}

function normalizeAddress(addr) {
  if (!addr) return '';
  addr = addr.trim();
  // 서울 → 서울특별시 (서울시도 통일)
  addr = addr.replace(/^서울시\s/, '서울특별시 ');
  addr = addr.replace(/^서울\s/, '서울특별시 ');
  // 경기 → 경기도
  addr = addr.replace(/^경기\s/, '경기도 ');
  // 인천시 → 인천광역시
  addr = addr.replace(/^인천시\s/, '인천광역시 ');
  addr = addr.replace(/^인천\s/, '인천광역시 ');
  return addr;
}

function normalizeDate(val) {
  if (!val) return '';
  const s = String(val).trim();
  // 이미 YYYY-MM-DD 형태면 그대로
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function isNAValue(val) {
  return !val || val.toString().trim() === '' || val.toString().trim().toUpperCase() === 'N/A';
}

function pickEmail(e1, e2) {
  // 둘 다 없으면 빈값
  if (!e1 && !e2) return { email: '', conflict: false };
  if (!e1) return { email: e2, conflict: false };
  if (!e2) return { email: e1, conflict: false };
  if (e1 === e2) return { email: e1, conflict: false };
  // 다르면: firstName.lastName 형태(점 포함) 우선
  const hasDot1 = e1.split('@')[0].includes('.');
  const hasDot2 = e2.split('@')[0].includes('.');
  if (hasDot1 && !hasDot2) return { email: e1, conflict: true };
  if (!hasDot1 && hasDot2) return { email: e2, conflict: true };
  // 둘 다 점 있으면 첫 번째 우선
  return { email: e1, conflict: true };
}

// ── 행 객체 변환 ────────────────────────────

function rowToObj(row) {
  return {
    회사명: (row[COL.회사명] || '').trim(),
    담당자: isNAValue(row[COL.담당자]) ? '' : (row[COL.담당자] || '').trim(),
    직급: isNAValue(row[COL.직급]) ? '' : (row[COL.직급] || '').trim(),
    연락처: (row[COL.연락처] || '').trim(),
    이메일: (row[COL.이메일] || '').trim(),
    주소: normalizeAddress((row[COL.주소] || '').trim()),
    거래시작일: normalizeDate(row[COL.거래시작일]),
    거래유형: (row[COL.거래유형] || '').trim(),
    비고: (row[COL.비고] || '').trim(),
  };
}

// ── 두 행이 중복인지 판단 ────────────────────────────────
// 담당자가 같고, 연락처 또는 이메일이 같거나 한쪽이 비어있으면 중복

function isSamePerson(a, b) {
  if (!a.담당자 || !b.담당자) return false; // 담당자 없으면 중복 판단 불가
  if (a.담당자 !== b.담당자) return false;
  // 연락처 비교
  const 연일치 = a.연락처 && b.연락처 && a.연락처 === b.연락처;
  // 이메일 비교 (도메인까지만)
  const aDomain = a.이메일.includes('@') ? a.이메일.split('@')[1] : '';
  const bDomain = b.이메일.includes('@') ? b.이메일.split('@')[1] : '';
  const 이메일일치 = aDomain && bDomain && aDomain === bDomain;
  // 한쪽 연락처가 비어있고 이메일이 같으면 중복
  const 연락처하나비어있음 = (!a.연락처 || !b.연락처) && 이메일일치;
  return 연일치 || 이메일일치 || 연락처하나비어있음;
}

// ── 두 행 병합 ────────────────────────────────
// a를 기준으로 b의 비어있는 필드를 채움

function mergeRows(a, b) {
  const merged = { ...a };

  // 회사명: (주) 붙은 버전 우선
  const names = [a.회사명, b.회사명].filter(Boolean);
  const withJoo = names.find(n => n.includes('(주)'));
  if (withJoo) {
    merged.회사명 = normalizeCompanyName(withJoo);
  } else {
    merged.회사명 = normalizeCompanyName(a.회사명 || b.회사명);
  }

  // 나머지 필드: 비어있으면 b에서 채움
  for (const key of ['담당자', '직급', '연락처', '주소', '거래시작일', '거래유형']) {
    if (!merged[key] && b[key]) merged[key] = b[key];
  }

  // 이메일: 충돌 처리
  const { email, conflict } = pickEmail(a.이메일, b.이메일);
  merged.이메일 = email;

  // 비고: 두 행 비고를 합침 (중복 제거)
  const notes = [];
  const addNote = (note) => {
    if (!note) return;
    note = note.trim();
    if (note && !notes.some(n => n.toLowerCase() === note.toLowerCase())) {
      notes.push(note);
    }
  };
  addNote(a.비고);
  addNote(b.비고);
  if (conflict) addNote('[이메일 확인필요]');

  // 내부 메모성 비고 제거 (중복 표시, 도메인확인 원문 등)
  const filteredNotes = notes.filter(n => {
    if (/\d번째\s*중복/.test(n)) return false;       // "3번째 중복" 제거
    if (/회사명\s*확인\s*필요/.test(n)) return false; // "회사명 확인필요" 제거
    if (/이메일\s*도메인\s*확인/.test(n)) return false; // "이메일 도메인 확인" 제거 (→ [이메일 확인필요]로 대체)
    if (/연락처\s*누락/.test(n)) return false;        // "연락처 누락" 제거 (데이터로 해결됨)
    return true;
  });

  // 거래중단 정규화
  const noteStr = filteredNotes.join(' / ');
  if (/거래\s*중단/.test(noteStr)) {
    const withoutOld = filteredNotes.filter(n => !/거래\s*중단/.test(n));
    merged.비고 = ['[거래중단]', ...withoutOld].join(' / ');
  } else {
    merged.비고 = noteStr;
  }

  return merged;
}

// ── 메인 처리 ────────────────────────────────

const validRows = rows
  .map(rowToObj)
  .filter(obj => obj.회사명); // 회사명 없는 행 제거

// 다온인터내셔날 → 다온인터내셔널 교정
validRows.forEach(obj => {
  obj.회사명 = obj.회사명.replace('인터내셔날', '인터내셔널');
});

// 중복 그룹핑 (Union-Find 방식)
const used = new Array(validRows.length).fill(false);
const groups = [];

for (let i = 0; i < validRows.length; i++) {
  if (used[i]) continue;
  const group = [validRows[i]];
  used[i] = true;
  for (let j = i + 1; j < validRows.length; j++) {
    if (!used[j] && isSamePerson(validRows[i], validRows[j])) {
      group.push(validRows[j]);
      used[j] = true;
    }
  }
  groups.push(group);
}

// 그룹별 병합
const cleaned = [];
for (const group of groups) {
  let merged = group[0];
  for (let i = 1; i < group.length; i++) {
    merged = mergeRows(merged, group[i]);
  }
  cleaned.push(merged);
}

// 정렬: 거래유형(A→B→C→기타), 회사명 가나다순
const gradeOrder = { 'A등급': 0, 'B등급': 1, 'C등급': 2 };
cleaned.sort((a, b) => {
  const ga = gradeOrder[a.거래유형] ?? 9;
  const gb = gradeOrder[b.거래유형] ?? 9;
  if (ga !== gb) return ga - gb;
  return a.회사명.localeCompare(b.회사명, 'ko');
});

// ── 결과 출력 ────────────────────────────────

console.log('=== 정리 결과 ===');
console.log(`원본 데이터 행 수: ${rows.length}`);
console.log(`유효 행 수 (회사명 있는): ${validRows.length}`);
console.log(`정리 후 행 수: ${cleaned.length}`);
console.log(`제거된 중복/빈 행 수: ${rows.length - cleaned.length}`);
console.log('');
console.log('=== 최종 데이터 ===');
cleaned.forEach((row, i) => {
  console.log(`${i + 1}. ${row.회사명} | ${row.담당자} ${row.직급} | ${row.연락처} | ${row.이메일} | ${row.거래유형} | 비고: ${row.비고}`);
});

// ── Excel 파일 생성 ───────────────────────────

const outputData = [
  ['회사명', '담당자', '직급', '연락처', '이메일', '주소', '거래시작일', '거래유형', '비고'],
  ...cleaned.map(r => [r.회사명, r.담당자, r.직급, r.연락처, r.이메일, r.주소, r.거래시작일, r.거래유형, r.비고])
];

const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.aoa_to_sheet(outputData);

// 컬럼 너비 설정
newWs['!cols'] = [
  { wch: 22 }, // 회사명
  { wch: 8 },  // 담당자
  { wch: 6 },  // 직급
  { wch: 14 }, // 연락처
  { wch: 28 }, // 이메일
  { wch: 36 }, // 주소
  { wch: 12 }, // 거래시작일
  { wch: 8 },  // 거래유형
  { wch: 30 }, // 비고
];

XLSX.utils.book_append_sheet(newWb, newWs, '거래처 명단');
XLSX.writeFile(newWb, outputFile);

console.log(`\n✅ 저장 완료: ${outputFile}`);
