/**
 * generate_sales_report.js
 *
 * 월간 매출 엑셀 12개 → 통합 엑셀 + 대시보드 엑셀 + PPT 자동 생성
 *
 * 사용법:
 *   node generate_sales_report.js [월간파일_폴더] [결과물_저장_폴더]
 *
 * 예시:
 *   node generate_sales_report.js "D:/2026년/월간매출" "D:/2026년/결과물"
 *   node generate_sales_report.js ./월간매출          (결과물은 상위 폴더에 저장)
 */

const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');

// ── 경로 설정 ──────────────────────────────────────────────
const inputDir  = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(__dirname, '월간매출');

const outputDir = process.argv[3]
  ? path.resolve(process.argv[3])
  : path.dirname(inputDir);  // 폴더를 지정 안 하면 월간파일 폴더의 상위에 저장

// 입력 폴더 검증
if (!fs.existsSync(inputDir)) {
  console.error(`\n❌ 입력 폴더를 찾을 수 없습니다: ${inputDir}`);
  console.error('   사용법: node generate_sales_report.js [월간파일_폴더] [결과물_저장_폴더]');
  process.exit(1);
}

const xlsxFiles = fs.readdirSync(inputDir).filter(f => f.endsWith('.xlsx')).sort();

if (xlsxFiles.length === 0) {
  console.error(`\n❌ "${inputDir}" 안에 xlsx 파일이 없습니다.`);
  process.exit(1);
}

if (xlsxFiles.length !== 12) {
  console.warn(`\n⚠️  xlsx 파일이 ${xlsxFiles.length}개입니다 (12개 권장). 그대로 진행합니다.`);
}

// 파일명에서 연도 추출 (예: "2026년_01월_매출.xlsx" → "2026")
const yearMatch = xlsxFiles[0].match(/(\d{4})년/);
const year = yearMatch ? yearMatch[1] : new Date().getFullYear();

if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

// 출력 파일 경로
const unifiedXlsx   = path.join(outputDir, `${year}_연간매출_통합.xlsx`);
const dashboardXlsx = path.join(outputDir, `${year}_연간매출_대시보드.xlsx`);
const reportPptx    = path.join(outputDir, `${year}_연간매출_보고서.pptx`);

const scriptsDir = __dirname;
const env = {
  ...process.env,
  MONTHLY_DIR:    inputDir,
  UNIFIED_XLSX:   unifiedXlsx,
  DASHBOARD_XLSX: dashboardXlsx,
  REPORT_PPTX:    reportPptx,
};

console.log('='.repeat(60));
console.log('  월간 매출 자동 보고서 생성기');
console.log('='.repeat(60));
console.log(`📂 입력 폴더  : ${inputDir}`);
console.log(`📂 출력 폴더  : ${outputDir}`);
console.log(`📅 대상 연도  : ${year}년`);
console.log(`📄 파일 수    : ${xlsxFiles.length}개`);
console.log('-'.repeat(60));

const steps = [
  { label: '1단계: 12개 파일 통합 (헤더 매핑 + 숫자 포맷 통일)', script: 'merge_sales.js' },
  { label: '2단계: 엑셀 대시보드 생성 (차트 6종 + 분석 시트)', script: 'make_dashboard.js' },
  { label: '3단계: PPT 보고서 생성 (11슬라이드)',               script: 'make_ppt.js' },
];

const startTotal = Date.now();

steps.forEach(({ label, script }) => {
  console.log(`\n▶ ${label}`);
  const t0 = Date.now();
  try {
    execSync(`node "${path.join(scriptsDir, script)}"`, { env, stdio: 'inherit' });
    console.log(`  ✅ 완료 (${((Date.now()-t0)/1000).toFixed(1)}s)`);
  } catch (e) {
    console.error(`  ❌ 오류 발생: ${script}`);
    process.exit(1);
  }
});

console.log('\n' + '='.repeat(60));
console.log(`  🎉 모든 작업 완료! (총 ${((Date.now()-startTotal)/1000).toFixed(1)}s)`);
console.log('='.repeat(60));
console.log(`\n📊 통합 엑셀     : ${unifiedXlsx}`);
console.log(`📊 대시보드 엑셀 : ${dashboardXlsx}`);
console.log(`📑 PPT 보고서    : ${reportPptx}`);
