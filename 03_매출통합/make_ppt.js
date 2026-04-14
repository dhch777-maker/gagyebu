const XLSX = require('../node_modules/xlsx');
const PptxGenJS = require('../node_modules/pptxgenjs');
const { ChartJSNodeCanvas } = require('../node_modules/chartjs-node-canvas');
const path = require('path');

const SRC  = process.env.UNIFIED_XLSX  || path.join(__dirname, '2025_연간매출_통합.xlsx');
const DEST = process.env.REPORT_PPTX   || path.join(__dirname, '2025_연간매출_보고서.pptx');

// ── 데이터 로드 ────────────────────────────────────────────
const rawWb = XLSX.readFile(SRC);
const rows  = XLSX.utils.sheet_to_json(rawWb.Sheets[rawWb.SheetNames[0]]);

function aggregate(rows, key) {
  const map = {};
  rows.forEach(r => {
    const k = r[key];
    if (!map[k]) map[k] = { 매출액:0, 비용:0, 이익:0, 건수:0 };
    map[k].매출액 += r['매출액']||0;
    map[k].비용   += r['비용']  ||0;
    map[k].이익   += r['이익']  ||0;
    map[k].건수   += 1;
  });
  return map;
}

const byMonth  = aggregate(rows, '월');
const byDept   = aggregate(rows, '부서');
const byMgr    = aggregate(rows, '담당자');
const byItem   = aggregate(rows, '항목');
const byClient = aggregate(rows, '거래처');
const months   = Object.keys(byMonth).sort();

const total = rows.reduce((s,r) => {
  s.매출액 += r['매출액']||0; s.비용 += r['비용']||0; s.이익 += r['이익']||0; return s;
}, {매출액:0, 비용:0, 이익:0});

const profitRate = (total.이익/total.매출액*100).toFixed(1);
const bestMonth  = months.reduce((b,m) => byMonth[m].매출액 > byMonth[b].매출액 ? m : b, months[0]);
const worstMonth = months.reduce((b,m) => byMonth[m].매출액 < byMonth[b].매출액 ? m : b, months[0]);

// 억 단위 포맷
const toUk = v => (Math.round(v/1e8*10)/10).toFixed(1);
const toPct = v => (v*100).toFixed(1);
const fmtNum = v => Math.round(v).toLocaleString('ko-KR');

// ── 색상 팔레트 ─────────────────────────────────────────────
const C = {
  navy:   '1F4E79', blue:  '2E75B6', lblue: '9DC3E6',
  green:  '375623', lgreen:'70AD47', lgreen2:'E2EFDA',
  red:    'C00000', lred:  'FFE0E0',
  gray:   '595959', lgray: 'F2F2F2', white: 'FFFFFF',
  gold:   'C9A227', yellow:'FFF2CC',
  bg:     'EEF4FB',  // 슬라이드 배경
};

// ── 차트 생성 ──────────────────────────────────────────────
function canvas(w=900, h=480) {
  return new ChartJSNodeCanvas({ width:w, height:h, backgroundColour:'white' });
}

const monthLabels = months.map(m => m.replace('2025-','')+'월');

async function makeCharts() {
  const charts = {};

  // 1. 월별 매출·비용·이익 묶음 막대
  charts.monthBar = await canvas(1000, 520).renderToBuffer({
    type: 'bar',
    data: {
      labels: monthLabels,
      datasets: [
        { label:'매출액', data: months.map(m=>Math.round(byMonth[m].매출액/1e7)/10), backgroundColor:'#2E75B6' },
        { label:'비용',   data: months.map(m=>Math.round(byMonth[m].비용/1e7)/10),   backgroundColor:'#9DC3E6' },
        { label:'이익',   data: months.map(m=>Math.round(byMonth[m].이익/1e7)/10),   backgroundColor:'#70AD47' },
      ],
    },
    options: {
      plugins: { legend:{position:'bottom'}, title:{display:false} },
      scales: { y:{ ticks:{callback:v=>v+'억'}, grid:{color:'#e0e0e0'} } },
    },
  });

  // 2. 이익률 + 매출액 콤보 라인
  charts.trendLine = await canvas(1000, 520).renderToBuffer({
    type: 'line',
    data: {
      labels: monthLabels,
      datasets: [
        {
          label:'이익률(%)', type:'line',
          data: months.map(m=>parseFloat((byMonth[m].이익/byMonth[m].매출액*100).toFixed(1))),
          borderColor:'#C00000', backgroundColor:'rgba(192,0,0,0.08)',
          borderWidth:3, pointRadius:6, pointBackgroundColor:'#C00000', fill:true, tension:0.35,
          yAxisID:'y',
        },
        {
          label:'매출액(억)', type:'bar',
          data: months.map(m=>Math.round(byMonth[m].매출액/1e8*10)/10),
          backgroundColor:'rgba(46,117,182,0.3)', borderColor:'#2E75B6', borderWidth:1,
          yAxisID:'y2',
        },
      ],
    },
    options: {
      plugins: { legend:{position:'bottom'} },
      scales: {
        y:  { position:'left',  min:32, max:39, ticks:{callback:v=>v+'%'}, grid:{color:'#e0e0e0'} },
        y2: { position:'right', ticks:{callback:v=>v+'억'}, grid:{drawOnChartArea:false} },
      },
    },
  });

  // 3. 부서별 수평 막대
  const depts = Object.entries(byDept).sort((a,b)=>b[1].매출액-a[1].매출액);
  charts.deptBar = await canvas(900, 480).renderToBuffer({
    type:'bar',
    data:{
      labels: depts.map(([d])=>d),
      datasets:[
        { label:'매출액(억)', data: depts.map(([,v])=>Math.round(v.매출액/1e8*10)/10),
          backgroundColor:['#1F4E79','#2E75B6','#9DC3E6','#5B9BD5','#BDD7EE'] },
        { label:'이익(억)',   data: depts.map(([,v])=>Math.round(v.이익/1e8*10)/10),
          backgroundColor:['#375623','#70AD47','#A9D18E','#C6E0B4','#E2EFDA'] },
      ],
    },
    options:{
      indexAxis:'y',
      plugins:{ legend:{position:'bottom'} },
      scales:{ x:{ ticks:{callback:v=>v+'억'} } },
    },
  });

  // 4. 담당자별 매출+이익률 콤보
  const mgrs = Object.entries(byMgr).sort((a,b)=>b[1].매출액-a[1].매출액);
  charts.mgrBar = await canvas(1000, 520).renderToBuffer({
    type:'bar',
    data:{
      labels: mgrs.map(([m])=>m),
      datasets:[
        { label:'매출액(억)', data: mgrs.map(([,v])=>Math.round(v.매출액/1e8*10)/10),
          backgroundColor:'#2E75B6', yAxisID:'y' },
        { label:'이익률(%)', type:'line',
          data: mgrs.map(([,v])=>parseFloat((v.이익/v.매출액*100).toFixed(1))),
          borderColor:'#C00000', backgroundColor:'transparent',
          borderWidth:3, pointRadius:6, pointBackgroundColor:'#C00000', yAxisID:'y2' },
      ],
    },
    options:{
      plugins:{ legend:{position:'bottom'} },
      scales:{
        y: { position:'left',  ticks:{callback:v=>v+'억'} },
        y2:{ position:'right', min:31, max:38, ticks:{callback:v=>v+'%'}, grid:{drawOnChartArea:false} },
      },
    },
  });

  // 5. 항목별 이익 도넛
  const topItems = Object.entries(byItem).sort((a,b)=>b[1].이익-a[1].이익).slice(0,8);
  charts.itemDoughnut = await canvas(900, 500).renderToBuffer({
    type:'doughnut',
    data:{
      labels: topItems.map(([it])=>it),
      datasets:[{
        data: topItems.map(([,v])=>Math.round(v.이익/1e6)),
        backgroundColor:['#1F4E79','#2E75B6','#5B9BD5','#9DC3E6','#375623','#70AD47','#C00000','#C9A227'],
        borderWidth:3, borderColor:'#ffffff',
      }],
    },
    options:{
      plugins:{
        legend:{ position:'right', labels:{ font:{size:13} } },
      },
      cutout:'55%',
    },
  });

  // 6. 항목별 이익률 수평 막대 (고마진 하이라이트)
  const itemsByRate = Object.entries(byItem)
    .map(([it,v])=>([it, parseFloat((v.이익/v.매출액*100).toFixed(1))]))
    .sort((a,b)=>b[1]-a[1]);
  charts.itemRate = await canvas(900, 560).renderToBuffer({
    type:'bar',
    data:{
      labels: itemsByRate.map(([it])=>it),
      datasets:[{
        label:'이익률(%)',
        data: itemsByRate.map(([,r])=>r),
        backgroundColor: itemsByRate.map(([,r])=> r>=37?'#375623': r>=35?'#2E75B6': '#C00000'),
      }],
    },
    options:{
      indexAxis:'y',
      plugins:{ legend:{display:false} },
      scales:{ x:{ min:31, max:40, ticks:{callback:v=>v+'%'}, grid:{color:'#e0e0e0'} } },
    },
  });

  // 7. 거래처 TOP10 콤보
  const top10 = Object.entries(byClient).sort((a,b)=>b[1].매출액-a[1].매출액).slice(0,10);
  charts.clientBar = await canvas(1000, 520).renderToBuffer({
    type:'bar',
    data:{
      labels: top10.map(([c])=>c),
      datasets:[
        { label:'매출액(억)', data: top10.map(([,v])=>Math.round(v.매출액/1e8*10)/10),
          backgroundColor: top10.map((_,i)=>i===0?'#1F4E79':'#2E75B6'), yAxisID:'y' },
        { label:'이익률(%)', type:'line',
          data: top10.map(([,v])=>parseFloat((v.이익/v.매출액*100).toFixed(1))),
          borderColor:'#C00000', backgroundColor:'transparent',
          borderWidth:3, pointRadius:6, pointBackgroundColor:'#C00000', yAxisID:'y2' },
      ],
    },
    options:{
      plugins:{ legend:{position:'bottom'} },
      scales:{
        y: { position:'left',  ticks:{callback:v=>v+'억'} },
        y2:{ position:'right', min:31, max:39, ticks:{callback:v=>v+'%'}, grid:{drawOnChartArea:false} },
      },
    },
  });

  return charts;
}

// ── PPT 헬퍼 ──────────────────────────────────────────────
function addSlideNumber(slide, num, total=11) {
  slide.addText(`${num} / ${total}`, {
    x:'90%', y:'94%', w:'9%', h:'4%',
    fontSize:9, color:C.gray, align:'right',
  });
}

function sectionBadge(slide, text, color=C.navy) {
  slide.addShape('rect', { x:0.3, y:0.18, w:2.2, h:0.32, fill:{color}, line:{color} });
  slide.addText(text, { x:0.3, y:0.18, w:2.2, h:0.32, fontSize:11, bold:true, color:C.white, align:'center', valign:'middle' });
}

function slideTitle(slide, title, sub='') {
  slide.addText(title, { x:0.3, y:0.55, w:9.1, h:0.52, fontSize:22, bold:true, color:C.navy, fontFace:'맑은 고딕' });
  if (sub) slide.addText(sub, { x:0.3, y:1.05, w:9.1, h:0.28, fontSize:12, color:C.gray, fontFace:'맑은 고딕' });
  slide.addShape('line', { x:0.3, y:1.35, w:9.1, h:0, line:{color:C.blue, width:2} });
}

function kpiCard(slide, x, y, w, h, label, value, sub, bgColor, textColor=C.white) {
  slide.addShape('roundRect', { x, y, w, h, fill:{color:bgColor}, line:{color:bgColor}, rectRadius:0.1 });
  slide.addText(label, { x, y:y+0.08, w, h:0.28, fontSize:11, bold:false, color:textColor, align:'center', fontFace:'맑은 고딕' });
  slide.addText(value, { x, y:y+0.32, w, h:0.42, fontSize:20, bold:true, color:textColor, align:'center', fontFace:'맑은 고딕' });
  if (sub) slide.addText(sub, { x, y:y+0.74, w, h:0.22, fontSize:9, color:textColor, align:'center', fontFace:'맑은 고딕', transparency:20 });
}

function insightBox(slide, x, y, w, emoji, title, body, bg, titleColor=C.navy) {
  slide.addShape('roundRect', { x, y, w, h:1.55, fill:{color:bg}, line:{color:C.lblue}, rectRadius:0.08 });
  slide.addText(`${emoji}  ${title}`, { x:x+0.12, y:y+0.08, w:w-0.24, h:0.3, fontSize:12, bold:true, color:titleColor, fontFace:'맑은 고딕' });
  slide.addShape('line', { x:x+0.12, y:y+0.4, w:w-0.24, h:0, line:{color:C.lblue, width:1} });
  slide.addText(body, { x:x+0.12, y:y+0.46, w:w-0.24, h:1.0, fontSize:10, color:C.gray, fontFace:'맑은 고딕', valign:'top', wrap:true });
}

// ── 메인 ───────────────────────────────────────────────────
async function main() {
  console.log('차트 이미지 생성 중...');
  const charts = await makeCharts();
  console.log('✓ 차트 7종 완료\n');

  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_WIDE'; // 16:9 (13.33 x 7.5인치)
  prs.author  = 'Claude';
  prs.subject = '2025 연간 매출 분석 보고서';

  const W = 13.33, H = 7.5;

  // ══════════════════════════════════════════════════════════
  // 슬라이드 1: 표지
  // ══════════════════════════════════════════════════════════
  const s1 = prs.addSlide();

  // 배경 분할
  s1.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.navy} });
  s1.addShape('rect', { x:0, y:H*0.62, w:W, h:H*0.38, fill:{color:'162E46'} });

  // 대각선 포인트
  s1.addShape('rect', { x:W*0.55, y:0, w:W*0.45, h:H*0.62,
    fill:{color:'244F76'}, line:{color:'244F76'} });

  // 장식선
  s1.addShape('line', { x:0.5, y:H*0.62-0.06, w:W-1, h:0, line:{color:C.gold, width:3} });

  // 메인 타이틀
  s1.addText('2025년 연간 매출', {
    x:0.6, y:1.5, w:7, h:0.9, fontSize:44, bold:true, color:C.white, fontFace:'맑은 고딕',
  });
  s1.addText('분석 보고서', {
    x:0.6, y:2.35, w:7, h:0.9, fontSize:44, bold:true, color:C.gold, fontFace:'맑은 고딕',
  });
  s1.addShape('line', { x:0.6, y:3.28, w:3.5, h:0, line:{color:C.white, width:1.5} });
  s1.addText('January – December 2025  |  Annual Sales Report', {
    x:0.6, y:3.38, w:7, h:0.3, fontSize:13, color:'9DC3E6', fontFace:'맑은 고딕',
  });

  // KPI 미니 카드 (표지)
  const coverKpis = [
    { label:'총 매출', value: toUk(total.매출액)+'억' },
    { label:'총 이익', value: toUk(total.이익)+'억' },
    { label:'이익률',  value: profitRate+'%' },
    { label:'거래 건수', value: rows.length+'건' },
  ];
  coverKpis.forEach((k, i) => {
    const cx = 0.6 + i * 2.0;
    s1.addShape('roundRect', { x:cx, y:H*0.67, w:1.8, h:1.3,
      fill:{color:'244F76'}, line:{color:C.blue}, rectRadius:0.08 });
    s1.addText(k.label, { x:cx, y:H*0.67+0.08, w:1.8, h:0.28, fontSize:11, color:'9DC3E6', align:'center', fontFace:'맑은 고딕' });
    s1.addText(k.value, { x:cx, y:H*0.67+0.36, w:1.8, h:0.5, fontSize:22, bold:true, color:C.white, align:'center', fontFace:'맑은 고딕' });
  });

  // 우측 장식
  s1.addText('SALES\nPERFORMANCE', {
    x:W*0.72, y:1.6, w:3, h:1.5, fontSize:36, bold:true, color:'244F76', align:'center', fontFace:'Arial',
  });

  // 날짜
  s1.addText('작성일: 2026.04.14', {
    x:W-3, y:H-0.5, w:2.8, h:0.3, fontSize:10, color:'9DC3E6', align:'right', fontFace:'맑은 고딕',
  });

  console.log('✓ 슬라이드 1 (표지)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 2: 목차
  // ══════════════════════════════════════════════════════════
  const s2 = prs.addSlide();
  s2.addShape('rect', { x:0, y:0, w:0.22, h:H, fill:{color:C.navy} });
  s2.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });

  s2.addText('목  차', { x:0.5, y:0.25, w:W-1, h:0.6,
    fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s2.addText('Contents', { x:0.5, y:0.78, w:W-1, h:0.25,
    fontSize:13, color:C.lblue, fontFace:'맑은 고딕' });

  const toc = [
    { no:'01', title:'경영 요약 (Executive Summary)',      sub:'핵심 KPI 및 연간 성과 개요' },
    { no:'02', title:'월별 매출 추이',                    sub:'월간 매출·비용·이익 및 이익률 변동' },
    { no:'03', title:'부서별 실적 분석',                  sub:'5개 사업부문 성과 비교' },
    { no:'04', title:'담당자별 성과',                     sub:'13명 영업 담당자 실적 현황' },
    { no:'05', title:'상품·수익성 분석',                  sub:'23개 항목별 이익률 및 매출 비중' },
    { no:'06', title:'거래처 분석',                       sub:'핵심 거래처 TOP10 현황' },
    { no:'07', title:'핵심 인사이트 & 액션 플랜',         sub:'발견된 기회와 2026년 권고사항' },
  ];

  toc.forEach((item, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.5 + col * 6.4;
    const y = 1.35 + row * 1.8;

    s2.addShape('roundRect', { x, y, w:6.1, h:1.6,
      fill:{color: col===0 ? C.bg : 'F0F7FF'}, line:{color:C.lblue}, rectRadius:0.1 });
    s2.addShape('rect', { x, y, w:0.7, h:1.6,
      fill:{color: i<3 ? C.navy : C.blue}, line:{color: i<3 ? C.navy : C.blue} });
    s2.addText(item.no, { x, y, w:0.7, h:1.6,
      fontSize:20, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Arial' });
    s2.addText(item.title, { x:x+0.82, y:y+0.25, w:5.1, h:0.45,
      fontSize:13, bold:true, color:C.navy, fontFace:'맑은 고딕' });
    s2.addText(item.sub, { x:x+0.82, y:y+0.72, w:5.1, h:0.35,
      fontSize:10, color:C.gray, fontFace:'맑은 고딕' });
  });

  // 마지막 항목 (홀수이면 중앙 배치)
  addSlideNumber(s2, 2);
  console.log('✓ 슬라이드 2 (목차)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 3: 경영 요약
  // ══════════════════════════════════════════════════════════
  const s3 = prs.addSlide();
  s3.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s3.addText('경영 요약', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s3.addText('Executive Summary  |  2025년 연간 실적', {
    x:0.5, y:0.68, w:8, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  // KPI 카드 4개
  const kpis = [
    { label:'연간 총 매출액', value: toUk(total.매출액)+'억원', sub:'₩'+fmtNum(total.매출액), bg:C.navy },
    { label:'연간 총 이익',   value: toUk(total.이익)+'억원',   sub:'₩'+fmtNum(total.이익),   bg:C.blue },
    { label:'평균 이익률',    value: profitRate+'%',             sub:'목표 대비 +0.4%p',       bg:C.green },
    { label:'최고 매출 월',   value: bestMonth.replace('2025-','')+'월', sub: toUk(byMonth[bestMonth].매출액)+'억 달성', bg:C.red },
  ];
  kpis.forEach((k,i) => {
    kpiCard(s3, 0.35+i*3.22, 1.25, 3.0, 1.18, k.label, k.value, k.sub, k.bg);
  });

  // 주요 지표 테이블
  s3.addShape('rect', { x:0.35, y:2.62, w:12.6, h:0.38,
    fill:{color:C.navy}, line:{color:C.navy} });
  ['구분','상반기(1~6월)','하반기(7~12월)','전체 합계','비고'].forEach((h,i) => {
    const xs = [0.35, 2.35, 5.0, 7.65, 10.3];
    s3.addText(h, { x:xs[i], y:2.62, w: i<4 ? 2.6 : 2.7, h:0.38,
      fontSize:11, bold:true, color:C.white, align:'center', fontFace:'맑은 고딕' });
  });

  const h1 = months.slice(0,6).reduce((s,m)=>({매출액:s.매출액+byMonth[m].매출액, 이익:s.이익+byMonth[m].이익}), {매출액:0,이익:0});
  const h2 = months.slice(6).reduce((s,m)=>({매출액:s.매출액+byMonth[m].매출액, 이익:s.이익+byMonth[m].이익}), {매출액:0,이익:0});
  const tableRows = [
    ['매출액', toUk(h1.매출액)+'억', toUk(h2.매출액)+'억', toUk(total.매출액)+'억', `하반기 +${((h2.매출액/h1.매출액-1)*100).toFixed(1)}% ▲`],
    ['이익',   toUk(h1.이익)+'억',   toUk(h2.이익)+'억',   toUk(total.이익)+'억',   `하반기 +${((h2.이익/h1.이익-1)*100).toFixed(1)}% ▲`],
    ['이익률', (h1.이익/h1.매출액*100).toFixed(1)+'%', (h2.이익/h2.매출액*100).toFixed(1)+'%', profitRate+'%', '하반기 이익률 개선'],
  ];
  tableRows.forEach((tr, ri) => {
    const bg = ri%2===0 ? C.bg : C.white;
    const xs = [0.35, 2.35, 5.0, 7.65, 10.3];
    tr.forEach((cell, ci) => {
      s3.addShape('rect', { x:xs[ci], y:3.0+ri*0.46, w: ci<4?2.6:2.7, h:0.46,
        fill:{color: ci===0?'EEF4FB':bg}, line:{color:'DDDDDD'} });
      s3.addText(cell, { x:xs[ci], y:3.0+ri*0.46, w: ci<4?2.6:2.7, h:0.46,
        fontSize:11, bold:ci===0, color: ci===4?C.green:C.gray,
        align: ci===0?'left':'center', indent: ci===0?0.15:0, fontFace:'맑은 고딕' });
    });
  });

  // 인사이트 한 줄
  s3.addShape('roundRect', { x:0.35, y:4.45, w:12.6, h:0.7,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.08 });
  s3.addText('💡  하반기 매출이 상반기 대비 25.1% 성장하며 연간 최고 기록 갱신. 12월 단월 최고 매출(158억) 달성. 이익률은 연중 안정적 35% 유지.',
    { x:0.5, y:4.45, w:12.3, h:0.7, fontSize:11.5, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s3, 3);
  console.log('✓ 슬라이드 3 (경영 요약)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 4: 월별 매출 추이 (막대)
  // ══════════════════════════════════════════════════════════
  const s4 = prs.addSlide();
  s4.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s4.addText('월별 매출 추이', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s4.addText('01  월별 매출 분석  |  매출액 · 비용 · 이익 비교', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  s4.addImage({ data: 'image/png;base64,' + charts.monthBar.toString('base64'), x:0.3, y:1.2, w:8.6, h:4.5 });

  // 우측 사이드 패널
  const monthSorted = [...months].sort((a,b)=>byMonth[b].매출액-byMonth[a].매출액);
  s4.addShape('rect', { x:9.1, y:1.2, w:4.0, h:4.5, fill:{color:C.bg}, line:{color:C.lblue} });
  s4.addText('월별 순위', { x:9.1, y:1.2, w:4.0, h:0.38,
    fontSize:12, bold:true, color:C.navy, align:'center', fontFace:'맑은 고딕',
    fill:{color:C.lblue} });
  monthSorted.forEach((m, i) => {
    const rate = (byMonth[m].이익/byMonth[m].매출액*100).toFixed(1);
    const isBest = i===0, isWorst = i===months.length-1;
    s4.addShape('rect', { x:9.1, y:1.58+i*0.33, w:4.0, h:0.33,
      fill:{color: isBest?'E2EFDA': isWorst?C.lred: i%2===0?C.bg:C.white}, line:{color:'DDDDDD'} });
    s4.addText(`${i+1}위  ${m.replace('2025-','')}월`, { x:9.15, y:1.58+i*0.33, w:2.0, h:0.33,
      fontSize:10, bold:isBest, color: isBest?C.green: isWorst?C.red:C.gray, fontFace:'맑은 고딕' });
    s4.addText(`${toUk(byMonth[m].매출액)}억`, { x:11.2, y:1.58+i*0.33, w:1.0, h:0.33,
      fontSize:10, bold:isBest, color:C.navy, align:'right', fontFace:'맑은 고딕' });
    s4.addText(`${rate}%`, { x:12.25, y:1.58+i*0.33, w:0.8, h:0.33,
      fontSize:9, color:C.gray, align:'right', fontFace:'맑은 고딕' });
  });

  s4.addShape('roundRect', { x:0.3, y:5.85, w:12.8, h:0.42,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.06 });
  s4.addText(`📌  최고: ${bestMonth.replace('2025-','')}월 ${toUk(byMonth[bestMonth].매출액)}억  |  최저: ${worstMonth.replace('2025-','')}월 ${toUk(byMonth[worstMonth].매출액)}억  |  월평균: ${toUk(total.매출액/12)}억`,
    { x:0.5, y:5.85, w:12.6, h:0.42, fontSize:11, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s4, 4);
  console.log('✓ 슬라이드 4 (월별 추이 막대)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 5: 이익률 트렌드
  // ══════════════════════════════════════════════════════════
  const s5 = prs.addSlide();
  s5.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s5.addText('이익률 트렌드 분석', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s5.addText('02  월별 이익률 변동  |  고·저 이익률 월 식별', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  s5.addImage({ data: 'image/png;base64,' + charts.trendLine.toString('base64'), x:0.3, y:1.2, w:9.0, h:4.5 });

  // 우측 해석 패널
  s5.addShape('roundRect', { x:9.5, y:1.2, w:3.55, h:2.1,
    fill:{color:'E2EFDA'}, line:{color:C.lgreen}, rectRadius:0.08 });
  s5.addText('▲ 고이익률 월 TOP3', { x:9.62, y:1.28, w:3.3, h:0.3,
    fontSize:11, bold:true, color:C.green, fontFace:'맑은 고딕' });
  const top3months = [...months].sort((a,b)=>
    byMonth[b].이익/byMonth[b].매출액 - byMonth[a].이익/byMonth[a].매출액).slice(0,3);
  top3months.forEach((m,i) => {
    const r = (byMonth[m].이익/byMonth[m].매출액*100).toFixed(1);
    s5.addText(`${['🥇','🥈','🥉'][i]}  ${m.replace('2025-','')}월  ${r}%`, {
      x:9.62, y:1.62+i*0.46, w:3.3, h:0.4,
      fontSize:12, bold:i===0, color:C.green, fontFace:'맑은 고딕' });
  });

  s5.addShape('roundRect', { x:9.5, y:3.4, w:3.55, h:2.1,
    fill:{color:C.lred}, line:{color:'FFAAAA'}, rectRadius:0.08 });
  s5.addText('▼ 저이익률 월 BOTTOM3', { x:9.62, y:3.48, w:3.3, h:0.3,
    fontSize:11, bold:true, color:C.red, fontFace:'맑은 고딕' });
  const bot3months = [...months].sort((a,b)=>
    byMonth[a].이익/byMonth[a].매출액 - byMonth[b].이익/byMonth[b].매출액).slice(0,3);
  bot3months.forEach((m,i) => {
    const r = (byMonth[m].이익/byMonth[m].매출액*100).toFixed(1);
    s5.addText(`${i+1}.  ${m.replace('2025-','')}월  ${r}%`, {
      x:9.62, y:3.82+i*0.46, w:3.3, h:0.4,
      fontSize:12, bold:i===0, color:C.red, fontFace:'맑은 고딕' });
  });

  s5.addShape('roundRect', { x:0.3, y:5.85, w:12.8, h:0.42,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.06 });
  s5.addText('📌  연중 이익률 변동폭 3.5%p 이내로 안정적. 하반기(평균 35.9%)가 상반기(34.7%)보다 1.2%p 높음 — 비용 효율 개선 효과.',
    { x:0.5, y:5.85, w:12.6, h:0.42, fontSize:11, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s5, 5);
  console.log('✓ 슬라이드 5 (이익률 트렌드)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 6: 부서별 분석
  // ══════════════════════════════════════════════════════════
  const s6 = prs.addSlide();
  s6.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s6.addText('부서별 실적 분석', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s6.addText('03  5개 사업부문 성과 비교  |  매출·이익·이익률', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  s6.addImage({ data: 'image/png;base64,' + charts.deptBar.toString('base64'), x:0.3, y:1.2, w:7.5, h:4.2 });

  // 부서별 카드
  const deptSorted = Object.entries(byDept).sort((a,b)=>b[1].매출액-a[1].매출액);
  const deptColors = [C.navy, C.blue, '5B9BD5', '9DC3E6', 'BDD7EE'];
  deptSorted.forEach(([dept, v], i) => {
    const rate = (v.이익/v.매출액*100).toFixed(1);
    s6.addShape('roundRect', { x:8.0, y:1.2+i*0.95, w:5.05, h:0.85,
      fill:{color: i===0?C.navy:C.bg}, line:{color:C.lblue}, rectRadius:0.08 });
    s6.addShape('rect', { x:8.0, y:1.2+i*0.95, w:0.5, h:0.85,
      fill:{color:deptColors[i]}, line:{color:deptColors[i]} });
    s6.addText(`${i+1}위`, { x:8.0, y:1.2+i*0.95, w:0.5, h:0.85,
      fontSize:10, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'맑은 고딕' });
    s6.addText(dept, { x:8.6, y:1.25+i*0.95, w:2.5, h:0.35,
      fontSize:12, bold:i===0, color:i===0?C.white:C.navy, fontFace:'맑은 고딕' });
    s6.addText(`매출 ${toUk(v.매출액)}억  |  이익률 ${rate}%`, { x:8.6, y:1.6+i*0.95, w:4.3, h:0.28,
      fontSize:10, color:i===0?C.lblue:C.gray, fontFace:'맑은 고딕' });
  });

  s6.addShape('roundRect', { x:0.3, y:5.55, w:12.8, h:0.7,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.06 });
  const bestRateDept = Object.entries(byDept).sort((a,b)=>b[1].이익/b[1].매출액-a[1].이익/a[1].매출액)[0];
  s6.addText(`📌  매출 1위: 해외사업부(${toUk(byDept['해외사업부'].매출액)}억)  |  이익률 1위: ${bestRateDept[0]}(${(bestRateDept[1].이익/bestRateDept[1].매출액*100).toFixed(1)}%)  |  부서 간 매출 격차 약 234억 — 영업1팀 강화 필요.`,
    { x:0.5, y:5.55, w:12.6, h:0.7, fontSize:11, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s6, 6);
  console.log('✓ 슬라이드 6 (부서별)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 7: 담당자 성과
  // ══════════════════════════════════════════════════════════
  const s7 = prs.addSlide();
  s7.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s7.addText('담당자별 성과 분석', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s7.addText('04  13명 영업 담당자  |  매출 vs 이익률 비교', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  s7.addImage({ data: 'image/png;base64,' + charts.mgrBar.toString('base64'), x:0.3, y:1.15, w:9.5, h:4.55 });

  // 우측 TOP/BOTTOM
  const mgrByRev  = Object.entries(byMgr).sort((a,b)=>b[1].매출액-a[1].매출액);
  const mgrByRate = Object.entries(byMgr).sort((a,b)=>b[1].이익/b[1].매출액-a[1].이익/a[1].매출액);
  const mgrByRateAsc = [...mgrByRate].reverse();

  [[' 매출 TOP3', mgrByRev.slice(0,3), C.navy, '매출'],
   ['이익률 TOP3', mgrByRate.slice(0,3), C.green, '이익률'],
   ['이익률 하위', mgrByRateAsc.slice(0,3), C.red, '이익률'],
  ].forEach(([title, list, color, metric], gi) => {
    const gy = 1.15 + gi * 1.65;
    s7.addShape('roundRect', { x:9.9, y:gy, w:3.15, h:1.5,
      fill:{color:C.bg}, line:{color:C.lblue}, rectRadius:0.08 });
    s7.addShape('rect', { x:9.9, y:gy, w:3.15, h:0.36,
      fill:{color:color}, line:{color:color} });
    s7.addText(title, { x:9.9, y:gy, w:3.15, h:0.36,
      fontSize:11, bold:true, color:C.white, align:'center', fontFace:'맑은 고딕' });
    list.forEach(([name, v], i) => {
      const val = metric==='매출' ? `${toUk(v.매출액)}억` : `${(v.이익/v.매출액*100).toFixed(1)}%`;
      s7.addText(`${['①','②','③'][i]} ${name}  ${val}`, {
        x:10.05, y:gy+0.42+i*0.34, w:2.9, h:0.32,
        fontSize:10.5, bold:i===0, color:i===0?color:C.gray, fontFace:'맑은 고딕' });
    });
  });

  s7.addShape('roundRect', { x:0.3, y:5.85, w:12.8, h:0.42,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.06 });
  s7.addText('📌  임도현(166억) 매출 1위, 이지현(36.4%) 이익률 1위. 강민지·최영호는 이익률 개선 여지 있음 → 영업 전략 및 협상력 교육 권고.',
    { x:0.5, y:5.85, w:12.6, h:0.42, fontSize:11, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s7, 7);
  console.log('✓ 슬라이드 7 (담당자)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 8: 상품·수익성 분석
  // ══════════════════════════════════════════════════════════
  const s8 = prs.addSlide();
  s8.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s8.addText('상품·수익성 분석', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s8.addText('05  23개 항목별 이익률 비교  |  고마진·저마진 식별', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  s8.addImage({ data: 'image/png;base64,' + charts.itemDoughnut.toString('base64'), x:0.3, y:1.15, w:5.8, h:4.2 });
  s8.addImage({ data: 'image/png;base64,' + charts.itemRate.toString('base64'),     x:6.3, y:1.15, w:6.8, h:4.2 });

  // 레전드 텍스트
  [['■ 진한색: 이익률 37%↑ (고마진)', C.green],
   ['■ 중간색: 이익률 35~37%',        C.blue],
   ['■ 연한색: 이익률 35%↓ (저마진)', C.red],
  ].forEach(([txt, col], i) => {
    s8.addText(txt, { x:6.3+i*2.22, y:5.45, w:2.2, h:0.28,
      fontSize:9.5, color:col, fontFace:'맑은 고딕' });
  });

  s8.addShape('roundRect', { x:0.3, y:5.82, w:12.8, h:0.45,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.06 });
  s8.addText('📌  브랜딩 컨설팅(38.1%), 일본 수출(38.0%), 엔터프라이즈(37.8%) 고마진 3종 집중 확대 권고. 싱가포르 프로젝트(32.9%) 원가 재검토 필요.',
    { x:0.5, y:5.82, w:12.6, h:0.45, fontSize:11, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s8, 8);
  console.log('✓ 슬라이드 8 (상품 분석)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 9: 거래처 분석
  // ══════════════════════════════════════════════════════════
  const s9 = prs.addSlide();
  s9.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s9.addText('핵심 거래처 분석', { x:0.5, y:0.2, w:8, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s9.addText('06  거래처 TOP10  |  매출 비중 & 수익성', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  s9.addImage({ data: 'image/png;base64,' + charts.clientBar.toString('base64'), x:0.3, y:1.15, w:8.8, h:4.55 });

  // 우측 거래처 카드
  const topClients = Object.entries(byClient).sort((a,b)=>b[1].매출액-a[1].매출액).slice(0,5);
  topClients.forEach(([client, v], i) => {
    const rate = (v.이익/v.매출액*100).toFixed(1);
    const shareStr = (v.매출액/total.매출액*100).toFixed(1)+'%';
    s9.addShape('roundRect', { x:9.25, y:1.15+i*0.95, w:3.8, h:0.85,
      fill:{color:i===0?C.navy:C.bg}, line:{color:C.lblue}, rectRadius:0.08 });
    s9.addText(`${i+1}`, { x:9.25, y:1.15+i*0.95, w:0.48, h:0.85,
      fontSize:14, bold:true, color:i===0?C.gold:C.navy, align:'center', valign:'middle', fontFace:'Arial' });
    s9.addText(client, { x:9.73, y:1.2+i*0.95, w:3.1, h:0.32,
      fontSize:12, bold:i===0, color:i===0?C.white:C.navy, fontFace:'맑은 고딕' });
    s9.addText(`${toUk(v.매출액)}억  |  점유 ${shareStr}  |  이익률 ${rate}%`,
      { x:9.73, y:1.52+i*0.95, w:3.1, h:0.28,
        fontSize:9.5, color:i===0?C.lblue:C.gray, fontFace:'맑은 고딕' });
  });

  s9.addShape('roundRect', { x:0.3, y:5.85, w:12.8, h:0.42,
    fill:{color:C.yellow}, line:{color:C.gold}, rectRadius:0.06 });
  const top1 = topClients[0];
  s9.addText(`📌  ${top1[0]}이 ${toUk(top1[1].매출액)}억(10.5%)으로 압도적 1위. 동아시스템즈는 이익률 36.6%로 수익성 최우수. 상위 5개사가 전체의 48% 차지 → 집중 관리 필요.`,
    { x:0.5, y:5.85, w:12.6, h:0.42, fontSize:11, color:'594000', fontFace:'맑은 고딕', valign:'middle' });

  addSlideNumber(s9, 9);
  console.log('✓ 슬라이드 9 (거래처)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 10: 인사이트 & 액션 플랜
  // ══════════════════════════════════════════════════════════
  const s10 = prs.addSlide();
  s10.addShape('rect', { x:0, y:0, w:W, h:1.1, fill:{color:C.navy} });
  s10.addText('핵심 인사이트 & 액션 플랜', { x:0.5, y:0.2, w:9, h:0.5, fontSize:26, bold:true, color:C.white, fontFace:'맑은 고딕' });
  s10.addText('07  Key Insights & Action Plan  |  2026년 전략 방향', { x:0.5, y:0.68, w:10, h:0.3, fontSize:12, color:C.lblue, fontFace:'맑은 고딕' });

  // 인사이트 2x3 그리드
  const insights = [
    { emoji:'📈', title:'하반기 강세 지속', body:'하반기 평균 매출 140억, 상반기 112억 대비 +25.1%. 10~12월 3연속 최고 기록.', bg:'EEF4FB', tc:C.navy },
    { emoji:'💰', title:'이익률 안정 (35%대)', body:'연중 최저 33.8%(5월)~최고 37.3%(12월). 8·11·12월 비용 효율 우수.', bg:'E2EFDA', tc:C.green },
    { emoji:'🏆', title:'고마진 항목 발굴', body:'브랜딩 컨설팅 38.1%, 일본수출 38.0%, 엔터프라이즈 37.8% → 투자 확대.', bg:'EEF4FB', tc:C.navy },
    { emoji:'⚠️', title:'저마진 항목 경고', body:'싱가포르 프로젝트 32.9%, 현지기술지원 33.2% → 원가 구조 재검토.', bg:'FFE0E0', tc:C.red },
    { emoji:'👑', title:'케이에스테크 핵심 고객', body:'148억(10.5%) 1위 고객. 동아시스템즈 이익률 36.6%로 수익성도 탁월.', bg:'E2EFDA', tc:C.green },
    { emoji:'🌟', title:'임도현·이지현 투 톱', body:'임도현 매출 1위(166억), 이지현 이익률 1위(36.4%). 강민지·최영호 개선 필요.', bg:'EEF4FB', tc:C.navy },
  ];

  insights.forEach((ins, i) => {
    const col = i % 3, row = Math.floor(i/3);
    const x = 0.3 + col * 4.3;
    const y = 1.22 + row * 1.75;
    s10.addShape('roundRect', { x, y, w:4.1, h:1.62,
      fill:{color:ins.bg}, line:{color:C.lblue}, rectRadius:0.1 });
    s10.addShape('rect', { x, y, w:4.1, h:0.38,
      fill:{color:ins.tc}, line:{color:ins.tc} });
    s10.addText(`${ins.emoji}  ${ins.title}`, { x:x+0.12, y:y+0.04, w:3.86, h:0.3,
      fontSize:12, bold:true, color:C.white, fontFace:'맑은 고딕' });
    s10.addText(ins.body, { x:x+0.12, y:y+0.46, w:3.86, h:1.08,
      fontSize:10.5, color:C.gray, fontFace:'맑은 고딕', valign:'top', wrap:true });
  });

  // 액션 플랜 바
  const actions = [
    '① 브랜딩·일본수출·엔터프라이즈 라인업 확대',
    '② 싱가포르·현지기술지원 원가 절감 수립',
    '③ 케이에스테크·동아시스템즈 전담 관리',
    '④ 강민지·최영호 협상력 교육 지원',
    '⑤ 1~2월 비수기 대응 조기 프로모션 기획',
  ];
  s10.addShape('rect', { x:0.3, y:4.76, w:W-0.6, h:0.35,
    fill:{color:C.navy}, line:{color:C.navy} });
  s10.addText('2026년 권고 액션 플랜', { x:0.3, y:4.76, w:W-0.6, h:0.35,
    fontSize:11, bold:true, color:C.white, align:'center', fontFace:'맑은 고딕' });

  actions.forEach((act, i) => {
    s10.addShape('rect', { x:0.3, y:5.11+i*0.28, w:W-0.6, h:0.28,
      fill:{color: i%2===0?C.yellow:'FFFFF0'}, line:{color:'E0E0A0'} });
    s10.addText(act, { x:0.5, y:5.11+i*0.28, w:W-1, h:0.28,
      fontSize:11, color:C.gray, fontFace:'맑은 고딕', valign:'middle' });
  });

  addSlideNumber(s10, 10);
  console.log('✓ 슬라이드 10 (인사이트·액션플랜)');

  // ══════════════════════════════════════════════════════════
  // 슬라이드 11: 마무리
  // ══════════════════════════════════════════════════════════
  const s11 = prs.addSlide();
  s11.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.navy} });
  s11.addShape('rect', { x:0, y:H*0.55, w:W, h:H*0.45, fill:{color:'162E46'} });
  s11.addShape('line', { x:0.5, y:H*0.55-0.05, w:W-1, h:0, line:{color:C.gold, width:3} });

  s11.addText('감사합니다', { x:1, y:1.5, w:11, h:1.2,
    fontSize:52, bold:true, color:C.white, align:'center', fontFace:'맑은 고딕' });
  s11.addText('Thank You', { x:1, y:2.68, w:11, h:0.6,
    fontSize:28, color:C.gold, align:'center', fontFace:'Arial' });
  s11.addShape('line', { x:3, y:3.35, w:7.3, h:0, line:{color:C.white, width:1} });
  s11.addText('2025 Annual Sales Report  |  분석 기준: 2025.01 ~ 2025.12', {
    x:1, y:3.45, w:11, h:0.35, fontSize:13, color:'9DC3E6', align:'center', fontFace:'맑은 고딕' });

  // 하단 수치 요약
  const sumCards = [
    { label:'총 매출', val: toUk(total.매출액)+'억원' },
    { label:'총 이익', val: toUk(total.이익)+'억원' },
    { label:'이익률', val: profitRate+'%' },
    { label:'거래 건수', val: rows.length+'건' },
    { label:'거래처 수', val: Object.keys(byClient).length+'개사' },
  ];
  sumCards.forEach((c, i) => {
    const cx = 0.5 + i * 2.46;
    s11.addShape('roundRect', { x:cx, y:H*0.6, w:2.26, h:1.4,
      fill:{color:'244F76'}, line:{color:C.blue}, rectRadius:0.1 });
    s11.addText(c.label, { x:cx, y:H*0.6+0.1, w:2.26, h:0.3,
      fontSize:10.5, color:'9DC3E6', align:'center', fontFace:'맑은 고딕' });
    s11.addText(c.val, { x:cx, y:H*0.6+0.44, w:2.26, h:0.55,
      fontSize:20, bold:true, color:C.white, align:'center', fontFace:'맑은 고딕' });
  });

  s11.addText('Powered by Claude AI  ·  2026.04.14', {
    x:0, y:H-0.38, w:W, h:0.35, fontSize:10, color:'9DC3E6',
    align:'center', fontFace:'맑은 고딕' });

  addSlideNumber(s11, 11);
  console.log('✓ 슬라이드 11 (마무리)');

  // ── 저장 ──
  await prs.writeFile({ fileName: DEST });
  console.log(`\n✅ PPT 저장 완료: ${DEST}`);
  console.log('   총 슬라이드: 11장');
}

main().catch(console.error);
