# 인건비 엑셀 기간별 급여유형 지원 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `build_payroll.py`의 `EMPLOYEES` 상수와 세 시트 빌드 함수를 수정하여, 고경민처럼 시점에 따라 월급/시급 유형이 바뀐 직원의 이력을 기간별로 저장하고 급여 요약이 선택월에 맞는 유형·금액을 표시하도록 개선한다.

**Architecture:** 직원 마스터를 한 직원 여러 행 구조로 확장(F/G 컬럼을 "적용시작월/적용종료월"로 의미 재정의). 근무 기록 G열 수식을 기간 기반 `INDEX/MATCH` 배열 수식으로 교체. 급여 요약은 기존 `active_check` 로직을 그대로 활용하면서 루프 범위만 마스터 행 수에 맞게 확장.

**Tech Stack:** Python 3, openpyxl(.xlsx 생성), pyxlsb(원본 .xlsb 파싱 — 근무기록 추출용 기존 로직 유지)

**Spec:** `docs/superpowers/specs/2026-04-19-payroll-period-based-paytype.md`

**대상 파일 (단일):**
- `BUNT_finance/build_payroll.py` (수정)
- `BUNT_finance/분트_인건비관리_v2.xlsx` (재생성 — 커밋 대상)

---

## Task 1: EMPLOYEES 상수를 구간 기반 리스트로 재작성

**Files:**
- Modify: `BUNT_finance/build_payroll.py:22-36`

**의도:** 기존 한 직원 한 dict 구조를 "구간별 한 dict" 구조로 바꾼다. 마스터 시트에 쓰일 행의 순서 그대로, 고경민은 3개 dict로 분해.

- [ ] **Step 1: EMPLOYEES 리스트 교체**

현재 코드 ([BUNT_finance/build_payroll.py:22-36](BUNT_finance/build_payroll.py#L22-L36)):

```python
EMPLOYEES = [
    {"name": "고경민", "pay_type": "월급", "amount": 1300000, "account": "국민은행 82240104164295",
     "status": "재직", "start_month": 202502, "end_month": None},
    {"name": "장예원", "pay_type": "시급", "amount": 12000, "account": "카카오뱅크 3333-07-7072641",
     "status": "퇴직", "start_month": 202502, "end_month": 202507},
    {"name": "유은비", "pay_type": "시급", "amount": 12000, "account": "우리 1002-533-898534",
     "status": "퇴직", "start_month": 202502, "end_month": 202503},
    {"name": "전민아", "pay_type": "시급", "amount": 12000, "account": "하나은행 558-910330-30707",
     "status": "재직", "start_month": 202503, "end_month": None},
    {"name": "김채현", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-351-542250",
     "status": "퇴직", "start_month": 202507, "end_month": 202510},
    {"name": "이진화", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-166-678383",
     "status": "재직", "start_month": 202511, "end_month": None},
]
```

신규 코드 — 필드명은 그대로 두고(수식 호환성), `start_month/end_month`의 의미만 "적용시작월/적용종료월"로 변경:

```python
# 직원 급여 이력 (한 직원 여러 행 가능, 각 행은 하나의 급여 구간)
# start_month ~ end_month는 해당 급여유형/금액이 적용되는 기간 (YYYYMM)
# 같은 이름의 여러 행은 기간이 서로 겹치지 않아야 함 (불변 조건)
EMPLOYEES = [
    # 고경민: 월급(1,2월) → 시급(3~9월) → 월급(10월~)
    {"name": "고경민", "pay_type": "월급", "amount": 1200000, "account": "국민은행 82240104164295",
     "status": "재직", "start_month": 202501, "end_month": 202502},
    {"name": "고경민", "pay_type": "시급", "amount": 12000,   "account": "국민은행 82240104164295",
     "status": "재직", "start_month": 202503, "end_month": 202509},
    {"name": "고경민", "pay_type": "월급", "amount": 1300000, "account": "국민은행 82240104164295",
     "status": "재직", "start_month": 202510, "end_month": None},
    {"name": "장예원", "pay_type": "시급", "amount": 12000, "account": "카카오뱅크 3333-07-7072641",
     "status": "퇴직", "start_month": 202502, "end_month": 202507},
    {"name": "유은비", "pay_type": "시급", "amount": 12000, "account": "우리 1002-533-898534",
     "status": "퇴직", "start_month": 202502, "end_month": 202503},
    {"name": "전민아", "pay_type": "시급", "amount": 12000, "account": "하나은행 558-910330-30707",
     "status": "재직", "start_month": 202503, "end_month": None},
    {"name": "김채현", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-351-542250",
     "status": "퇴직", "start_month": 202507, "end_month": 202510},
    {"name": "이진화", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-166-678383",
     "status": "재직", "start_month": 202511, "end_month": None},
]
```

- [ ] **Step 2: 구조 검증**

Run:
```bash
cd BUNT_finance && python -c "
from build_payroll import EMPLOYEES
assert len(EMPLOYEES) == 8, f'Expected 8 rows, got {len(EMPLOYEES)}'
kgm = [e for e in EMPLOYEES if e['name']=='고경민']
assert len(kgm) == 3
assert kgm[0] == {'name':'고경민','pay_type':'월급','amount':1200000,'account':'국민은행 82240104164295','status':'재직','start_month':202501,'end_month':202502}
assert kgm[2]['end_month'] is None
print('OK: EMPLOYEES structure valid, 8 rows, 고경민 3개 구간')
"
```

Expected: `OK: EMPLOYEES structure valid, 8 rows, 고경민 3개 구간`

- [ ] **Step 3: 커밋**

```bash
cd "c:/Users/dhchd/work"
git add BUNT_finance/build_payroll.py
git commit -m "$(cat <<'EOF'
refactor: EMPLOYEES 상수를 구간 기반 리스트로 재작성

- 직원 마스터 한 직원 여러 행 허용
- 고경민 3구간 등록 (월급 120만→시급 12,000→월급 130만)
- start_month/end_month 의미: 입사/퇴사 → 급여유형 적용시작/종료월

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: 직원 마스터 시트 헤더 및 드롭다운 범위 조정

**Files:**
- Modify: `BUNT_finance/build_payroll.py:120-160`

**의도:** 헤더 텍스트 "입사월/퇴사월" → "적용시작월/적용종료월"로 변경. 드롭다운 범위(현재 `+20`)는 그대로 두되, 다중 행 구조 지원을 위해 마스터 데이터 영역을 `len(EMPLOYEES)`(8행) 기준으로 유지.

- [ ] **Step 1: 헤더 텍스트 변경**

`build_master_sheet` 함수 내 `headers` 리스트 (라인 126):

변경 전:
```python
headers = ["이름", "급여유형", "금액", "입금계좌", "재직상태", "입사월", "퇴사월", "비고"]
```

변경 후:
```python
headers = ["이름", "급여유형", "금액", "입금계좌", "재직상태", "적용시작월", "적용종료월", "비고"]
```

- [ ] **Step 2: 생성 검증**

Run:
```bash
cd BUNT_finance && python build_payroll.py
```

Expected 출력 하단:
```
  → [직원 마스터] 시트 생성 완료
  → [근무 기록] 시트 생성 완료
  → [급여 요약] 시트 생성 완료
[완료] 저장: .../분트_인건비관리.xlsx
```

(주의: 현재 스크립트의 `OUTPUT_PATH`는 `분트_인건비관리.xlsx` — 기존 `_v2` 파일과 이름이 다름. Task 5에서 일치시킨다.)

- [ ] **Step 3: 마스터 시트 내용 검증**

Run:
```bash
cd BUNT_finance && python -c "
from openpyxl import load_workbook
wb = load_workbook('분트_인건비관리.xlsx')
ws = wb['직원 마스터']
assert ws.cell(1,6).value == '적용시작월'
assert ws.cell(1,7).value == '적용종료월'
# 8개 직원 행 (2~9행)
assert ws.cell(2,1).value == '고경민'
assert ws.cell(3,1).value == '고경민'
assert ws.cell(4,1).value == '고경민'
assert ws.cell(2,2).value == '월급' and ws.cell(2,3).value == 1200000
assert ws.cell(3,2).value == '시급' and ws.cell(3,3).value == 12000
assert ws.cell(4,2).value == '월급' and ws.cell(4,3).value == 1300000
assert ws.cell(4,7).value is None  # 현재 활성 = 적용종료월 공란
assert ws.cell(5,1).value == '장예원'
assert ws.cell(9,1).value == '이진화'
print('OK: 마스터 시트 8행, 고경민 3행, 헤더 변경됨')
"
```

Expected: `OK: 마스터 시트 8행, 고경민 3행, 헤더 변경됨`

- [ ] **Step 4: 커밋**

```bash
cd "c:/Users/dhchd/work"
git add BUNT_finance/build_payroll.py
git commit -m "$(cat <<'EOF'
refactor: 마스터 시트 헤더를 적용시작월/적용종료월로 변경

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: 근무 기록 G열(급여유형) 수식을 기간 기반 조회로 교체

**Files:**
- Modify: `BUNT_finance/build_payroll.py:206-208` (마이그레이션 데이터 G열)
- Modify: `BUNT_finance/build_payroll.py:231-233` (빈 입력행 G열)
- Modify: `BUNT_finance/build_payroll.py:227-230` (빈 입력행 F열 — 시급 조회도 기간 기반으로)

**의도:** 현재 `VLOOKUP('직원 마스터'!A:B, 2, FALSE)`는 고경민처럼 여러 행이 있으면 **첫 번째 매칭 행의 급여유형**만 반환 → 과거/현재 유형을 구분 못함. 근무기록 날짜 기준으로 **해당 시점의 활성 구간**을 찾아야 한다.

근무기록 시트의 날짜 컬럼은 A열, 근무자는 B열. 직원 마스터의 F/G는 YYYYMM 정수. 날짜를 YYYYMM으로 변환 후 매칭.

- [ ] **Step 1: 기간 기반 조회 수식 헬퍼 함수 추가**

파일 상단 `apply_cell_style` 함수(라인 111) 바로 아래에 헬퍼 함수를 추가한다.

```python
def master_lookup_formula(name_ref, date_ref, return_col, master_last_row=30):
    """
    직원 마스터에서 특정 시점에 활성인 구간을 찾아 지정 컬럼 값을 반환하는 수식.
    - name_ref: 근무자 이름 셀 참조 (예: "B5")
    - date_ref: 날짜 셀 참조 (예: "A5")
    - return_col: 반환할 마스터 컬럼 문자 (예: "B" = 급여유형, "C" = 금액)
    - master_last_row: 마스터 조회 범위 끝 행 (여유 있게 30)

    매칭 조건: 이름 일치 AND 적용시작월 <= 근무월 AND (적용종료월 공란 또는 >= 근무월)
    근무월 = YEAR*100 + MONTH
    """
    ym = f'(YEAR({date_ref})*100+MONTH({date_ref}))'
    master_a = f"'직원 마스터'!A$2:A${master_last_row}"
    master_f = f"'직원 마스터'!F$2:F${master_last_row}"
    master_g = f"'직원 마스터'!G$2:G${master_last_row}"
    master_ret = f"'직원 마스터'!{return_col}$2:{return_col}${master_last_row}"
    match_expr = (
        f'MATCH(1,'
        f'({master_a}={name_ref})*'
        f'({master_f}<={ym})*'
        f'(({master_g}="")+({master_g}>={ym})),'
        f'0)'
    )
    return f'INDEX({master_ret},{match_expr})'
```

- [ ] **Step 2: 마이그레이션 데이터 G열 수식 교체**

현재 코드 (라인 206-208):
```python
# 급여유형: 수식
ws.cell(row=r, column=7).value = (
    f'=IF(B{r}<>"",VLOOKUP(B{r},\'직원 마스터\'!A:B,2,FALSE),"")'
)
```

변경:
```python
# 급여유형: 날짜 기반 기간 조회
lookup = master_lookup_formula(f"B{r}", f"A{r}", "B")
ws.cell(row=r, column=7).value = f'=IFERROR(IF(B{r}<>"",{lookup},""),"")'
```

- [ ] **Step 3: 빈 입력행 G열 수식 교체**

현재 코드 (라인 231-233):
```python
ws.cell(row=r, column=7).value = (
    f'=IF(B{r}<>"",VLOOKUP(B{r},\'직원 마스터\'!A:B,2,FALSE),"")'
)
```

변경:
```python
lookup_type = master_lookup_formula(f"B{r}", f"A{r}", "B")
ws.cell(row=r, column=7).value = f'=IFERROR(IF(B{r}<>"",{lookup_type},""),"")'
```

- [ ] **Step 4: 빈 입력행 F열(일급여) 수식 교체**

현재 코드 (라인 227-230):
```python
ws.cell(row=r, column=6).value = (
    f'=IF(AND(B{r}<>"",G{r}="시급"),E{r}*24*VLOOKUP(B{r},\'직원 마스터\'!A:C,3,FALSE),0)'
)
```

변경 — 시급 금액도 기간 기반으로 조회:
```python
lookup_amount = master_lookup_formula(f"B{r}", f"A{r}", "C")
ws.cell(row=r, column=6).value = (
    f'=IFERROR(IF(AND(B{r}<>"",G{r}="시급"),E{r}*24*{lookup_amount},0),0)'
)
```

- [ ] **Step 5: 재생성 + 수식 검증**

Run:
```bash
cd BUNT_finance && python build_payroll.py
```

Expected: 정상 생성.

Run (생성된 파일에서 근무기록 G열 수식 확인):
```bash
cd BUNT_finance && python -c "
from openpyxl import load_workbook
wb = load_workbook('분트_인건비관리.xlsx')
ws = wb['근무 기록']
# 첫 데이터 행의 G열 수식 확인
g2 = ws.cell(2,7).value
print('G2 수식:', g2[:80], '...' if len(g2)>80 else '')
assert 'INDEX' in g2 and 'MATCH' in g2, 'INDEX/MATCH 수식 필요'
assert 'YEAR(A2)' in g2, '날짜 기반 조회 필요'
print('OK: G열 수식이 기간 기반으로 교체됨')
"
```

Expected: `OK: G열 수식이 기간 기반으로 교체됨`

- [ ] **Step 6: Excel에서 실제 계산 확인 (수동)**

생성된 `분트_인건비관리.xlsx`를 Excel에서 열고:
1. 근무기록 탭에서 고경민 2025-04-XX 행의 G열이 **"시급"**으로 계산되는지 확인
2. 고경민 2025-02-XX 행이 있다면 G열이 **"월급"**으로 계산되는지 확인 (원본 데이터엔 2월 근무기록 없을 가능성 — 3월 이후가 대부분)

수동 확인 결과를 커밋 메시지에 반영.

- [ ] **Step 7: 커밋**

```bash
cd "c:/Users/dhchd/work"
git add BUNT_finance/build_payroll.py
git commit -m "$(cat <<'EOF'
feat: 근무기록 급여유형 수식을 기간 기반 INDEX/MATCH로 교체

- master_lookup_formula 헬퍼 추가
- 근무일의 YYYYMM을 기준으로 직원 마스터의 활성 구간 조회
- 고경민 같은 다구간 직원의 시점별 유형 정확 반환

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: 급여 요약 시트 루프 범위 확장

**Files:**
- Modify: `BUNT_finance/build_payroll.py:252-417`

**의도:** `build_summary_sheet`는 `EMPLOYEES` 리스트를 순회하며 4행부터 한 행씩 배치한다. 기존 6명 → 신규 8행. 수식 로직(`active_check`, `daily_sum`, 조건부 서식 등)은 이미 F/G 기반이므로 그대로 동작. 루프와 하단 합계 행 좌표만 자동으로 재계산된다. 검증 위주 태스크.

- [ ] **Step 1: 루프 수정 없음 확인 (코드 리뷰)**

`build_summary_sheet` 함수의 핵심 루프 (라인 294-367):

```python
num_emp = len(EMPLOYEES)
for idx, emp in enumerate(EMPLOYEES):
    r = 4 + idx
    mr = idx + 2  # master row
    ...
```

`num_emp`와 `len(EMPLOYEES)`가 8로 확장되므로 4~11행까지 자동 생성된다. 수정 불필요.

- [ ] **Step 2: 월별 고경민 급여 유형 검증 스크립트 실행**

Run:
```bash
cd BUNT_finance && python build_payroll.py
```

Run (생성된 xlsx에서 급여요약 수식 활성 체크 검증):
```bash
cd BUNT_finance && python -c "
from openpyxl import load_workbook
wb = load_workbook('분트_인건비관리.xlsx')
ws = wb['급여 요약']
# 8명(행) 모두 4~11행에 매핑되어 있는지
for idx in range(8):
    r = 4 + idx
    name_formula = ws.cell(r,2).value
    assert name_formula is not None, f'Row {r}: 이름 수식 없음'
    assert '직원 마스터' in name_formula, f'Row {r}: 마스터 참조 없음'
# 하단 합계는 마지막 직원 행(11) + 2 = 13행
last_emp_row = 3 + 8  # =11
summary_row = last_emp_row + 2  # =13
total_formula = ws.cell(summary_row, 5).value
assert 'SUM(E4:E11)' in total_formula, f'합계 수식 범위 오류: {total_formula}'
print('OK: 급여 요약 8행 매핑, 합계 범위 E4:E11')
"
```

Expected: `OK: 급여 요약 8행 매핑, 합계 범위 E4:E11`

- [ ] **Step 3: Excel에서 월별 수동 검증**

생성된 `분트_인건비관리.xlsx`를 Excel에서 열고 **급여 요약** 탭 B1에 아래 월을 순서대로 입력하며 확인:

| 선택월 | 고경민 기대값 | 비고 |
|--------|---------------|------|
| 202502 | 급여유형="월급", 총급여(세전)=1,200,000 | 이진화는 표시 안 됨 |
| 202505 | 급여유형="시급", 총급여=근무기록 5월 일급합(600,000) | |
| 202511 | 급여유형="월급", 총급여=1,300,000 | |

각 월에서 고경민 행이 **3개 모두 보이지 않고 단 1개**만 표시되어야 한다. (다른 2개 구간 행은 해당월 비활성 → 이름 공란)

3개 월 모두 기대값과 일치하면 성공. 불일치 시 `active_check` 수식 디버깅.

- [ ] **Step 4: 커밋**

```bash
cd "c:/Users/dhchd/work"
git add BUNT_finance/build_payroll.py
git commit -m "$(cat <<'EOF'
feat: 급여 요약 시트 8행으로 확장 및 수동 검증 통과

- EMPLOYEES 8행 자동 매핑
- 고경민 3구간 중 선택월별 1개만 활성 표시 확인
- 합계 범위 E4:E11로 자동 확장

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: 출력 파일명 정렬 및 최종 생성·커밋

**Files:**
- Modify: `BUNT_finance/build_payroll.py:18`
- Create/Overwrite: `BUNT_finance/분트_인건비관리_v2.xlsx`

**의도:** 현재 스크립트의 `OUTPUT_PATH`는 `분트_인건비관리.xlsx`(v2 아님). 기존 git에 존재하는 파일은 `분트_인건비관리_v2.xlsx`. 이름을 일치시켜 재생성 시 기존 파일이 갱신되도록 한다.

- [ ] **Step 1: OUTPUT_PATH 수정**

현재 코드 (라인 18):
```python
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "분트_인건비관리.xlsx")
```

변경:
```python
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "분트_인건비관리_v2.xlsx")
```

- [ ] **Step 2: 기존 임시 파일 정리 및 재생성**

Excel에서 파일이 열려 있으면 닫는다(`~$분트_인건비관리_v2.xlsx` 잠금 파일 해제 확인).

Run:
```bash
cd BUNT_finance && ls ~\$*.xlsx 2>/dev/null && echo "WARNING: Excel이 파일을 잠그고 있음 - 닫아주세요" || echo "잠금 없음, 진행 가능"
python build_payroll.py
```

Expected: `[완료] 저장: .../분트_인건비관리_v2.xlsx`

- [ ] **Step 3: 구 파일 삭제 (있으면)**

```bash
cd BUNT_finance && [ -f 분트_인건비관리.xlsx ] && rm 분트_인건비관리.xlsx && echo "구 파일 삭제" || echo "구 파일 없음"
```

- [ ] **Step 4: 최종 커밋**

```bash
cd "c:/Users/dhchd/work"
git add BUNT_finance/build_payroll.py BUNT_finance/분트_인건비관리_v2.xlsx
# 혹시 구 파일 이름으로 추적되던 게 있으면 제거
git ls-files BUNT_finance/분트_인건비관리.xlsx 2>/dev/null && git rm BUNT_finance/분트_인건비관리.xlsx || true
git commit -m "$(cat <<'EOF'
chore: 출력 파일명을 _v2로 통일 및 최신 빌드 반영

- OUTPUT_PATH를 기존 커밋된 분트_인건비관리_v2.xlsx와 일치
- 기간별 급여유형 지원 빌드 결과물 포함

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: 종합 검증

**Files:** 없음 (검증만)

**의도:** 모든 변경이 누적된 상태에서 최종 동작을 한 번 더 종합 확인.

- [ ] **Step 1: 전체 검증 스크립트 실행**

```bash
cd BUNT_finance && python -c "
from openpyxl import load_workbook
wb = load_workbook('분트_인건비관리_v2.xlsx')

# 1. 시트 존재
assert set(wb.sheetnames) == {'직원 마스터','근무 기록','급여 요약'}

# 2. 마스터 8행
m = wb['직원 마스터']
assert m.cell(1,6).value == '적용시작월'
assert m.cell(1,7).value == '적용종료월'
kgm_rows = [r for r in range(2,10) if m.cell(r,1).value=='고경민']
assert kgm_rows == [2,3,4], f'고경민 행 위치: {kgm_rows}'

# 3. 근무기록 G열 기간기반
wr = wb['근무 기록']
g2 = wr.cell(2,7).value
assert 'INDEX' in g2 and 'YEAR' in g2

# 4. 급여요약 8행
s = wb['급여 요약']
for r in range(4,12):
    assert s.cell(r,2).value is not None
assert 'SUM(E4:E11)' in s.cell(13,5).value

print('OK: 전체 검증 통과')
"
```

Expected: `OK: 전체 검증 통과`

- [ ] **Step 2: 수동 브라우저 확인 체크리스트**

Excel에서 `분트_인건비관리_v2.xlsx` 열고 급여 요약 B1에 다음 입력하며 결과를 기록:

- [ ] `202501`: 고경민 월급 1,200,000 / 장예원 등 미표시
- [ ] `202502`: 고경민 월급 1,200,000 / 장예원·유은비 시급 표시
- [ ] `202505`: 고경민 **시급** 유형, 총급여 = 5월 일급 합(600,000)
- [ ] `202509`: 고경민 시급 유형, 총급여 = 9월 일급 합(672,000)
- [ ] `202510`: 고경민 **월급** 1,300,000 (시급 구간 끝, 월급 복귀)
- [ ] `202511`: 고경민 월급 1,300,000 / 이진화 시급 표시 시작

각 월에서 고경민 행이 **정확히 1개**만 표시되어야 함.

- [ ] **Step 3: 검증 완료 보고**

전체 체크리스트 통과 시 사용자에게 완료 보고. 실패 항목이 있으면 해당 Task로 돌아가 수정.

---

## Self-Review 결과

**Spec coverage:**
- ✅ 직원 마스터 한 직원 여러 행 → Task 1, 2
- ✅ F/G 컬럼 의미 변경 및 헤더 변경 → Task 2
- ✅ 고경민 3구간 초기 데이터 → Task 1
- ✅ 근무 기록 G열 기간 기반 수식 → Task 3
- ✅ 빈 입력행 F열 시급 조회도 기간 기반 → Task 3 Step 4
- ✅ 급여 요약 루프 범위 확장 → Task 4
- ✅ 재직상태 E열 필터링 미사용 (기존 로직 유지) → Task 4 (코드 변경 없음으로 자연 충족)
- ✅ 중복 표시 방지(불변 조건) → Task 4 Step 3 수동 검증

**Placeholder scan:** 모든 단계에 구체적 코드·명령·기대 출력 포함. "적절히 처리" 등의 모호 표현 없음.

**Type consistency:** `master_lookup_formula`는 Task 3 Step 1에서 정의되고 Step 2~4에서 `master_lookup_formula(name_ref, date_ref, return_col)` 시그니처로 일관되게 호출됨. 컬럼 상수("B", "C")는 급여유형/금액 컬럼으로 일관.
