# 분트 인건비 엑셀 고도화 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 기존 xlsb 인건비 시트를 독립 xlsx 파일로 재설계하여 입력 편의성, 시각 완성도, 급여 자동 계산을 개선한다.

**Architecture:** 단일 Python 스크립트(`build_payroll.py`)가 기존 xlsb에서 데이터를 읽고, openpyxl로 3개 시트(직원 마스터, 근무 기록, 급여 요약)를 갖춘 xlsx를 생성한다. 수식/드롭다운/조건부서식은 모두 엑셀 네이티브 기능으로 심어서, 생성 이후에는 Python 없이 엑셀만으로 운영 가능하다.

**Tech Stack:** Python 3, openpyxl, pyxlsb

---

## File Structure

| 파일 | 역할 |
|------|------|
| `BUNT_finance/build_payroll.py` | 엑셀 생성 스크립트 (실행 시 xlsx 출력) |
| `BUNT_finance/분트_인건비관리.xlsx` | 최종 산출물 (스크립트가 생성) |

소스 파일 1개, 산출물 1개. 스크립트는 4개 함수로 구성:
- `extract_data(xlsb_path)` → 기존 데이터 파싱, dict 반환
- `build_master_sheet(wb, employees)` → 직원 마스터 시트 생성
- `build_records_sheet(wb, records, employees)` → 근무 기록 시트 생성
- `build_summary_sheet(wb, employees)` → 급여 요약 시트 생성

---

### Task 1: 데이터 추출 함수 작성

**Files:**
- Create: `BUNT_finance/build_payroll.py`

- [ ] **Step 1: 스크립트 기본 구조 + extract_data 함수 작성**

```python
"""분트 미술학원 인건비 관리 엑셀 생성 스크립트"""
from pyxlsb import open_workbook
from datetime import datetime, timedelta

XLSB_PATH = "분트 재무제표(25.12월).xlsb"
OUTPUT_PATH = "분트_인건비관리.xlsx"

# 직원 마스터 정보 (xlsb에서 확인된 데이터)
EMPLOYEES = [
    {"name": "고경민", "pay_type": "월급", "amount": 1200000, "account": "국민은행 82240104164295", "status": "재직"},
    {"name": "장예원", "pay_type": "시급", "amount": 12000, "account": "카카오뱅크 3333-07-7072641", "status": "퇴직"},
    {"name": "유은비", "pay_type": "시급", "amount": 12000, "account": "우리 1002-533-898534", "status": "퇴직"},
    {"name": "전민아", "pay_type": "시급", "amount": 12000, "account": "하나은행 558-910330-30707", "status": "재직"},
    {"name": "김채현", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-351-542250", "status": "재직"},
    {"name": "이진화", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-166-678383", "status": "재직"},
]

def excel_serial_to_date(serial):
    """엑셀 시리얼 넘버를 datetime으로 변환"""
    if serial is None or not isinstance(serial, (int, float)):
        return None
    return datetime(1899, 12, 30) + timedelta(days=int(serial))


def extract_data(xlsb_path):
    """기존 xlsb에서 근무 기록을 추출한다.
    
    Returns:
        list[dict]: 각 항목은 {date, worker, start, end, hours, daily_pay, note}
    """
    records = []
    wb = open_workbook(xlsb_path)
    
    with wb.get_sheet("2. 인건비 계산") as sheet:
        rows = []
        for row in sheet.rows():
            cells = [c.v for c in row]
            rows.append(cells)
    
    for row in rows:
        # 근무 기록 행 판별: B열(index 1)이 엑셀 날짜 시리얼(숫자), C열(index 2)이 근무자명(문자열)
        if (len(row) > 6 
            and isinstance(row[1], (int, float)) 
            and row[1] > 40000  # 엑셀 날짜 시리얼 범위
            and isinstance(row[2], str)
            and row[2].strip() not in ("휴무", "근무자", "")):
            
            dt = excel_serial_to_date(row[1])
            if dt is None:
                continue
            
            worker = row[2].strip()
            start_time = row[3] if isinstance(row[3], (int, float)) else None
            end_time = row[4] if isinstance(row[4], (int, float)) else None
            hours = row[5] if isinstance(row[5], (int, float)) else 0
            daily_pay = row[6] if isinstance(row[6], (int, float)) else 0
            note = str(row[7]).strip() if row[7] is not None else ""
            # 비고에 "경민쌤근무", "전민아쌤 X" 등이 있는 경우
            if note == "0" or note == "0.0":
                note = ""
            
            records.append({
                "date": dt,
                "worker": worker,
                "start": start_time,
                "end": end_time,
                "hours": hours,
                "daily_pay": daily_pay,
                "note": note,
            })
    
    # 날짜 오름차순 정렬
    records.sort(key=lambda r: (r["date"], r["worker"]))
    return records
```

- [ ] **Step 2: 추출 결과 검증용 임시 코드 추가 및 실행**

파일 맨 아래에 추가:

```python
if __name__ == "__main__":
    records = extract_data(XLSB_PATH)
    print(f"총 {len(records)}건 추출")
    for r in records[:5]:
        print(f"  {r['date'].strftime('%Y-%m-%d')} | {r['worker']} | {r['start']}~{r['end']} | {r['hours']}h | {r['daily_pay']}원")
    print("...")
    for r in records[-5:]:
        print(f"  {r['date'].strftime('%Y-%m-%d')} | {r['worker']} | {r['start']}~{r['end']} | {r['hours']}h | {r['daily_pay']}원")
    
    # 근무자별 건수
    from collections import Counter
    counts = Counter(r["worker"] for r in records)
    print("\n근무자별 건수:")
    for name, cnt in counts.most_common():
        print(f"  {name}: {cnt}건")
```

Run: `cd BUNT_finance && python build_payroll.py`
Expected: 수백 건의 근무 기록이 출력되고, 6명의 근무자가 나열됨

- [ ] **Step 3: 커밋**

```bash
git add BUNT_finance/build_payroll.py
git commit -m "feat: 인건비 엑셀 생성 스크립트 - 데이터 추출 함수 구현"
```

---

### Task 2: 직원 마스터 시트 생성

**Files:**
- Modify: `BUNT_finance/build_payroll.py`

- [ ] **Step 1: openpyxl import 및 스타일 상수 정의**

파일 상단 import 영역에 추가:

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# 스타일 상수
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
CALC_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
MONEY_FMT = '#,##0'
```

- [ ] **Step 2: build_master_sheet 함수 작성**

```python
def build_master_sheet(wb, employees):
    """직원 마스터 시트를 생성한다."""
    ws = wb.active
    ws.title = "직원 마스터"
    ws.sheet_properties.tabColor = "4472C4"  # 파랑
    
    # 헤더
    headers = ["이름", "급여유형", "금액", "입금계좌", "재직상태", "입사일", "비고"]
    col_widths = [12, 12, 16, 32, 12, 14, 20]
    
    for col_idx, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    
    # 데이터
    for row_idx, emp in enumerate(employees, 2):
        ws.cell(row=row_idx, column=1, value=emp["name"]).border = THIN_BORDER
        ws.cell(row=row_idx, column=2, value=emp["pay_type"]).border = THIN_BORDER
        
        amount_cell = ws.cell(row=row_idx, column=3, value=emp["amount"])
        amount_cell.number_format = MONEY_FMT
        amount_cell.border = THIN_BORDER
        
        ws.cell(row=row_idx, column=4, value=emp["account"]).border = THIN_BORDER
        ws.cell(row=row_idx, column=5, value=emp["status"]).border = THIN_BORDER
        ws.cell(row=row_idx, column=6).border = THIN_BORDER  # 입사일은 비워둠
        ws.cell(row=row_idx, column=7).border = THIN_BORDER
        
        # 입력 셀 배경
        for col in range(1, 8):
            ws.cell(row=row_idx, column=col).fill = INPUT_FILL
    
    # 드롭다운: 급여유형
    dv_pay = DataValidation(type="list", formula1='"월급,시급"', allow_blank=True)
    dv_pay.error = "월급 또는 시급만 선택 가능합니다"
    dv_pay.errorTitle = "입력 오류"
    ws.add_data_validation(dv_pay)
    dv_pay.add(f"B2:B50")
    
    # 드롭다운: 재직상태
    dv_status = DataValidation(type="list", formula1='"재직,퇴직"', allow_blank=True)
    dv_status.error = "재직 또는 퇴직만 선택 가능합니다"
    dv_status.errorTitle = "입력 오류"
    ws.add_data_validation(dv_status)
    dv_status.add(f"E2:E50")
    
    # 행 높이
    ws.row_dimensions[1].height = 25
    
    return ws
```

- [ ] **Step 3: main에 워크북 생성 + 마스터 시트 호출 추가**

기존 `if __name__ == "__main__":` 블록을 교체:

```python
if __name__ == "__main__":
    # 1. 데이터 추출
    print("기존 xlsb에서 데이터 추출 중...")
    records = extract_data(XLSB_PATH)
    print(f"  → {len(records)}건 근무 기록 추출 완료")
    
    # 2. 새 워크북 생성
    wb = Workbook()
    
    # 3. 직원 마스터
    print("직원 마스터 시트 생성 중...")
    build_master_sheet(wb, EMPLOYEES)
    
    # 저장
    wb.save(OUTPUT_PATH)
    print(f"\n✓ {OUTPUT_PATH} 생성 완료")
```

- [ ] **Step 4: 실행 및 검증**

Run: `cd BUNT_finance && python build_payroll.py`
Expected: `분트_인건비관리.xlsx` 생성, 엑셀에서 열어 직원 마스터 시트에 6명 데이터 + 드롭다운 확인

- [ ] **Step 5: 커밋**

```bash
git add BUNT_finance/build_payroll.py
git commit -m "feat: 직원 마스터 시트 생성 (드롭다운, 스타일 적용)"
```

---

### Task 3: 근무 기록 시트 생성

**Files:**
- Modify: `BUNT_finance/build_payroll.py`

- [ ] **Step 1: build_records_sheet 함수 작성**

```python
def build_records_sheet(wb, records, employees):
    """근무 기록 시트를 생성한다. 기존 데이터 마이그레이션 + 수식/드롭다운 설정."""
    ws = wb.create_sheet("근무 기록")
    ws.sheet_properties.tabColor = "70AD47"  # 초록
    
    # 헤더
    headers = ["날짜", "근무자", "시작시간", "종료시간", "근무시간", "일 급여", "급여유형", "비고"]
    col_widths = [14, 12, 12, 12, 12, 14, 10, 24]
    
    for col_idx, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    
    ws.row_dimensions[1].height = 25
    
    # 기존 데이터 입력
    for row_idx, rec in enumerate(records, 2):
        # A: 날짜
        date_cell = ws.cell(row=row_idx, column=1, value=rec["date"])
        date_cell.number_format = "YYYY-MM-DD"
        date_cell.fill = INPUT_FILL
        date_cell.border = THIN_BORDER
        
        # B: 근무자
        name_cell = ws.cell(row=row_idx, column=2, value=rec["worker"])
        name_cell.fill = INPUT_FILL
        name_cell.border = THIN_BORDER
        
        # C: 시작시간
        start_cell = ws.cell(row=row_idx, column=3, value=rec["start"])
        start_cell.fill = INPUT_FILL
        start_cell.border = THIN_BORDER
        
        # D: 종료시간
        end_cell = ws.cell(row=row_idx, column=4, value=rec["end"])
        end_cell.fill = INPUT_FILL
        end_cell.border = THIN_BORDER
        
        # E: 근무시간 (수식)
        hours_cell = ws.cell(row=row_idx, column=5)
        hours_cell.value = f"=IF(AND(C{row_idx}<>\"\",D{row_idx}<>\"\"),D{row_idx}-C{row_idx},\"\")"
        hours_cell.fill = CALC_FILL
        hours_cell.border = THIN_BORDER
        
        # F: 일 급여 (수식 - 시급제만 계산)
        pay_cell = ws.cell(row=row_idx, column=6)
        pay_cell.value = f'=IF(G{row_idx}="시급",E{row_idx}*VLOOKUP(B{row_idx},\'직원 마스터\'!A:C,3,FALSE),"")'
        pay_cell.number_format = MONEY_FMT
        pay_cell.fill = CALC_FILL
        pay_cell.border = THIN_BORDER
        
        # G: 급여유형 (수식)
        type_cell = ws.cell(row=row_idx, column=7)
        type_cell.value = f"=IF(B{row_idx}<>\"\",VLOOKUP(B{row_idx},'직원 마스터'!A:B,2,FALSE),\"\")"
        type_cell.fill = CALC_FILL
        type_cell.border = THIN_BORDER
        
        # H: 비고
        note_cell = ws.cell(row=row_idx, column=8, value=rec["note"] if rec["note"] else None)
        note_cell.fill = INPUT_FILL
        note_cell.border = THIN_BORDER
    
    # 빈 행 50개 추가 (향후 입력용, 수식 포함)
    last_data_row = len(records) + 1
    for row_idx in range(last_data_row + 1, last_data_row + 51):
        for col in range(1, 9):
            cell = ws.cell(row=row_idx, column=col)
            cell.border = THIN_BORDER
            if col in (1, 2, 3, 4, 8):
                cell.fill = INPUT_FILL
            else:
                cell.fill = CALC_FILL
        
        # 날짜 서식
        ws.cell(row=row_idx, column=1).number_format = "YYYY-MM-DD"
        
        # E: 근무시간
        ws.cell(row=row_idx, column=5).value = f"=IF(AND(C{row_idx}<>\"\",D{row_idx}<>\"\"),D{row_idx}-C{row_idx},\"\")"
        
        # F: 일 급여
        ws.cell(row=row_idx, column=6).value = f'=IF(G{row_idx}="시급",E{row_idx}*VLOOKUP(B{row_idx},\'직원 마스터\'!A:C,3,FALSE),"")'
        ws.cell(row=row_idx, column=6).number_format = MONEY_FMT
        
        # G: 급여유형
        ws.cell(row=row_idx, column=7).value = f"=IF(B{row_idx}<>\"\",VLOOKUP(B{row_idx},'직원 마스터'!A:B,2,FALSE),\"\")"
    
    # 드롭다운: 근무자 (직원 마스터의 이름 참조)
    emp_count = len(employees)
    dv_worker = DataValidation(
        type="list",
        formula1=f"'직원 마스터'!$A$2:$A${emp_count + 1}",
        allow_blank=True,
    )
    dv_worker.error = "직원 마스터에 등록된 이름만 선택 가능합니다"
    dv_worker.errorTitle = "입력 오류"
    ws.add_data_validation(dv_worker)
    dv_worker.add(f"B2:B{last_data_row + 50}")
    
    # 자동 필터
    ws.auto_filter.ref = f"A1:H{last_data_row + 50}"
    
    return ws
```

- [ ] **Step 2: main에 근무 기록 시트 호출 추가**

`build_master_sheet` 호출 이후에 추가:

```python
    # 4. 근무 기록
    print("근무 기록 시트 생성 중...")
    build_records_sheet(wb, records, EMPLOYEES)
```

- [ ] **Step 3: 실행 및 검증**

Run: `cd BUNT_finance && python build_payroll.py`
Expected: xlsx에 근무 기록 시트 추가됨. 엑셀에서 열어 확인:
- 기존 데이터가 날짜순으로 정렬되어 있음
- 근무시간/일급여/급여유형 열에 수식이 작동함
- 근무자 드롭다운이 직원 마스터 이름을 참조함
- 하단에 빈 입력 행 50개가 수식 포함으로 준비됨

- [ ] **Step 4: 커밋**

```bash
git add BUNT_finance/build_payroll.py
git commit -m "feat: 근무 기록 시트 생성 (데이터 마이그레이션 + 수식 + 드롭다운)"
```

---

### Task 4: 급여 요약 시트 생성

**Files:**
- Modify: `BUNT_finance/build_payroll.py`

- [ ] **Step 1: build_summary_sheet 함수 작성**

```python
def build_summary_sheet(wb, employees):
    """급여 요약 시트를 생성한다. 월 선택 → 자동 집계."""
    ws = wb.create_sheet("급여 요약")
    ws.sheet_properties.tabColor = "ED7D31"  # 주황
    
    # --- 상단: 월 선택 ---
    ws.cell(row=1, column=1, value="조회 월:").font = Font(bold=True, size=12)
    month_cell = ws.cell(row=1, column=2, value=202501)
    month_cell.font = Font(bold=True, size=14, color="1F4E79")
    month_cell.fill = INPUT_FILL
    month_cell.border = THIN_BORDER
    month_cell.number_format = "0"
    
    # 월 드롭다운
    dv_month = DataValidation(
        type="list",
        formula1='"202501,202502,202503,202504,202505,202506,202507,202508,202509,202510,202511,202512"',
        allow_blank=False,
    )
    ws.add_data_validation(dv_month)
    dv_month.add("B1")
    
    # --- 날짜 계산 보조 셀 (숨김 영역 J1:J2) ---
    # J1: 해당월 1일, J2: 해당월 말일
    ws.cell(row=1, column=10).value = '=DATE(INT(B1/100),MOD(B1,100),1)'
    ws.cell(row=1, column=10).number_format = "YYYY-MM-DD"
    ws.cell(row=2, column=10).value = '=EOMONTH(J1,0)'
    ws.cell(row=2, column=10).number_format = "YYYY-MM-DD"
    # J열 숨기기
    ws.column_dimensions["J"].hidden = True
    
    # --- 헤더 (Row 3) ---
    headers = ["이름", "급여유형", "총 근무시간", "총 급여(세전)", "3.3% 원천징수", "실지급액", "입금계좌", "지급여부"]
    col_widths = [12, 12, 14, 16, 16, 16, 32, 10]
    
    for col_idx, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    
    ws.row_dimensions[3].height = 25
    
    # --- 직원별 급여 행 (Row 4~) ---
    for i, emp in enumerate(employees):
        row = 4 + i
        master_row = 2 + i  # 직원 마스터에서의 행 번호
        
        # A: 이름 (마스터 참조)
        ws.cell(row=row, column=1).value = f"='직원 마스터'!A{master_row}"
        ws.cell(row=row, column=1).border = THIN_BORDER
        
        # B: 급여유형 (마스터 참조)
        ws.cell(row=row, column=2).value = f"='직원 마스터'!B{master_row}"
        ws.cell(row=row, column=2).border = THIN_BORDER
        ws.cell(row=row, column=2).fill = CALC_FILL
        
        # C: 총 근무시간 (SUMIFS)
        ws.cell(row=row, column=3).value = (
            f"=SUMIFS('근무 기록'!E:E,"
            f"'근무 기록'!B:B,A{row},"
            f"'근무 기록'!A:A,\">=\"&$J$1,"
            f"'근무 기록'!A:A,\"<=\"&$J$2)"
        )
        ws.cell(row=row, column=3).fill = CALC_FILL
        ws.cell(row=row, column=3).border = THIN_BORDER
        ws.cell(row=row, column=3).number_format = "0.0"
        
        # D: 총 급여(세전)
        ws.cell(row=row, column=4).value = (
            f'=IF(B{row}="월급",\'직원 마스터\'!C{master_row},C{row}*\'직원 마스터\'!C{master_row})'
        )
        ws.cell(row=row, column=4).number_format = MONEY_FMT
        ws.cell(row=row, column=4).fill = CALC_FILL
        ws.cell(row=row, column=4).border = THIN_BORDER
        
        # E: 3.3% 원천징수
        ws.cell(row=row, column=5).value = f"=ROUND(D{row}*0.033,0)"
        ws.cell(row=row, column=5).number_format = MONEY_FMT
        ws.cell(row=row, column=5).fill = CALC_FILL
        ws.cell(row=row, column=5).border = THIN_BORDER
        
        # F: 실지급액
        ws.cell(row=row, column=6).value = f"=D{row}-E{row}"
        ws.cell(row=row, column=6).number_format = MONEY_FMT
        ws.cell(row=row, column=6).fill = CALC_FILL
        ws.cell(row=row, column=6).border = THIN_BORDER
        
        # G: 입금계좌 (마스터 참조)
        ws.cell(row=row, column=7).value = f"='직원 마스터'!D{master_row}"
        ws.cell(row=row, column=7).fill = CALC_FILL
        ws.cell(row=row, column=7).border = THIN_BORDER
        
        # H: 지급여부 (입력)
        ws.cell(row=row, column=8).fill = INPUT_FILL
        ws.cell(row=row, column=8).border = THIN_BORDER
        ws.cell(row=row, column=8).alignment = Alignment(horizontal="center")
    
    # 지급여부 드롭다운
    last_emp_row = 3 + len(employees)
    dv_paid = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
    ws.add_data_validation(dv_paid)
    dv_paid.add(f"H4:H{last_emp_row}")
    
    # 조건부 서식: 지급여부
    green_font = Font(color="008000", bold=True)
    red_font = Font(color="FF0000", bold=True)
    ws.conditional_formatting.add(
        f"H4:H{last_emp_row}",
        CellIsRule(operator="equal", formula=['"O"'], font=green_font),
    )
    ws.conditional_formatting.add(
        f"H4:H{last_emp_row}",
        CellIsRule(operator="equal", formula=['"X"'], font=red_font),
    )
    
    # --- 하단 요약 ---
    summary_row = last_emp_row + 2
    
    ws.cell(row=summary_row, column=3, value="합계").font = Font(bold=True)
    ws.cell(row=summary_row, column=3).alignment = Alignment(horizontal="right")
    ws.cell(row=summary_row, column=3).border = THIN_BORDER
    
    # 총 급여 합계
    ws.cell(row=summary_row, column=4).value = f"=SUM(D4:D{last_emp_row})"
    ws.cell(row=summary_row, column=4).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=4).font = Font(bold=True, size=12)
    ws.cell(row=summary_row, column=4).border = THIN_BORDER
    
    # 원천징수 합계
    ws.cell(row=summary_row, column=5).value = f"=SUM(E4:E{last_emp_row})"
    ws.cell(row=summary_row, column=5).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=5).font = Font(bold=True)
    ws.cell(row=summary_row, column=5).border = THIN_BORDER
    
    # 실지급 합계
    ws.cell(row=summary_row, column=6).value = f"=SUM(F4:F{last_emp_row})"
    ws.cell(row=summary_row, column=6).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=6).font = Font(bold=True, size=12, color="1F4E79")
    ws.cell(row=summary_row, column=6).border = THIN_BORDER
    
    # 당월 총 인건비 라벨
    label_row = summary_row + 1
    ws.cell(row=label_row, column=3, value="당월 총 인건비").font = Font(bold=True, size=11)
    ws.cell(row=label_row, column=3).alignment = Alignment(horizontal="right")
    ws.cell(row=label_row, column=4).value = f"=D{summary_row}"
    ws.cell(row=label_row, column=4).number_format = MONEY_FMT
    ws.cell(row=label_row, column=4).font = Font(bold=True, size=14, color="C00000")
    ws.cell(row=label_row, column=4).border = THIN_BORDER
    
    return ws
```

- [ ] **Step 2: main에 급여 요약 시트 호출 추가**

근무 기록 시트 호출 이후에 추가:

```python
    # 5. 급여 요약
    print("급여 요약 시트 생성 중...")
    build_summary_sheet(wb, EMPLOYEES)
```

- [ ] **Step 3: 실행 및 검증**

Run: `cd BUNT_finance && python build_payroll.py`
Expected: 3개 시트 모두 생성. 엑셀에서 열어 확인:
- 월 선택 드롭다운 작동
- 202502 선택 시 2월 근무 데이터 기반 급여 자동 집계
- 3.3% 원천징수, 실지급액 자동 계산
- 합계/당월 총 인건비 표시
- 지급여부 O/X 조건부 서식 (초록/빨강)

- [ ] **Step 4: 커밋**

```bash
git add BUNT_finance/build_payroll.py
git commit -m "feat: 급여 요약 시트 생성 (SUMIFS 자동 집계, 3.3% 계산, 조건부서식)"
```

---

### Task 5: 최종 검증 및 정리

**Files:**
- Modify: `BUNT_finance/build_payroll.py`

- [ ] **Step 1: main 함수의 디버그 코드 제거 및 최종 정리**

`if __name__ == "__main__":` 블록 최종 형태:

```python
if __name__ == "__main__":
    print("=" * 50)
    print("분트 미술학원 인건비 관리 엑셀 생성")
    print("=" * 50)
    
    # 1. 데이터 추출
    print("\n[1/4] 기존 xlsb에서 데이터 추출 중...")
    records = extract_data(XLSB_PATH)
    print(f"  → {len(records)}건 근무 기록 추출 완료")
    
    # 2. 새 워크북 생성
    wb = Workbook()
    
    # 3. 직원 마스터
    print("[2/4] 직원 마스터 시트 생성 중...")
    build_master_sheet(wb, EMPLOYEES)
    
    # 4. 근무 기록
    print("[3/4] 근무 기록 시트 생성 중...")
    build_records_sheet(wb, records, EMPLOYEES)
    
    # 5. 급여 요약
    print("[4/4] 급여 요약 시트 생성 중...")
    build_summary_sheet(wb, EMPLOYEES)
    
    # 6. 저장
    wb.save(OUTPUT_PATH)
    print(f"\n{'=' * 50}")
    print(f"  {OUTPUT_PATH} 생성 완료!")
    print(f"  - 직원 마스터: {len(EMPLOYEES)}명")
    print(f"  - 근무 기록: {len(records)}건")
    print(f"  - 급여 요약: 월 선택 → 자동 집계")
    print(f"{'=' * 50}")
```

- [ ] **Step 2: 전체 실행 및 최종 검증**

Run: `cd BUNT_finance && python build_payroll.py`

엑셀에서 최종 검증 체크리스트:
1. 직원 마스터: 6명 데이터, 드롭다운(급여유형/재직상태) 작동
2. 근무 기록: 기존 데이터 마이그레이션됨, 수식(근무시간/일급여/급여유형) 작동, 드롭다운 작동
3. 급여 요약: 월 선택 시 집계 변경, 3.3% 계산 정확, 조건부서식 작동
4. 시각: 헤더 남색, 입력셀 노랑, 계산셀 회색, 시트탭 색상
5. 기존 xlsb 파일은 그대로 보존됨

- [ ] **Step 3: 최종 커밋**

```bash
git add BUNT_finance/build_payroll.py
git commit -m "feat: 인건비 엑셀 생성 스크립트 완성 - 3시트 구조, 데이터 마이그레이션, 자동 급여 계산"
```
