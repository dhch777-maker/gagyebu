"""
분트 미술학원 인건비 관리 엑셀 생성 스크립트
- 원본: 분트 재무제표(25.12월).xlsb → "2. 인건비 계산" 시트
- 출력: 분트_인건비관리.xlsx (3개 시트)
"""

import os
from datetime import datetime, timedelta
from pyxlsb import open_workbook as open_xlsb
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ── Constants ──────────────────────────────────────────────────────────
XLSB_PATH = os.path.join(os.path.dirname(__file__), "분트 재무제표(25.12월).xlsb")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "분트_인건비관리_v2.xlsx")
SHEET_NAME = "2. 인건비 계산"

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

# 직원 마스터 조회 범위 끝 행 — EMPLOYEES 길이 + 신규 추가 여유분(20)
# 근무기록 수식들은 이 범위를 직원 마스터 조회 대상으로 참조
MASTER_LAST_ROW = len(EMPLOYEES) + 21  # header row(1) + data rows + 20 buffer

# 지급 이력 — xlsb 원본에서 추출한 (이름, YYYYMM, 지급여부) 튜플
# - 이름: 직원 마스터의 이름과 일치
# - YYYYMM: 지급 대상 월 (xlsb 지급일을 기준월로 정규화: 1~10일=전월분, 11일 이후=당월분)
# - 지급여부: 해당 월의 모든 지급 기록이 xlsb에서 'O' 표시된 경우 "O", 하나라도 누락 시 "X"
# 과거 지급 이력 수정이나 신규 지급 반영 시 이 리스트 또는 "지급 이력" 시트에 직접 행 추가
PAYMENT_HISTORY = [
    ("고경민", 202501, "O"),
    ("유은비", 202501, "O"),
    ("장예원", 202501, "O"),
    ("고경민", 202502, "O"),
    ("유은비", 202502, "O"),
    ("장예원", 202502, "O"),
    ("전민아", 202502, "O"),
    ("고경민", 202503, "O"),
    ("장예원", 202503, "O"),
    ("전민아", 202503, "O"),
    ("고경민", 202504, "O"),
    ("장예원", 202504, "O"),
    ("전민아", 202504, "O"),
    # 5·6월: xlsb 원본엔 'O' 표시 누락이나 실제 지급 확인되어 보정
    ("고경민", 202505, "O"),
    ("장예원", 202505, "O"),
    ("전민아", 202505, "O"),
    ("고경민", 202506, "O"),
    ("김채현", 202506, "O"),
    ("장예원", 202506, "O"),
    ("전민아", 202506, "O"),
    ("고경민", 202507, "O"),
    ("김채현", 202507, "O"),
    ("전민아", 202507, "O"),
    ("고경민", 202508, "O"),
    ("김채현", 202508, "O"),
    ("전민아", 202508, "O"),
    ("고경민", 202509, "O"),
    ("김채현", 202509, "O"),
    ("전민아", 202509, "O"),
    ("고경민", 202510, "O"),
    ("이진화", 202510, "O"),
    ("전민아", 202510, "O"),
    ("고경민", 202511, "X"),
    ("이진화", 202511, "X"),
    ("전민아", 202511, "X"),
]

EXCEL_EPOCH = datetime(1899, 12, 30)

# ── Styles ─────────────────────────────────────────────────────────────
NAVY_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=11)
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GRAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
MONEY_FMT = "#,##0"


def excel_serial_to_date(serial):
    """Excel serial number → datetime"""
    return EXCEL_EPOCH + timedelta(days=int(serial))


def extract_data():
    """xlsb에서 근무 기록 추출"""
    records = []
    wb = open_xlsb(XLSB_PATH)
    with wb.get_sheet(SHEET_NAME) as sheet:
        for row in sheet.rows():
            if len(row) < 7:
                continue
            b_val = row[1].v
            c_val = row[2].v
            if b_val is None or c_val is None:
                continue
            try:
                serial = float(b_val)
            except (ValueError, TypeError):
                continue
            if serial <= 40000:
                continue

            worker = str(c_val).strip()
            if worker in ("휴무", "근무자", "") or not worker:
                continue

            dt = excel_serial_to_date(serial)
            start_time = row[3].v
            end_time = row[4].v
            hours = row[5].v
            daily_pay = row[6].v

            records.append({
                "date": dt,
                "worker": worker,
                "start": start_time if start_time else None,
                "end": end_time if end_time else None,
                "hours": float(hours) if hours else 0,
                "daily_pay": float(daily_pay) if daily_pay else 0,
            })
    records.sort(key=lambda r: r["date"])
    return records


def apply_header_style(ws, row, col_count):
    """Navy header row styling"""
    ws.row_dimensions[row].height = 25
    for c in range(1, col_count + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = NAVY_FILL
        cell.font = WHITE_BOLD
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def apply_cell_style(cell, fill=None, fmt=None):
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt


def master_lookup_formula(name_ref, date_ref, return_col, master_last_row=None):
    """
    직원 마스터에서 특정 시점에 활성인 구간을 찾아 지정 컬럼 값을 반환하는 수식.
    - name_ref: 근무자 이름 셀 참조 (예: "B5")
    - date_ref: 날짜 셀 참조 (예: "A5")
    - return_col: 반환할 마스터 컬럼 문자 (예: "B" = 급여유형, "C" = 금액)
    - master_last_row: 마스터 조회 범위 끝 행 (기본값: MASTER_LAST_ROW = len(EMPLOYEES)+21)

    매칭 조건: 이름 일치 AND 적용시작월 <= 근무월 AND (적용종료월 공란 또는 >= 근무월)
    근무월 = YEAR*100 + MONTH
    """
    if master_last_row is None:
        master_last_row = MASTER_LAST_ROW
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


def build_master_sheet(wb):
    """Sheet 1: 직원 마스터"""
    ws = wb.active
    ws.title = "직원 마스터"
    ws.sheet_properties.tabColor = "4472C4"

    headers = ["이름", "급여유형", "금액", "입금계좌", "재직상태", "적용시작월", "적용종료월", "비고"]
    col_widths = [12, 12, 15, 30, 12, 12, 12, 20]

    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    apply_header_style(ws, 1, len(headers))

    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r, emp in enumerate(EMPLOYEES, 2):
        ws.cell(row=r, column=1, value=emp["name"])
        ws.cell(row=r, column=2, value=emp["pay_type"])
        ws.cell(row=r, column=3, value=emp["amount"])
        ws.cell(row=r, column=4, value=emp["account"])
        ws.cell(row=r, column=5, value=emp["status"])
        ws.cell(row=r, column=6, value=emp["start_month"])
        ws.cell(row=r, column=6).number_format = "0"
        ws.cell(row=r, column=7, value=emp["end_month"])
        ws.cell(row=r, column=7).number_format = "0"
        for c in range(1, len(headers) + 1):
            apply_cell_style(ws.cell(row=r, column=c), fill=YELLOW_FILL)
        ws.cell(row=r, column=3).number_format = MONEY_FMT

    # Dropdowns
    dv_pay = DataValidation(type="list", formula1='"월급,시급"', allow_blank=True)
    dv_pay.error = "월급 또는 시급만 입력 가능합니다"
    ws.add_data_validation(dv_pay)
    dv_pay.add(f"B2:B{len(EMPLOYEES) + 20}")

    dv_status = DataValidation(type="list", formula1='"재직,퇴직"', allow_blank=True)
    ws.add_data_validation(dv_status)
    dv_status.add(f"E2:E{len(EMPLOYEES) + 20}")

    return ws


def build_records_sheet(wb, records):
    """Sheet 2: 근무 기록
    마이그레이션 데이터: 일 급여를 원본 값 그대로 저장 (수식 X)
    신규 입력 행: 수식으로 자동 계산
    """
    ws = wb.create_sheet("근무 기록")
    ws.sheet_properties.tabColor = "70AD47"

    headers = ["날짜", "근무자", "시작시간", "종료시간", "근무시간", "일 급여", "급여유형", "비고"]
    col_widths = [14, 12, 12, 12, 12, 15, 12, 20]

    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    apply_header_style(ws, 1, len(headers))

    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    INPUT_COLS = {1, 2, 3, 4, 8}

    # Write extracted data — 원본 값 그대로 저장
    for idx, rec in enumerate(records):
        r = idx + 2
        ws.cell(row=r, column=1, value=rec["date"])
        ws.cell(row=r, column=1).number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=2, value=rec["worker"])

        if rec["start"] is not None:
            ws.cell(row=r, column=3, value=float(rec["start"]) / 24)
            ws.cell(row=r, column=3).number_format = "H:MM"
        if rec["end"] is not None:
            ws.cell(row=r, column=4, value=float(rec["end"]) / 24)
            ws.cell(row=r, column=4).number_format = "H:MM"

        # 근무시간: 수식
        ws.cell(row=r, column=5).value = f'=IF(AND(C{r}<>"",D{r}<>""),D{r}-C{r},0)'
        ws.cell(row=r, column=5).number_format = "0.0"

        # 일 급여: 원본 값 그대로 (수식 아님!)
        ws.cell(row=r, column=6, value=int(rec["daily_pay"]) if rec["daily_pay"] else 0)
        ws.cell(row=r, column=6).number_format = MONEY_FMT

        # 급여유형: 날짜 기반 기간 조회
        lookup = master_lookup_formula(f"B{r}", f"A{r}", "B")
        ws.cell(row=r, column=7).value = f'=IFERROR(IF(B{r}<>"",{lookup},""),"")'

        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c in INPUT_COLS or c == 6:  # 일 급여도 마이그레이션 데이터는 입력값
                apply_cell_style(cell, fill=YELLOW_FILL)
            else:
                apply_cell_style(cell, fill=GRAY_FILL)

    # 빈 행 50개 (신규 입력용 — 수식으로 자동 계산)
    last_data_row = len(records) + 1
    for i in range(50):
        r = last_data_row + 1 + i
        ws.cell(row=r, column=1).number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=3).number_format = "H:MM"
        ws.cell(row=r, column=4).number_format = "H:MM"
        ws.cell(row=r, column=5).value = f'=IF(AND(C{r}<>"",D{r}<>""),D{r}-C{r},0)'
        ws.cell(row=r, column=5).number_format = "0.0"
        # 신규 행만 수식으로 일 급여 계산
        lookup_amount = master_lookup_formula(f"B{r}", f"A{r}", "C")
        ws.cell(row=r, column=6).value = (
            f'=IFERROR(IF(AND(B{r}<>"",G{r}="시급"),E{r}*24*{lookup_amount},0),0)'
        )
        ws.cell(row=r, column=6).number_format = MONEY_FMT
        lookup_type = master_lookup_formula(f"B{r}", f"A{r}", "B")
        ws.cell(row=r, column=7).value = f'=IFERROR(IF(B{r}<>"",{lookup_type},""),"")'
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c in INPUT_COLS:
                apply_cell_style(cell, fill=YELLOW_FILL)
            else:
                apply_cell_style(cell, fill=GRAY_FILL)

    # Worker dropdown
    emp_names = ",".join(e["name"] for e in EMPLOYEES)
    dv_worker = DataValidation(type="list", formula1=f'"{emp_names}"', allow_blank=True)
    ws.add_data_validation(dv_worker)
    total_rows = last_data_row + 50
    dv_worker.add(f"B2:B{total_rows}")

    ws.auto_filter.ref = f"A1:H{total_rows}"
    return ws


def build_payment_history_sheet(wb):
    """Sheet: 지급 이력 — (이름, 월, 지급여부) 룩업 테이블
    급여 요약 I열 수식이 (이름, 선택월) 조합으로 이 시트를 조회한다.
    과거 이력 수정 또는 신규 지급 반영 시 이 시트에 직접 행 추가 가능.
    """
    ws = wb.create_sheet("지급 이력")
    ws.sheet_properties.tabColor = "A5A5A5"

    headers = ["이름", "월", "지급여부"]
    widths = [12, 10, 12]

    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    apply_header_style(ws, 1, len(headers))

    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for idx, (name, ym, paid) in enumerate(PAYMENT_HISTORY):
        r = idx + 2
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=ym)
        ws.cell(row=r, column=2).number_format = "0"
        ws.cell(row=r, column=3, value=paid)
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if paid == "O":
                cell.fill = GREEN_FILL
            elif paid == "X":
                cell.fill = RED_FILL

    ws.auto_filter.ref = f"A1:C{len(PAYMENT_HISTORY) + 1}"
    return ws


def build_summary_sheet(wb):
    """Sheet 3: 급여 요약
    - 선택한 월에 활성 구간인 직원만 표시 (적용시작월~적용종료월 범위)
    - 총 급여 = 월급제면 마스터 고정액, 시급제면 해당월 근무기록 일급여 합산
    """
    ws = wb.create_sheet("급여 요약")
    ws.sheet_properties.tabColor = "ED7D31"

    widths = {1: 4, 2: 14, 3: 12, 4: 14, 5: 16, 6: 16, 7: 14, 8: 28, 9: 12, 10: 14}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # B1: 월 선택
    ws.cell(row=1, column=2, value=202502)
    ws.cell(row=1, column=2).font = Font(bold=True, size=14)
    ws.cell(row=1, column=2).number_format = "0"
    ws.cell(row=1, column=2).alignment = Alignment(horizontal="center")
    apply_cell_style(ws.cell(row=1, column=2), fill=YELLOW_FILL)

    months_list = ",".join(str(m) for m in range(202501, 202513))
    dv_month = DataValidation(type="list", formula1=f'"{months_list}"', allow_blank=False)
    ws.add_data_validation(dv_month)
    dv_month.add("B1")

    ws.cell(row=1, column=3, value="← 월 선택")
    ws.cell(row=1, column=3).font = Font(color="888888", italic=True)

    # Hidden J column: 날짜 계산 보조
    ws.cell(row=1, column=10, value='=DATE(INT(B1/100),MOD(B1,100),1)')
    ws.cell(row=1, column=10).number_format = "YYYY-MM-DD"
    ws.cell(row=2, column=10, value='=EOMONTH(J1,0)')
    ws.cell(row=2, column=10).number_format = "YYYY-MM-DD"
    ws.column_dimensions["J"].hidden = True

    # Header row 3
    headers = ["", "이름", "급여유형", "총 근무시간", "총 급여(세전)", "3.3% 원천징수", "실지급액", "입금계좌", "지급여부"]
    for i, h in enumerate(headers, 1):
        if i == 1:
            continue
        ws.cell(row=3, column=i, value=h)
    apply_header_style(ws, 3, 9)

    num_emp = len(EMPLOYEES)
    for idx, emp in enumerate(EMPLOYEES):
        r = 4 + idx
        mr = idx + 2  # master row

        # 해당 월에 재직 중이었는지 판별:
        # 입사월(F열) <= 선택월 AND (퇴사월(G열)이 비어있거나 >= 선택월)
        active_check = (
            f'AND(\'직원 마스터\'!F{mr}<>"",'
            f'\'직원 마스터\'!F{mr}<=$B$1,'
            f'OR(\'직원 마스터\'!G{mr}="",'
            f'\'직원 마스터\'!G{mr}>=$B$1))'
        )

        name_ref = f"'직원 마스터'!A{mr}"

        # 이름
        ws.cell(row=r, column=2).value = f'=IF({active_check},{name_ref},"")'
        # 급여유형
        ws.cell(row=r, column=3).value = f'=IF({active_check},\'직원 마스터\'!B{mr},"")'

        # 총 근무시간
        ws.cell(row=r, column=4).value = (
            f'=IF({active_check},'
            f'SUMPRODUCT(('
            f"'근무 기록'!B$2:B$500={name_ref})*"
            f"('근무 기록'!A$2:A$500>=J$1)*"
            f"('근무 기록'!A$2:A$500<=J$2)*"
            f"('근무 기록'!E$2:E$500))*24,\"\")"
        )
        ws.cell(row=r, column=4).number_format = "0.0"

        # 일 급여 합산 (근무기록 F열)
        daily_sum = (
            f'SUMPRODUCT(('
            f"'근무 기록'!B$2:B$500={name_ref})*"
            f"('근무 기록'!A$2:A$500>=J$1)*"
            f"('근무 기록'!A$2:A$500<=J$2)*"
            f"('근무 기록'!F$2:F$500))"
        )

        # 총 급여: 월급제면 마스터 고정액, 시급제면 해당월 일급여 합산
        # (월급제 구간에 근무기록 일급 기록이 남아있어도 월급액이 우선)
        ws.cell(row=r, column=5).value = (
            f'=IF(B{r}="","",'
            f"IF(C{r}=\"월급\",'직원 마스터'!C{mr},{daily_sum}))"
        )
        ws.cell(row=r, column=5).number_format = MONEY_FMT

        # 3.3% 원천징수
        ws.cell(row=r, column=6).value = f'=IF(E{r}="","",ROUND(E{r}*0.033,0))'
        ws.cell(row=r, column=6).number_format = MONEY_FMT

        # 실지급액
        ws.cell(row=r, column=7).value = f'=IF(E{r}="","",E{r}-F{r})'
        ws.cell(row=r, column=7).number_format = MONEY_FMT

        # 입금계좌
        ws.cell(row=r, column=8).value = f'=IF(B{r}="","",\'직원 마스터\'!D{mr})'

        # 지급여부: 지급 이력 시트에서 (이름, 선택월)로 조회, 없으면 "X"(미지급)
        paid_last_row = len(PAYMENT_HISTORY) + 51  # data rows + 50 buffer for future additions
        ws.cell(row=r, column=9).value = (
            f'=IF(B{r}="","",'
            f'IFERROR(INDEX(\'지급 이력\'!C$2:C${paid_last_row},'
            f'MATCH(1,'
            f"('지급 이력'!A$2:A${paid_last_row}=B{r})*"
            f"('지급 이력'!B$2:B${paid_last_row}=$B$1),"
            f'0)),"X"))'
        )

        # Styling
        for c in range(2, 10):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c in (4, 5, 6, 7):
                cell.fill = GRAY_FILL
            elif c == 9:
                pass
            else:
                cell.fill = YELLOW_FILL

    # 지급여부 드롭다운 / 조건부 서식
    last_emp_row = 3 + num_emp
    dv_paid = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
    ws.add_data_validation(dv_paid)
    dv_paid.add(f"I4:I{last_emp_row}")

    ws.conditional_formatting.add(
        f"I4:I{last_emp_row}",
        CellIsRule(operator="equal", formula=['"O"'], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"I4:I{last_emp_row}",
        CellIsRule(operator="equal", formula=['"X"'], fill=RED_FILL),
    )

    # 하단 요약
    summary_row = last_emp_row + 2
    ws.cell(row=summary_row, column=2, value="합계")
    ws.cell(row=summary_row, column=2).font = Font(bold=True, size=11)
    ws.cell(row=summary_row, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=summary_row, column=2).border = THIN_BORDER

    ws.cell(row=summary_row, column=5).value = f"=SUM(E4:E{last_emp_row})"
    ws.cell(row=summary_row, column=5).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=5).font = Font(bold=True)
    ws.cell(row=summary_row, column=5).border = THIN_BORDER

    ws.cell(row=summary_row, column=6).value = f"=SUM(F4:F{last_emp_row})"
    ws.cell(row=summary_row, column=6).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=6).font = Font(bold=True)
    ws.cell(row=summary_row, column=6).border = THIN_BORDER

    ws.cell(row=summary_row, column=7).value = f"=SUM(G4:G{last_emp_row})"
    ws.cell(row=summary_row, column=7).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=7).font = Font(bold=True)
    ws.cell(row=summary_row, column=7).border = THIN_BORDER

    label_row = summary_row + 1
    ws.merge_cells(start_row=label_row, start_column=2, end_row=label_row, end_column=4)
    ws.cell(row=label_row, column=2, value="당월 총 인건비")
    ws.cell(row=label_row, column=2).font = Font(bold=True, size=13, color="1F4E79")
    ws.cell(row=label_row, column=2).alignment = Alignment(horizontal="right")

    ws.cell(row=label_row, column=5).value = f"=E{summary_row}"
    ws.cell(row=label_row, column=5).number_format = MONEY_FMT
    ws.cell(row=label_row, column=5).font = Font(bold=True, size=13, color="1F4E79")
    ws.cell(row=label_row, column=5).border = THIN_BORDER

    return ws


def main():
    print("분트 인건비 관리 엑셀 생성 시작...")

    records = extract_data()
    print(f"  → 근무 기록 {len(records)}건 추출 완료")

    workers = sorted(set(r["worker"] for r in records))
    print(f"  → 근무자: {', '.join(workers)}")

    if records:
        print(f"  → 기간: {records[0]['date'].strftime('%Y-%m-%d')} ~ {records[-1]['date'].strftime('%Y-%m-%d')}")

    wb = Workbook()

    build_master_sheet(wb)
    print("  → [직원 마스터] 시트 생성 완료")

    build_records_sheet(wb, records)
    print("  → [근무 기록] 시트 생성 완료")

    build_payment_history_sheet(wb)
    print("  → [지급 이력] 시트 생성 완료")

    build_summary_sheet(wb)
    print("  → [급여 요약] 시트 생성 완료")

    wb.save(OUTPUT_PATH)
    print(f"\n[완료] 저장: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
