"""
분트 미술학원 인건비 관리 엑셀 생성 스크립트
- 원본: 분트 재무제표(25.12월).xlsb → "2. 인건비 계산" 시트
- 출력: 분트_인건비관리.xlsx (3개 시트)
"""

import os
from datetime import datetime, timedelta
from pyxlsb import open_workbook as open_xlsb
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ── Constants ──────────────────────────────────────────────────────────
XLSB_PATH = os.path.join(os.path.dirname(__file__), "분트 재무제표(25.12월).xlsb")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "분트_인건비관리.xlsx")
SHEET_NAME = "2. 인건비 계산"

EMPLOYEES = [
    {"name": "고경민", "pay_type": "월급", "amount": 1200000, "account": "국민은행 82240104164295", "status": "재직"},
    {"name": "장예원", "pay_type": "시급", "amount": 12000, "account": "카카오뱅크 3333-07-7072641", "status": "퇴직"},
    {"name": "유은비", "pay_type": "시급", "amount": 12000, "account": "우리 1002-533-898534", "status": "퇴직"},
    {"name": "전민아", "pay_type": "시급", "amount": 12000, "account": "하나은행 558-910330-30707", "status": "재직"},
    {"name": "김채현", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-351-542250", "status": "재직"},
    {"name": "이진화", "pay_type": "시급", "amount": 12000, "account": "우리은행 1002-166-678383", "status": "재직"},
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
            start_time = row[3].v  # D column
            end_time = row[4].v    # E column
            hours = row[5].v       # F column
            daily_pay = row[6].v   # G column

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


def build_master_sheet(wb):
    """Sheet 1: 직원 마스터"""
    ws = wb.active
    ws.title = "직원 마스터"
    ws.sheet_properties.tabColor = "4472C4"

    headers = ["이름", "급여유형", "금액", "입금계좌", "재직상태", "입사일", "비고"]
    col_widths = [12, 12, 15, 30, 12, 14, 20]

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
    """Sheet 2: 근무 기록"""
    ws = wb.create_sheet("근무 기록")
    ws.sheet_properties.tabColor = "70AD47"

    headers = ["날짜", "근무자", "시작시간", "종료시간", "근무시간", "일 급여", "급여유형", "비고"]
    col_widths = [14, 12, 12, 12, 12, 15, 12, 20]

    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    apply_header_style(ws, 1, len(headers))

    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Input columns: A(날짜), B(근무자), C(시작시간), D(종료시간), H(비고) → yellow
    # Formula columns: E(근무시간), F(일 급여), G(급여유형) → gray
    INPUT_COLS = {1, 2, 3, 4, 8}
    FORMULA_COLS = {5, 6, 7}

    # Write extracted data
    for idx, rec in enumerate(records):
        r = idx + 2
        ws.cell(row=r, column=1, value=rec["date"])
        ws.cell(row=r, column=1).number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=2, value=rec["worker"])

        if rec["start"] is not None:
            # Store as decimal hours (e.g. 14.5 = 14:30)
            # Convert to Excel time fraction for display
            h_start = float(rec["start"])
            ws.cell(row=r, column=3, value=h_start / 24)
            ws.cell(row=r, column=3).number_format = "H:MM"
        if rec["end"] is not None:
            h_end = float(rec["end"])
            ws.cell(row=r, column=4, value=h_end / 24)
            ws.cell(row=r, column=4).number_format = "H:MM"

        # Formula: 근무시간
        ws.cell(row=r, column=5).value = f'=IF(AND(C{r}<>"",D{r}<>""),D{r}-C{r},0)'
        ws.cell(row=r, column=5).number_format = "0.0"
        # Formula: 일 급여
        ws.cell(row=r, column=6).value = (
            f'=IF(G{r}="시급",E{r}*24*VLOOKUP(B{r},\'직원 마스터\'!A:C,3,FALSE),0)'
        )
        ws.cell(row=r, column=6).number_format = MONEY_FMT
        # Formula: 급여유형
        ws.cell(row=r, column=7).value = (
            f'=IF(B{r}<>"",VLOOKUP(B{r},\'직원 마스터\'!A:B,2,FALSE),"")'
        )

        # Styling
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c in INPUT_COLS:
                apply_cell_style(cell, fill=YELLOW_FILL)
            else:
                apply_cell_style(cell, fill=GRAY_FILL)

    # Add 50 empty rows with formulas
    last_data_row = len(records) + 1
    for i in range(50):
        r = last_data_row + 1 + i
        ws.cell(row=r, column=1).number_format = "YYYY-MM-DD"
        if r > 1:
            ws.cell(row=r, column=3).number_format = "H:MM"
            ws.cell(row=r, column=4).number_format = "H:MM"
        ws.cell(row=r, column=5).value = f'=IF(AND(C{r}<>"",D{r}<>""),D{r}-C{r},0)'
        ws.cell(row=r, column=5).number_format = "0.0"
        ws.cell(row=r, column=6).value = (
            f'=IF(G{r}="시급",E{r}*24*VLOOKUP(B{r},\'직원 마스터\'!A:C,3,FALSE),0)'
        )
        ws.cell(row=r, column=6).number_format = MONEY_FMT
        ws.cell(row=r, column=7).value = (
            f'=IF(B{r}<>"",VLOOKUP(B{r},\'직원 마스터\'!A:B,2,FALSE),"")'
        )
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c in INPUT_COLS:
                apply_cell_style(cell, fill=YELLOW_FILL)
            else:
                apply_cell_style(cell, fill=GRAY_FILL)

    # Worker dropdown from master sheet names
    emp_names = ",".join(e["name"] for e in EMPLOYEES)
    dv_worker = DataValidation(type="list", formula1=f'"{emp_names}"', allow_blank=True)
    ws.add_data_validation(dv_worker)
    total_rows = last_data_row + 50
    dv_worker.add(f"B2:B{total_rows}")

    # Auto-filter
    ws.auto_filter.ref = f"A1:H{total_rows}"

    return ws


def build_summary_sheet(wb):
    """Sheet 3: 급여 요약"""
    ws = wb.create_sheet("급여 요약")
    ws.sheet_properties.tabColor = "ED7D31"

    # Column widths
    widths = {1: 4, 2: 14, 3: 12, 4: 14, 5: 16, 6: 16, 7: 14, 8: 28, 9: 12, 10: 14}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # B1: Month selector
    ws.cell(row=1, column=2, value=202501)
    ws.cell(row=1, column=2).font = Font(bold=True, size=14)
    ws.cell(row=1, column=2).number_format = "0"
    ws.cell(row=1, column=2).alignment = Alignment(horizontal="center")

    months_list = ",".join(str(m) for m in range(202501, 202513))
    dv_month = DataValidation(type="list", formula1=f'"{months_list}"', allow_blank=False)
    ws.add_data_validation(dv_month)
    dv_month.add("B1")

    ws.cell(row=1, column=3, value="← 월 선택")
    ws.cell(row=1, column=3).font = Font(color="888888", italic=True)

    # Hidden J column: helper dates
    # J1 = start of month, J2 = end of month
    ws.cell(row=1, column=10, value='=DATE(INT(B1/100),MOD(B1,100),1)')
    ws.cell(row=1, column=10).number_format = "YYYY-MM-DD"
    ws.cell(row=2, column=10, value='=EOMONTH(J1,0)')
    ws.cell(row=2, column=10).number_format = "YYYY-MM-DD"
    ws.column_dimensions["J"].hidden = True

    # Header row 3
    headers = ["", "이름", "급여유형", "총 근무시간", "총 급여(세전)", "3.3% 원천징수", "실지급액", "입금계좌", "지급여부"]
    for i, h in enumerate(headers, 1):
        if i == 1:
            continue  # column A empty
        ws.cell(row=3, column=i, value=h)
    apply_header_style(ws, 3, 9)

    # Employee rows (rows 4+)
    num_emp = len(EMPLOYEES)
    for idx, emp in enumerate(EMPLOYEES):
        r = 4 + idx
        master_row = idx + 2  # row in master sheet

        # 이름: reference master
        ws.cell(row=r, column=2, value=f"='직원 마스터'!A{master_row}")
        # 급여유형: reference master
        ws.cell(row=r, column=3, value=f"='직원 마스터'!B{master_row}")

        # 총 근무시간: SUMIFS (date as excel time fraction, multiply by 24)
        ws.cell(row=r, column=4).value = (
            f'=SUMPRODUCT(('
            f"'근무 기록'!B$2:B$500=B{r})*"
            f"('근무 기록'!A$2:A$500>=J$1)*"
            f"('근무 기록'!A$2:A$500<=J$2)*"
            f"('근무 기록'!E$2:E$500))*24"
        )
        ws.cell(row=r, column=4).number_format = "0.0"

        # 총 급여(세전): monthly=fixed, hourly=hours*rate
        ws.cell(row=r, column=5).value = (
            f'=IF(C{r}="월급",'
            f"'직원 마스터'!C{master_row},"
            f"D{r}*'직원 마스터'!C{master_row})"
        )
        ws.cell(row=r, column=5).number_format = MONEY_FMT

        # 3.3% 원천징수
        ws.cell(row=r, column=6).value = f"=ROUND(E{r}*0.033,0)"
        ws.cell(row=r, column=6).number_format = MONEY_FMT

        # 실지급액
        ws.cell(row=r, column=7).value = f"=E{r}-F{r}"
        ws.cell(row=r, column=7).number_format = MONEY_FMT

        # 입금계좌: reference master
        ws.cell(row=r, column=8, value=f"='직원 마스터'!D{master_row}")

        # 지급여부: dropdown
        ws.cell(row=r, column=9, value="X")

        # Styling
        for c in range(2, 10):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c in (4, 5, 6, 7):  # calculated
                cell.fill = GRAY_FILL
            elif c == 9:
                pass  # conditional formatting below
            else:
                cell.fill = YELLOW_FILL

    # 지급여부 dropdown O/X
    dv_paid = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
    ws.add_data_validation(dv_paid)
    last_emp_row = 3 + num_emp
    dv_paid.add(f"I4:I{last_emp_row}")

    # Conditional formatting for 지급여부
    ws.conditional_formatting.add(
        f"I4:I{last_emp_row}",
        CellIsRule(operator="equal", formula=['"O"'], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"I4:I{last_emp_row}",
        CellIsRule(operator="equal", formula=['"X"'], fill=RED_FILL),
    )

    # Summary rows
    summary_row = last_emp_row + 2
    ws.cell(row=summary_row, column=2, value="합계")
    ws.cell(row=summary_row, column=2).font = Font(bold=True, size=11)
    ws.cell(row=summary_row, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=summary_row, column=2).border = THIN_BORDER

    # 총 급여 합계
    ws.cell(row=summary_row, column=5).value = f"=SUM(E4:E{last_emp_row})"
    ws.cell(row=summary_row, column=5).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=5).font = Font(bold=True)
    ws.cell(row=summary_row, column=5).border = THIN_BORDER

    # 원천징수 합계
    ws.cell(row=summary_row, column=6).value = f"=SUM(F4:F{last_emp_row})"
    ws.cell(row=summary_row, column=6).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=6).font = Font(bold=True)
    ws.cell(row=summary_row, column=6).border = THIN_BORDER

    # 실지급 합계
    ws.cell(row=summary_row, column=7).value = f"=SUM(G4:G{last_emp_row})"
    ws.cell(row=summary_row, column=7).number_format = MONEY_FMT
    ws.cell(row=summary_row, column=7).font = Font(bold=True)
    ws.cell(row=summary_row, column=7).border = THIN_BORDER

    # "당월 총 인건비" label
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

    # 1. Extract data
    records = extract_data()
    print(f"  → 근무 기록 {len(records)}건 추출 완료")

    # Unique workers found
    workers = sorted(set(r["worker"] for r in records))
    print(f"  → 근무자: {', '.join(workers)}")

    # Date range
    if records:
        print(f"  → 기간: {records[0]['date'].strftime('%Y-%m-%d')} ~ {records[-1]['date'].strftime('%Y-%m-%d')}")

    # 2. Build workbook
    wb = Workbook()

    build_master_sheet(wb)
    print("  → [직원 마스터] 시트 생성 완료")

    build_records_sheet(wb, records)
    print("  → [근무 기록] 시트 생성 완료")

    build_summary_sheet(wb)
    print("  → [급여 요약] 시트 생성 완료")

    # 3. Save
    wb.save(OUTPUT_PATH)
    print(f"\n[완료] 저장: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
