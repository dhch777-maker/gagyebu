"""분트 인건비 관리 v3 — VBA 기반 동적 급여요약

v2와 데이터 구조는 동일. 다만 급여 요약 탭은 수식/마스터 복제 대신
VBA Worksheet_Change 이벤트로 월 선택 시 활성 직원만 순차 표시.

빌드 순서:
1. openpyxl로 xlsx 기본 생성 (급여 요약은 헤더·월선택만, 데이터 행 없음)
2. pywin32 COM으로 Excel 열어 VBA 주입 → .xlsm 저장
"""
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# v2 모듈에서 공통 구성요소 재사용
from build_payroll import (
    EMPLOYEES, PAYMENT_HISTORY,
    build_master_sheet, build_records_sheet, build_payment_history_sheet,
    apply_header_style, apply_cell_style,
    NAVY_FILL, WHITE_BOLD, YELLOW_FILL, GRAY_FILL, THIN_BORDER,
    GREEN_FILL, RED_FILL, MONEY_FMT,
    extract_data,
)

OUT_XLSX = os.path.join(os.path.dirname(__file__), "분트_인건비관리_v3.xlsx")
OUT_XLSM = os.path.join(os.path.dirname(__file__), "분트_인건비관리_v3.xlsm")

# ── VBA 코드 ────────────────────────────────────────────────────────
# 이 코드는 "급여 요약" 시트 모듈에 주입됨.
# B1(월 선택) 변경 시 Worksheet_Change 발동 → RefreshSummary 호출
# 직원 마스터에서 활성 구간만 추출, 지급 이력 기반 총급여·지급여부 기입.
VBA_CODE = r"""
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B1")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Cleanup
    Call RefreshSummary
Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshSummary()
    Dim wsSum As Worksheet: Set wsSum = Me
    Dim wsMaster As Worksheet: Set wsMaster = ThisWorkbook.Worksheets("직원 마스터")
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Worksheets("근무 기록")
    Dim wsPaid As Worksheet: Set wsPaid = ThisWorkbook.Worksheets("지급 이력")

    Dim selMonthRaw As Variant: selMonthRaw = wsSum.Range("B1").Value
    If Not IsNumeric(selMonthRaw) Then Exit Sub
    Dim sm As Long: sm = CLng(selMonthRaw)
    If sm < 200001 Or sm > 209912 Then Exit Sub

    Dim yr As Long: yr = sm \ 100
    Dim mo As Long: mo = sm Mod 100
    If mo < 1 Or mo > 12 Then Exit Sub
    Dim d1 As Date: d1 = DateSerial(yr, mo, 1)
    Dim d2 As Date: d2 = DateSerial(yr, mo + 1, 0)

    ' 기존 데이터·합계 영역 초기화 (병합 해제 + 값 삭제 + 서식 리셋)
    On Error Resume Next
    wsSum.Range("B4:I60").UnMerge
    On Error GoTo 0
    wsSum.Range("B4:I60").ClearContents
    wsSum.Range("B4:I60").Interior.ColorIndex = xlNone
    wsSum.Range("B4:I60").Borders.LineStyle = xlNone

    Dim outRow As Long: outRow = 4
    Dim mRow As Long, mLast As Long
    mLast = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row

    Dim totGross As Double, totTax As Double, totNet As Double
    totGross = 0: totTax = 0: totNet = 0

    For mRow = 2 To mLast
        Dim empName As String
        empName = Trim(CStr(wsMaster.Cells(mRow, 1).Value))
        If empName = "" Then GoTo NextIter

        Dim mStart As Variant, mEnd As Variant
        mStart = wsMaster.Cells(mRow, 6).Value
        mEnd = wsMaster.Cells(mRow, 7).Value
        If IsEmpty(mStart) Or Not IsNumeric(mStart) Then GoTo NextIter
        If CLng(mStart) > sm Then GoTo NextIter
        ' mEnd가 빈 셀이면 "현재까지" 의미 — skip 조건 검사 안 함
        ' 주의: IsNumeric(Empty)=True 이므로 IsEmpty 먼저 확인 필요
        If Not IsEmpty(mEnd) And IsNumeric(mEnd) Then
            If CLng(mEnd) < sm Then GoTo NextIter
        End If

        Dim payType As String: payType = CStr(wsMaster.Cells(mRow, 2).Value)
        Dim masterAmt As Double
        If IsNumeric(wsMaster.Cells(mRow, 3).Value) Then
            masterAmt = CDbl(wsMaster.Cells(mRow, 3).Value)
        Else
            masterAmt = 0
        End If
        Dim account As String: account = CStr(wsMaster.Cells(mRow, 4).Value)

        ' 근무기록 집계 (근무시간 + 일급여 합)
        Dim totalHours As Double: totalHours = 0
        Dim dailySum As Double: dailySum = 0
        Dim rRow As Long, rLast As Long
        rLast = wsRec.Cells(wsRec.Rows.Count, 1).End(xlUp).Row
        For rRow = 2 To rLast
            Dim dtVal As Variant: dtVal = wsRec.Cells(rRow, 1).Value
            If IsDate(dtVal) Then
                If CDate(dtVal) >= d1 And CDate(dtVal) <= d2 Then
                    If Trim(CStr(wsRec.Cells(rRow, 2).Value)) = empName Then
                        Dim hVal As Variant: hVal = wsRec.Cells(rRow, 5).Value
                        If IsNumeric(hVal) Then totalHours = totalHours + CDbl(hVal) * 24
                        Dim pVal As Variant: pVal = wsRec.Cells(rRow, 6).Value
                        If IsNumeric(pVal) Then dailySum = dailySum + CDbl(pVal)
                    End If
                End If
            End If
        Next rRow

        ' 지급 이력 조회
        Dim paidAmt As Variant: paidAmt = LookupPaid(wsPaid, empName, sm, 4)
        Dim paidStat As Variant: paidStat = LookupPaid(wsPaid, empName, sm, 3)

        ' 총급여: 지급이력 우선, 없으면 월급제=마스터액 / 시급제=일급합
        Dim gross As Double
        If IsNumeric(paidAmt) Then
            gross = CDbl(paidAmt)
        ElseIf payType = "월급" Then
            gross = masterAmt
        Else
            gross = dailySum
        End If

        Dim tax As Double: tax = Application.WorksheetFunction.Round(gross * 0.033, 0)
        Dim net As Double: net = gross - tax
        Dim statusStr As String
        If IsEmpty(paidStat) Then
            statusStr = "X"
        Else
            statusStr = CStr(paidStat)
        End If

        ' 행 기입
        wsSum.Cells(outRow, 2).Value = empName
        wsSum.Cells(outRow, 3).Value = payType
        wsSum.Cells(outRow, 4).Value = totalHours
        wsSum.Cells(outRow, 5).Value = gross
        wsSum.Cells(outRow, 6).Value = tax
        wsSum.Cells(outRow, 7).Value = net
        wsSum.Cells(outRow, 8).Value = account
        wsSum.Cells(outRow, 9).Value = statusStr

        ' 서식
        Call ApplyRowStyle(wsSum, outRow)
        If statusStr = "O" Then
            wsSum.Cells(outRow, 9).Interior.Color = RGB(198, 239, 206)
        ElseIf statusStr = "X" Then
            wsSum.Cells(outRow, 9).Interior.Color = RGB(255, 199, 206)
        End If

        totGross = totGross + gross
        totTax = totTax + tax
        totNet = totNet + net
        outRow = outRow + 1
NextIter:
    Next mRow

    ' 합계 행
    If outRow > 4 Then
        Dim sumRow As Long: sumRow = outRow + 1
        wsSum.Cells(sumRow, 2).Value = "합계"
        wsSum.Cells(sumRow, 2).Font.Bold = True
        wsSum.Cells(sumRow, 2).HorizontalAlignment = xlCenter
        wsSum.Cells(sumRow, 5).Value = totGross
        wsSum.Cells(sumRow, 6).Value = totTax
        wsSum.Cells(sumRow, 7).Value = totNet
        Dim c As Long
        For c = 5 To 7
            wsSum.Cells(sumRow, c).NumberFormat = "#,##0"
            wsSum.Cells(sumRow, c).Font.Bold = True
            wsSum.Cells(sumRow, c).Borders.LineStyle = xlContinuous
        Next c
        wsSum.Cells(sumRow, 2).Borders.LineStyle = xlContinuous

        ' 당월 총 인건비 강조 행
        Dim labelRow As Long: labelRow = sumRow + 1
        wsSum.Range(wsSum.Cells(labelRow, 2), wsSum.Cells(labelRow, 4)).Merge
        wsSum.Cells(labelRow, 2).Value = "당월 총 인건비"
        wsSum.Cells(labelRow, 2).Font.Bold = True
        wsSum.Cells(labelRow, 2).Font.Size = 13
        wsSum.Cells(labelRow, 2).Font.Color = RGB(31, 78, 121)
        wsSum.Cells(labelRow, 2).HorizontalAlignment = xlRight
        wsSum.Cells(labelRow, 5).Value = totGross
        wsSum.Cells(labelRow, 5).NumberFormat = "#,##0"
        wsSum.Cells(labelRow, 5).Font.Bold = True
        wsSum.Cells(labelRow, 5).Font.Size = 13
        wsSum.Cells(labelRow, 5).Font.Color = RGB(31, 78, 121)
        wsSum.Cells(labelRow, 5).Borders.LineStyle = xlContinuous
    End If
End Sub

Private Sub ApplyRowStyle(ws As Worksheet, r As Long)
    Dim c As Long
    For c = 2 To 9
        ws.Cells(r, c).Borders.LineStyle = xlContinuous
        ws.Cells(r, c).HorizontalAlignment = xlCenter
        ws.Cells(r, c).VerticalAlignment = xlCenter
    Next c
    ws.Cells(r, 4).NumberFormat = "0.0"
    ws.Cells(r, 5).NumberFormat = "#,##0"
    ws.Cells(r, 6).NumberFormat = "#,##0"
    ws.Cells(r, 7).NumberFormat = "#,##0"
    ' 계산 셀 회색 배경
    ws.Cells(r, 4).Interior.Color = RGB(242, 242, 242)
    ws.Cells(r, 5).Interior.Color = RGB(242, 242, 242)
    ws.Cells(r, 6).Interior.Color = RGB(242, 242, 242)
    ws.Cells(r, 7).Interior.Color = RGB(242, 242, 242)
End Sub

Private Function LookupPaid(ws As Worksheet, name As String, ym As Long, col As Long) As Variant
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lr
        If Trim(CStr(ws.Cells(r, 1).Value)) = name Then
            If IsNumeric(ws.Cells(r, 2).Value) Then
                If CLng(ws.Cells(r, 2).Value) = ym Then
                    LookupPaid = ws.Cells(r, col).Value
                    Exit Function
                End If
            End If
        End If
    Next r
    LookupPaid = Empty
End Function
"""


def build_summary_sheet_v3(wb):
    """급여 요약 v3: 헤더 + 월 선택만, 데이터 행은 VBA가 동적 생성."""
    ws = wb.create_sheet("급여 요약")
    ws.sheet_properties.tabColor = "ED7D31"

    widths = {1: 4, 2: 14, 3: 12, 4: 14, 5: 16, 6: 16, 7: 14, 8: 28, 9: 12}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # B1: 월 선택
    ws.cell(row=1, column=2, value=202502)
    ws.cell(row=1, column=2).font = Font(bold=True, size=14)
    ws.cell(row=1, column=2).number_format = "0"
    ws.cell(row=1, column=2).alignment = Alignment(horizontal="center")
    apply_cell_style(ws.cell(row=1, column=2), fill=YELLOW_FILL)

    months_list = ",".join(str(m) for m in range(202501, 202613))
    dv_month = DataValidation(type="list", formula1=f'"{months_list}"', allow_blank=False)
    ws.add_data_validation(dv_month)
    dv_month.add("B1")

    ws.cell(row=1, column=3, value="← 월 선택 (변경 시 VBA 자동 갱신)")
    ws.cell(row=1, column=3).font = Font(color="888888", italic=True)

    # 헤더 3행
    headers = ["", "이름", "급여유형", "총 근무시간", "총 급여(세전)",
               "3.3% 원천징수", "실지급액", "입금계좌", "지급여부"]
    for i, h in enumerate(headers, 1):
        if i == 1:
            continue
        ws.cell(row=3, column=i, value=h)
    apply_header_style(ws, 3, 9)

    # 안내 메시지 (병합 없이 K1에만 배치 — 데이터 영역 침범 금지)
    ws.cell(row=2, column=2, value="※ B1 월 선택 시 급여 요약 자동 갱신됨")
    ws.cell(row=2, column=2).font = Font(italic=True, color="888888")

    return ws


def inject_vba(xlsx_path, xlsm_path):
    """COM으로 Excel 열어 급여요약 시트 모듈에 VBA 주입 후 .xlsm 저장"""
    import win32com.client as win32

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

        # VBComponents에서 Properties("Name")로 "급여 요약" 시트 모듈 찾기
        # (openpyxl 생성 xlsx는 Worksheet.CodeName이 비어있어 COM 매칭 불가)
        vbproj = wb.VBProject
        target_comp = None
        for i in range(1, vbproj.VBComponents.Count + 1):
            comp = vbproj.VBComponents.Item(i)
            if comp.Type != 100:  # 100 = vbext_ct_Document (Sheet/Workbook module)
                continue
            try:
                if comp.Properties("Name").Value == "급여 요약":
                    target_comp = comp
                    break
            except Exception:
                continue
        if target_comp is None:
            raise RuntimeError("'급여 요약' 시트 VB 모듈을 찾을 수 없음")

        target_comp.CodeModule.AddFromString(VBA_CODE)

        # xlsm 저장 (FileFormat 52 = xlOpenXMLWorkbookMacroEnabled)
        if os.path.exists(xlsm_path):
            os.remove(xlsm_path)
        wb.SaveAs(os.path.abspath(xlsm_path), FileFormat=52)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def main():
    print("v3 (VBA 동적 급여요약) 빌드 시작...")
    records = extract_data()
    print(f"  → 근무 기록 {len(records)}건 추출")

    wb = Workbook()
    build_master_sheet(wb)
    print("  → [직원 마스터]")
    build_records_sheet(wb, records)
    print("  → [근무 기록]")
    build_payment_history_sheet(wb)
    print("  → [지급 이력]")
    build_summary_sheet_v3(wb)
    print("  → [급여 요약] (빈 템플릿)")

    wb.save(OUT_XLSX)
    print(f"  중간 xlsx 저장: {OUT_XLSX}")

    print("  VBA 주입 중...")
    inject_vba(OUT_XLSX, OUT_XLSM)

    if os.path.exists(OUT_XLSM) and os.path.exists(OUT_XLSX):
        os.remove(OUT_XLSX)

    print(f"\n[완료] {OUT_XLSM}")
    print("※ Excel에서 열 때 '매크로 사용' 허용 필요")
    print("※ B1에서 월 선택 시 자동으로 활성 직원만 순차 표시됨")


if __name__ == "__main__":
    main()
