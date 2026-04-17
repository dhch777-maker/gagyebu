import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import date, timedelta
import calendar

# === 설정 ===
BABY_NAME = "동해"
BABY_BIRTH_DATE = date(2025, 10, 28)
MEAL_START_DATE = date(2026, 4, 19)
OUTPUT_FILE = "EUSIK/동해_초기이유식_캘린더_v2.xlsx"

# === 28일 식단 데이터 ===
meal_plan = {
    1:  {"밥": "쌀죽과 퀵오트밀(10→7배죽)\n(잡곡 50%, 미음 대신 죽으로)", "고기류": "", "반찬": ""},
    2:  {"밥": "쌀죽과 퀵오트밀\n(서서히 늘려서 1~2주에\n50% 잡곡밥도 가능)", "고기류": "", "반찬": ""},
    3:  {"밥": "쌀오트밀 죽\n+ 소고기 다짐\n(얇게 썰어줘도 됨)", "고기류": "", "반찬": ""},
    4:  {"밥": "쌀오트밀 죽\n(빠르면 1배죽→7배죽)", "고기류": "소고기", "반찬": ""},
    5:  {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "+청경채\n(스틱 형태/토핑/죽에 섞기)\n(또는 청경채+당근)"},
    6:  {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "청경채\n(또는 청경채+당근)"},
    7:  {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "청경채(또는 청경채+당근)\n+사과 익힌 것\n(퓌레/슬라이스)\n(사과+바나나도 가능)"},
    8:  {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "청경채(또는 청경채+당근)\n사과 익힌 것\n(사과+바나나)"},
    9:  {"밥": "쌀오트밀 죽", "고기류": "소고기\n+계란 완전히 익힌 것", "반찬": "청경채(또는 청경채+당근)\n사과 익힌 것\n(사과+바나나)"},
    10: {"밥": "쌀오트밀 죽", "고기류": "소고기\n(완자로 줘도 됨)", "반찬": "청경채(또는 청경채+당근)\n사과 익힌 것\n(사과+바나나)"},
    11: {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "청경채(또는 청경채+당근)\n사과 익힌 것\n(사과+바나나)"},
    12: {"밥": "쌀오트밀 죽", "고기류": "소고기\n+계란", "반찬": "+시금치(또는 시금치+단호박)\n사과"},
    13: {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "+양배추\n시금치 단호박"},
    14: {"밥": "쌀오트밀 죽\n+ 밀가루 첨가", "고기류": "소고기", "반찬": "청경채+당근\n(또는 시금치+단호박)\n사과"},
    15: {"밥": "쌀오트밀 죽", "고기류": "소고기, 계란", "반찬": "청경채+당근\n(또는 시금치+단호박)\n사과"},
    16: {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "청경채+당근\n(또는 시금치+단호박)\n사과"},
    17: {"밥": "쌀오트밀 죽 + 밀가루", "고기류": "소고기", "반찬": "양배추, 청경채, 당근\n+토마토\n사과"},
    18: {"밥": "쌀오트밀 죽", "고기류": "소고기, 계란", "반찬": "양배추, 청경채, 당근\n토마토\n사과"},
    19: {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "단호박, 시금치\n+땅콩 소스(1주일에 3회)\n토마토, 사과"},
    20: {"밥": "쌀오트밀 죽 + 밀가루", "고기류": "소고기", "반찬": "단호박, 시금치\n또는 청경채, 당근\n토마토, 사과"},
    21: {"밥": "쌀오트밀 죽", "고기류": "소고기, 계란", "반찬": "당근, 시금치\n땅콩 소스\n토마토, 사과"},
    22: {"밥": "쌀오트밀 죽", "고기류": "소고기\n+생선(1주일에 2회)", "반찬": "당근, 시금치\n토마토, 사과"},
    23: {"밥": "쌀오트밀 죽 + 밀가루", "고기류": "소고기, 계란", "반찬": "단호박, 시금치\n땅콩 소스\n토마토, 사과"},
    24: {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "당근, 시금치\n토마토, 사과"},
    25: {"밥": "쌀오트밀 죽", "고기류": "소고기, 생선", "반찬": "단호박, 시금치\n+또는 아보카도+딸기"},
    26: {"밥": "쌀오트밀 죽 + 밀가루", "고기류": "소고기, 계란", "반찬": "당근, 시금치\n땅콩 소스\n또는 아보카도, 딸기"},
    27: {"밥": "쌀오트밀 죽", "고기류": "+돼지고기 또는 닭고기", "반찬": "당근, 시금치\n또는 아보카도, 딸기"},
    28: {"밥": "쌀오트밀 죽", "고기류": "소고기", "반찬": "당근, 시금치\n땅콩 소스\n또는 아보카도, 딸기"},
}

# === 동해바다 테마 색상 ===
# 깊은 바다 (헤더)
DEEP_SEA = "1B4F72"
DEEP_SEA_LIGHT = "2471A3"
# 바다 (요일 헤더)
OCEAN = "2E86C1"
# 날짜 행 배경
DATE_ROW_COLOR = "D4E6F1"       # 하늘빛 연한 파랑
DATE_ROW_WEEKEND_COLOR = "AED6F1"  # 조금 더 진한 하늘
# 주말 내용 배경
WEEKEND_BG = "EBF5FB"
# 이유식 기간 외
EMPTY_BG = "F2F4F4"
# 주차별 내용 행 배경 (파도 그라데이션)
WAVE_COLORS = ["F0F9FF", "E8F4FD", "E0F0FB", "D6ECF8", "CEEAF6"]
# 모래사장 (참고사항)
SAND = "F9E79F"
SAND_LIGHT = "FEF9E7"
# 테두리
BORDER_COLOR = "85C1E9"
BORDER_LIGHT = "AED6F1"

# 스타일 객체
THIN_BORDER = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="thin", color=BORDER_COLOR),
    bottom=Side(style="thin", color=BORDER_COLOR),
)
DATE_BORDER = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="thin", color=BORDER_COLOR),
    bottom=Side(style="hair", color=BORDER_LIGHT),
)
CONTENT_BORDER = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="hair", color=BORDER_LIGHT),
    bottom=Side(style="thin", color=BORDER_COLOR),
)
HEADER_BORDER = Border(
    left=Side(style="medium", color=DEEP_SEA),
    right=Side(style="medium", color=DEEP_SEA),
    top=Side(style="medium", color=DEEP_SEA),
    bottom=Side(style="medium", color=DEEP_SEA),
)

DOW_KR = ["일", "월", "화", "수", "목", "금", "토"]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = f"{BABY_NAME} 이유식 캘린더"

# 인쇄 설정
ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_setup.orientation = "landscape"

# 열 너비
for col in range(1, 8):
    ws.column_dimensions[get_column_letter(col)].width = 22

# === 상단 타이틀 ===
ws.merge_cells("A1:G1")
cell = ws["A1"]
cell.value = f"🌊 {BABY_NAME}의 초기 이유식 28일 항해 🐳"
cell.font = Font(name="맑은 고딕", size=16, bold=True, color=DEEP_SEA)
cell.fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
cell.alignment = Alignment(horizontal="center", vertical="center")
for c in range(1, 8):
    ws.cell(row=1, column=c).fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
ws.row_dimensions[1].height = 45

# === 상단 정보 ===
ws.merge_cells("A2:C2")
ws["A2"].value = f"🐣 {BABY_NAME} 탄생일: {BABY_BIRTH_DATE.strftime('%Y-%m-%d')}"
ws["A2"].font = Font(name="맑은 고딕", size=10, bold=True, color=DEEP_SEA_LIGHT)
ws["A2"].fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

ws.merge_cells("D2:G2")
d_start = (MEAL_START_DATE - BABY_BIRTH_DATE).days
ws["D2"].value = f"🥄 이유식 시작일: {MEAL_START_DATE.strftime('%Y-%m-%d')}  (D+{d_start})"
ws["D2"].font = Font(name="맑은 고딕", size=10, bold=True, color=DEEP_SEA_LIGHT)
ws["D2"].fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
ws["D2"].alignment = Alignment(horizontal="left", vertical="center")

for c in range(1, 8):
    ws.cell(row=2, column=c).fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
ws.row_dimensions[2].height = 28

# 물결 구분선
ws.merge_cells("A3:G3")
ws["A3"].value = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
ws["A3"].font = Font(name="맑은 고딕", size=8, color=OCEAN)
ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
ws["A3"].fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
for c in range(1, 8):
    ws.cell(row=3, column=c).fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
ws.row_dimensions[3].height = 14

# === 캘린더 생성 ===
end_date = MEAL_START_DATE + timedelta(days=27)

months = []
current = MEAL_START_DATE.replace(day=1)
while current <= end_date:
    months.append((current.year, current.month))
    if current.month == 12:
        current = current.replace(year=current.year + 1, month=1)
    else:
        current = current.replace(month=current.month + 1)

row_cursor = 4

for year, month in months:
    # 월 헤더
    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=7)
    month_cell = ws.cell(row=row_cursor, column=1)
    month_cell.value = f"⚓ {year}년 {month}월"
    month_cell.font = Font(name="맑은 고딕", size=13, bold=True, color="FFFFFF")
    month_cell.fill = PatternFill(start_color=DEEP_SEA, end_color=DEEP_SEA, fill_type="solid")
    month_cell.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, 8):
        ws.cell(row=row_cursor, column=c).fill = PatternFill(start_color=DEEP_SEA, end_color=DEEP_SEA, fill_type="solid")
        ws.cell(row=row_cursor, column=c).border = HEADER_BORDER
    ws.row_dimensions[row_cursor].height = 32
    row_cursor += 1

    # 요일 헤더
    for i, dow in enumerate(DOW_KR):
        cell = ws.cell(row=row_cursor, column=i + 1)
        cell.value = dow
        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=OCEAN, end_color=OCEAN, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    ws.row_dimensions[row_cursor].height = 24
    row_cursor += 1

    # 달력 본체 (일요일 시작)
    c = calendar.Calendar(firstweekday=6)
    cal = c.monthdayscalendar(year, month)

    for week_idx, week in enumerate(cal):
        wave = WAVE_COLORS[week_idx % len(WAVE_COLORS)]
        date_row = row_cursor
        content_row = row_cursor + 1

        for col_idx, day in enumerate(week):
            d_cell = ws.cell(row=date_row, column=col_idx + 1)
            c_cell = ws.cell(row=content_row, column=col_idx + 1)

            d_cell.border = DATE_BORDER
            c_cell.border = CONTENT_BORDER
            d_cell.alignment = Alignment(horizontal="left", vertical="center")
            c_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

            if day == 0:
                empty = PatternFill(start_color=EMPTY_BG, end_color=EMPTY_BG, fill_type="solid")
                d_cell.fill = empty
                c_cell.fill = empty
                continue

            current_date = date(year, month, day)
            d_plus = (current_date - BABY_BIRTH_DATE).days
            meal_day = (current_date - MEAL_START_DATE).days + 1
            is_weekend = col_idx == 0 or col_idx == 6  # 일(0), 토(6)

            if is_weekend:
                d_cell.fill = PatternFill(start_color=DATE_ROW_WEEKEND_COLOR, end_color=DATE_ROW_WEEKEND_COLOR, fill_type="solid")
                c_cell.fill = PatternFill(start_color=WEEKEND_BG, end_color=WEEKEND_BG, fill_type="solid")
            else:
                d_cell.fill = PatternFill(start_color=DATE_ROW_COLOR, end_color=DATE_ROW_COLOR, fill_type="solid")
                c_cell.fill = PatternFill(start_color=wave, end_color=wave, fill_type="solid")

            if 1 <= meal_day <= 28:
                # 날짜 행
                d_cell.value = f"{day}일  D+{d_plus}  [{meal_day}일차]"
                d_cell.font = Font(name="맑은 고딕", size=9, bold=True, color=DEEP_SEA)

                # 내용 행
                meal = meal_plan[meal_day]
                lines = []
                if meal["밥"]:
                    lines.append(f"🍚 {meal['밥'].replace(chr(10), ' ')}")
                if meal["고기류"]:
                    lines.append(f"🥩 {meal['고기류'].replace(chr(10), ' ')}")
                if meal["반찬"]:
                    lines.append(f"🥬 {meal['반찬'].replace(chr(10), ' ')}")
                c_cell.value = "\n".join(lines)
                c_cell.font = Font(name="맑은 고딕", size=8, color="2C3E50")
            else:
                d_cell.value = f"{day}일  D+{d_plus}"
                d_cell.font = Font(name="맑은 고딕", size=9, color="85C1E9")
                c_cell.value = ""

        ws.row_dimensions[date_row].height = 22
        ws.row_dimensions[content_row].height = 80
        row_cursor += 2

    # 월 사이 물결
    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=7)
    ws.cell(row=row_cursor, column=1).value = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
    ws.cell(row=row_cursor, column=1).font = Font(name="맑은 고딕", size=7, color=BORDER_COLOR)
    ws.cell(row=row_cursor, column=1).alignment = Alignment(horizontal="center")
    ws.row_dimensions[row_cursor].height = 12
    row_cursor += 1

# === 하단 참고사항 (모래사장 테마) ===
ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=7)
note_header = ws.cell(row=row_cursor, column=1)
note_header.value = f"🐚 {BABY_NAME}의 이유식 항해 안내서"
note_header.font = Font(name="맑은 고딕", size=11, bold=True, color=DEEP_SEA)
note_header.fill = PatternFill(start_color=SAND, end_color=SAND, fill_type="solid")
note_header.alignment = Alignment(horizontal="center", vertical="center")
for c in range(1, 8):
    ws.cell(row=row_cursor, column=c).fill = PatternFill(start_color=SAND, end_color=SAND, fill_type="solid")
ws.row_dimensions[row_cursor].height = 28
row_cursor += 1

notes = [
    "🐟 소고기는 매일, 닭고기는 어쩌다, 생선은 일주일에 2회 이하로 주는 것이 좋습니다.",
    "🥚 계란은 흰자 노른자 같이 줘도 됩니다.",
    "🥜 땅콩은 일주일에 3번 정도 주면 됩니다.",
    "🌿 한 번 첨가한 음식은 다음에 편하게 첨가해도 됩니다.",
    "🥕 2주가 지나면 양배추, 시금치, 당근 같은 상세 식재료를 기본 반찬으로 사용할 수 있습니다.",
    "🥣 미음으로 시작하지 말고 질감 있는 죽으로 시작하세요.",
    "🌾 이유식 초기에 잡곡을 50% 정도 첨가해서 먹여도 됩니다.",
]

for note in notes:
    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=7)
    cell = ws.cell(row=row_cursor, column=1)
    cell.value = note
    cell.font = Font(name="맑은 고딕", size=9, color="2C3E50")
    cell.fill = PatternFill(start_color=SAND_LIGHT, end_color=SAND_LIGHT, fill_type="solid")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, 8):
        ws.cell(row=row_cursor, column=c).fill = PatternFill(start_color=SAND_LIGHT, end_color=SAND_LIGHT, fill_type="solid")
    ws.row_dimensions[row_cursor].height = 20
    row_cursor += 1

wb.save(OUTPUT_FILE)
print(f"Done: {OUTPUT_FILE}")
