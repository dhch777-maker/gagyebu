import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "초기 이유식 식단표"

# ── Colors & Styles ──
TITLE_FILL = PatternFill(start_color="D4483B", end_color="D4483B", fill_type="solid")
TITLE_FONT = Font(name="맑은 고딕", size=16, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="F2D7D5", end_color="F2D7D5", fill_type="solid")
HEADER_FONT = Font(name="맑은 고딕", size=11, bold=True)
WEEK_FILL = PatternFill(start_color="E8B4B8", end_color="E8B4B8", fill_type="solid")
WEEK_FONT = Font(name="맑은 고딕", size=11, bold=True, color="5B2C2C")
EVEN_FILL = PatternFill(start_color="FDF2F2", end_color="FDF2F2", fill_type="solid")
ODD_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
NEW_FILL = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")  # yellow highlight for new foods
DATA_FONT = Font(name="맑은 고딕", size=10)
NOTE_FONT = Font(name="맑은 고딕", size=9, color="888888")
THIN_BORDER = Border(
    left=Side(style="thin", color="D5D5D5"),
    right=Side(style="thin", color="D5D5D5"),
    top=Side(style="thin", color="D5D5D5"),
    bottom=Side(style="thin", color="D5D5D5"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ── Column widths ──
col_widths = {"A": 8, "B": 10, "C": 22, "D": 18, "E": 24, "F": 18, "G": 32}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

# ── Title row ──
ws.merge_cells("A1:G1")
cell = ws["A1"]
cell.value = "초기 이유식 식단표 (만 6개월~)"
cell.font = TITLE_FONT
cell.fill = TITLE_FILL
cell.alignment = CENTER
ws.row_dimensions[1].height = 38

# ── Header row ──
headers = ["주차", "일차", "기본죽", "단백질(고기)", "야채(반찬)", "과일", "비고(새 식재료·TIP)"]
for i, h in enumerate(headers, 1):
    c = ws.cell(row=2, column=i, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = CENTER
    c.border = THIN_BORDER
ws.row_dimensions[2].height = 28

# ── Data ──
# (week, day, porridge, protein, veggie, fruit, note)
data = [
    # ─── 1주차 ───
    (1, 1,
     "쌀미음(10배죽)",
     "-",
     "-",
     "-",
     "★ 첫 이유식! 1~2숟가락부터 시작"),
    (1, 2,
     "쌀미음\n(또는 찹쌀미음)",
     "-",
     "-",
     "-",
     "농도 50%로 묽게"),
    (1, 3,
     "쌀미음",
     "소고기",
     "-",
     "-",
     "★ 소고기 첫 도입 (다진 것)"),
    (1, 4,
     "쌀미음",
     "소고기",
     "-",
     "-",
     "소고기 알레르기 반응 관찰"),
    (1, 5,
     "쌀죽",
     "소고기",
     "감자",
     "-",
     "★ 감자 첫 도입 (근채류)"),
    (1, 6,
     "쌀죽",
     "소고기",
     "감자",
     "-",
     ""),
    (1, 7,
     "쌀죽",
     "소고기",
     "청경채",
     "-",
     "★ 청경채 첫 도입 (엽채류)"),

    # ─── 2주차 ───
    (2, 8,
     "쌀죽",
     "소고기",
     "청경채, 감자",
     "-",
     ""),
    (2, 9,
     "쌀(오트밀)죽",
     "소고기",
     "애호박",
     "-",
     "★ 오트밀·애호박 첫 도입"),
    (2, 10,
     "쌀오트밀죽",
     "소고기",
     "애호박, 감자",
     "-",
     ""),
    (2, 11,
     "쌀오트밀죽",
     "소고기",
     "브로콜리",
     "-",
     "★ 브로콜리 첫 도입"),
    (2, 12,
     "쌀오트밀죽",
     "소고기",
     "양배추",
     "-",
     "★ 양배추 첫 도입"),
    (2, 13,
     "쌀오트밀죽",
     "소고기",
     "고구마",
     "사과",
     "★ 고구마·사과 첫 도입"),
    (2, 14,
     "쌀오트밀죽",
     "소고기",
     "양배추, 당근",
     "사과",
     "★ 당근 첫 도입"),

    # ─── 3주차 ───
    (3, 15,
     "쌀오트밀 죽",
     "소고기",
     "청경채, 단호박\n(또는 시금치+단호박)",
     "사과",
     "★ 단호박 첫 도입"),
    (3, 16,
     "쌀오트밀 죽",
     "소고기, 계란",
     "양배추, 청경채, 당근",
     "사과",
     "★ 계란(노른자) 첫 도입"),
    (3, 17,
     "쌀오트밀 죽",
     "소고기",
     "양배추, 청경채",
     "토마토, 사과",
     "★ 토마토 첫 도입"),
    (3, 18,
     "쌀오트밀 죽",
     "소고기, 계란",
     "양배추, 당근, 단호박",
     "토마토, 사과",
     ""),
    (3, 19,
     "쌀오트밀 죽",
     "소고기",
     "당근, 시금치",
     "토마토, 사과",
     "★ 시금치 첫 도입"),
    (3, 20,
     "쌀오트밀 죽",
     "소고기, 계란",
     "단호박, 청경채",
     "토마토, 사과",
     ""),
    (3, 21,
     "쌀오트밀 죽",
     "소고기, 계란",
     "당근, 시금치",
     "토마토, 사과",
     ""),

    # ─── 4주차 ───
    (4, 22,
     "쌀오트밀 죽\n+ 밀(1주 2회)",
     "소고기",
     "당근, 시금치",
     "토마토, 사과",
     "★ 밀 첫 도입 (1주 2회)"),
    (4, 23,
     "쌀오트밀 죽\n+ 밀가루",
     "소고기, 계란",
     "시금치, 당근",
     "토마토, 사과",
     ""),
    (4, 24,
     "쌀오트밀 죽",
     "소고기",
     "시금치, 당근",
     "토마토, 사과",
     ""),
    (4, 25,
     "쌀오트밀 죽",
     "소고기",
     "단호박, 시금치",
     "사과",
     ""),
    (4, 26,
     "쌀오트밀 죽\n+ 밀가루",
     "소고기, 계란",
     "당근, 시금치",
     "토마토, 사과",
     ""),
    (4, 27,
     "쌀오트밀 죽",
     "소고기",
     "시금치, 당근",
     "토마토, 사과",
     ""),
    (4, 28,
     "쌀오트밀 죽",
     "돼지고기\n(또는 닭고기)",
     "두부, 시금치",
     "사과",
     "★ 돼지고기(또는 닭고기)·두부 첫 도입"),
]

row = 3
current_week = 0
for week, day, porridge, protein, veggie, fruit, note in data:
    # Week divider row
    if week != current_week:
        current_week = week
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        c = ws.cell(row=row, column=1, value=f"── {week}주차 ──")
        c.font = WEEK_FONT
        c.fill = WEEK_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER
        ws.row_dimensions[row].height = 22
        row += 1

    fill = EVEN_FILL if day % 2 == 0 else ODD_FILL
    values = [week, f"{day}일차", porridge, protein, veggie, fruit, note]
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = DATA_FONT
        c.alignment = LEFT_WRAP if i >= 3 else CENTER
        c.fill = fill
        c.border = THIN_BORDER
        # Highlight new food introductions
        if "★" in str(note) and i in (3, 4, 5, 6):
            # check if this column has a new food
            pass

    # Yellow highlight for the note column on new food days
    if "★" in note:
        ws.cell(row=row, column=7).fill = NEW_FILL

    ws.row_dimensions[row].height = 36 if "\n" in porridge or "\n" in protein else 28
    row += 1

# ── Guidelines section ──
row += 1
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws.cell(row=row, column=1, value="📋 초기 이유식 가이드라인")
c.font = Font(name="맑은 고딕", size=12, bold=True, color="D4483B")
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[row].height = 30
row += 1

guidelines = [
    "• 이 식단표는 초기 이유식에서 중기 이유식으로 진행하기 전 단계입니다.",
    "• 1~2숟가락부터 시작하여 아기의 반응을 보며 서서히 양을 늘립니다.",
    "• 새로운 식재료는 2~3일 간격으로 하나씩 추가하며, 알레르기 반응을 관찰합니다.",
    "• 소고기는 철분 보충을 위해 매일 먹이는 것이 좋습니다.",
    "• 이유식 도중 알레르기(발진, 구토, 설사 등) 반응이 있으면 즉시 중단하고 소아과에 방문하세요.",
    "• 밀(밀가루)은 알레르기 유발 가능성이 있으므로 1주 2회 정도로 시작합니다.",
    "• 초기 이유식 완료 후 중기 이유식으로 자연스럽게 진행합니다.",
]
for g in guidelines:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = ws.cell(row=row, column=1, value=g)
    c.font = Font(name="맑은 고딕", size=10, color="555555")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 22
    row += 1

# ── New foods summary ──
row += 1
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws.cell(row=row, column=1, value="🆕 새 식재료 도입 순서")
c.font = Font(name="맑은 고딕", size=12, bold=True, color="D4483B")
ws.row_dimensions[row].height = 30
row += 1

new_foods = [
    ("1일차", "쌀(10배죽)"),
    ("3일차", "소고기"),
    ("5일차", "감자 (근채류)"),
    ("7일차", "청경채 (엽채류)"),
    ("9일차", "오트밀, 애호박"),
    ("11일차", "브로콜리"),
    ("12일차", "양배추"),
    ("13일차", "고구마, 사과"),
    ("14일차", "당근"),
    ("15일차", "단호박"),
    ("16일차", "계란(노른자)"),
    ("17일차", "토마토"),
    ("19일차", "시금치"),
    ("22일차", "밀(밀가루)"),
    ("28일차", "돼지고기/닭고기, 두부"),
]
for day_label, food in new_foods:
    ws.cell(row=row, column=2, value=day_label).font = Font(name="맑은 고딕", size=10, bold=True)
    ws.cell(row=row, column=2).alignment = CENTER
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
    ws.cell(row=row, column=3, value=food).font = Font(name="맑은 고딕", size=10)
    ws.cell(row=row, column=3).alignment = LEFT_WRAP
    ws.row_dimensions[row].height = 20
    row += 1

# ── Freeze panes ──
ws.freeze_panes = "A3"

# ── Print settings ──
ws.print_area = f"A1:G{row}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1

output_path = r"c:\Users\dhchd\work\EUSIK\초기_이유식_식단표.xlsx"
wb.save(output_path)
print(f"Excel saved to: {output_path}")
