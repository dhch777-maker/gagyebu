# ZETO 마케팅 시안 #03 — Paint Splatter ("MARK THE START") 디자인 스펙

- **날짜:** 2026-05-22
- **브랜드:** ZETO ART
- **목표:** 인스타그램 홍보물 시리즈에 1장 추가. 기존 2장(찢긴 종이 / 구겨진 종이)과 시각 언어가 확실히 다른 "잉크 스플래시" 변주 도입
- **결과물 디렉터리:** `ZETO_marketing/`

## As-built 노트 (2026-05-22)

실제 FLUX 출력에서 스플래시가 y=180~650 영역을 차지해 원안의 워드마크(y=180~276)·디바이더(y=300) 자리와 충돌함. 블루 워드마크가 블루 스플래시 위에 깔려 가독성 저하. 시드 변경 대신 **레이아웃을 양보**하는 쪽으로 결정:

- 슬롯 03의 `texts`에서 `wordmark` 키 제거 → 상단 워드마크 미표시
- `SPECS["03"]["draw_divider"] = False` → 디바이더 라인 미표시
- 브랜드 인지는 한글 카피 안의 "ZETO ART" + 하단 핸들 `@ZETO_magok`로 유지
- 스플래시가 시각적 히어로가 되고 "MARK THE START."가 스플래시 하단에 걸쳐 에디토리얼 톤 강화

코드 변경: `compositor.compose()`가 `draw_divider` 플래그와 누락된 텍스트 키를 허용하도록 일반화 (commit `c61b5f52`). 후속 슬롯(04+)은 텍스트-배경 충돌이 없으면 워드마크·디바이더를 다시 켤 수 있음 — 스펙 §3 의 원안 레이아웃은 그대로 유효한 템플릿.

## 1. 컨텍스트

`ZETO_marketing/` 안에 인스타용 홍보물 2장이 있다.

| # | 파일 | 톤 |
|---|---|---|
| 01 | `KakaoTalk_20260521_161056745.png` | 회색 종이 찢김 + 파란 내부에 ZETO 타이포 / "분트→ZETO" 리브랜드 |
| 02 | `KakaoTalk_20260521_165113695.png` | 파란 구겨진 종이 + 큰 "NEW IS COMING" / 3개 분원 통합 |

두 장은 모두 "종이 + 타이포" 계열이라 시리즈가 정적이다. 사용자 지시는 **"확실히 다른 느낌"**으로 1장 추가. 단, 같은 브랜드(ZETO ART, @ZETO_magok)·같은 인스타 정사각 포맷·같은 한글 카피 톤은 유지.

작업 트랙은 별도 폴더 `ZETO_insta/`(다크 브루탈리즘 인스타 론칭 그리드 9장)와는 무관한 독립 마케팅 시리즈로 본다.

## 2. 비주얼 컨셉

**컨셉명**: "MARK THE START"

**아이디어**: 아이보리 종이 위에 브랜드 블루 잉크가 한 방울 떨어진 순간. 잉크 자국이 곧 시작점이고, ZETO ART의 시작을 상징한다. 페인트 스플래시 형상이 기존 #1(찢김)과 #2(구겨짐)에 없던 "동적·에너지" 결을 도입한다.

**왜 페인트 스플래시인가:**
- ZETO ART 브랜드 정체성과 직결 (미술 = 페인트)
- 기존 두 장의 정적 톤과 시각 언어가 명확히 분리됨
- FLUX.1-dev가 안정적으로 잘 생성하는 영역 (잉크·페인트 텍스처)
- 이후 시안에서 색·형상만 바꿔 확장하기 쉬움 (시리즈 일관성 유지에 유리)

## 3. 디자인 사양

### 캔버스
- 1080 × 1080 (인스타 정사각, 기존 2장과 동일)
- 최종 포맷: PNG

### 색상 토큰

| 항목 | 값 | 출처 |
|---|---|---|
| 배경 | `#f4f1ea` (warm ivory) | ZETO_insta 시스템과 동일 |
| 스플래시 (메인) | `#266a9c` (ZETO 블루) | 기존 #2와 동일 |
| 텍스트 — 헤더 | `#266a9c` (블루) | 워드마크 강조 |
| 텍스트 — 본문 | `#1a1a1a` (다크 그레이) | 가독성 |
| 텍스트 — 메타 | `#9a978f` (뮤트 그레이) | 상하단 보조 |

### 타이포그래피

| 위치 | 폰트 | 사이즈(px @ 1080²) | 색 |
|---|---|---|---|
| 상단 메타 ("ZERO TO ART · Vol. 03") | ui-monospace / Courier | 18 | `#9a978f` |
| 메인 워드마크 ("ZETO ART") | Druk Heavy | 96 | `#266a9c` |
| 메인 카피 ("MARK THE START.") | Druk Heavy | 72 | `#1a1a1a` |
| 한글 서브카피 (2줄) | Noto Sans KR Bold | 26 | `#1a1a1a` |
| 하단 핸들 ("@ZETO_magok") | Noto Sans KR Regular | 20 | `#9a978f` |

폰트 파일: `ZETO_insta/fonts/Druk-Heavy-Trial.otf`를 `ZETO_marketing/fonts/`로 복사해 사용. Noto Sans KR은 시스템 설치본 또는 Google Fonts 로컬 다운로드.

### 카피 (확정안)

```
상단 메타:   ZERO TO ART · Vol. 03
워드마크:    ZETO ART
메인 카피:   MARK THE START.
한글 서브:   미술이 시작되는 순간을
            ZETO ART에서 그려갑니다.
하단 핸들:   @ZETO_magok
```

### 레이아웃 (좌표 — 1080² 기준, y는 텍스트 박스 상단 기준)

```
y=  60  상단 메타 (가로 중앙)                          ─ 메타 행 y=60~78
y= 180  ZETO ART 워드마크 (가로 중앙)                  ─ 워드마크 y=180~276
y= 300  ──── 가로 라인 (480px, 가로 중앙, 2px, #266a9c)
y= 360  잉크 스플래시 영역 시작                        ─ 스플래시 y=360~640
y= 720  MARK THE START. (가로 중앙)                    ─ 헤드라인 y=720~792
y= 850  한글 서브카피 1행                              ─ 한글1 y=850~876
y= 890  한글 서브카피 2행                              ─ 한글2 y=890~916
y=1020  @ZETO_magok (가로 중앙)                        ─ 핸들 y=1020~1040
```

스플래시는 텍스트와 겹치지 않도록 y=360~640 영역 안에 자리잡도록 FLUX 프롬프트에 명시.

## 4. 기술 파이프라인

```
[1] FLUX.1-dev 배경 생성        [2] PIL 텍스트 합성       [3] 결과
    HF Inference API              로컬 Python
    ────────────────             ───────────────
    프롬프트 → PNG               배경 + 폰트 + 색 + 좌표   1080×1080 PNG
    (블루 잉크 스플래시,         → composite               ZETO_marketing/
     텍스트 없음)                                          03-mark-the-start.png
```

### [1] 배경 생성 — FLUX.1-dev via HF Inference API

**모델**: `black-forest-labs/FLUX.1-dev`
**호출**: `huggingface_hub.InferenceClient.text_to_image()`

**프롬프트**:
```
minimalist art poster background, single deep cobalt blue ink splatter
on warm ivory paper, organic dynamic splash shape positioned in
upper-center area, small ink droplets scattered around the main splash,
generous negative space top and bottom, high contrast, fine paper grain
texture, no text no letters, 1:1 square format, editorial poster
aesthetic, clean composition
```

**negative_prompt**:
```
text, letters, words, watermark, signature, multiple colors, busy
composition, photograph, realistic person, frame, border
```

**파라미터**:
- `width`: 1024 (FLUX 권장값 → 1080으로 PIL 업스케일)
- `height`: 1024
- `guidance_scale`: 3.5
- `num_inference_steps`: 28
- `seed`: 고정값 (재현성용, 코드 상수)

### [2] 텍스트 합성 — Python + Pillow

`generate.py`가 FLUX 출력 PNG를 받아:
1. 1024 → 1080 리사이즈
2. `ImageDraw.text()`로 위 좌표/색/폰트대로 5개 텍스트 블록 그리기
3. 가운데 정렬은 `getbbox()` 기반 픽셀 정확 계산
4. 한글 줄바꿈은 수동 분할 (자동 wrap 미사용 — 디자인 의도 보존)

### [3] 출력
- `ZETO_marketing/03-mark-the-start.png` — 최종 결과물
- `ZETO_marketing/_bg/03-bg.png` — FLUX 원본 (디버깅·재합성용, gitignore 대상 아님 — 시안 검증용)

## 5. 파일 구조

```
ZETO_marketing/
├── KakaoTalk_20260521_161056745.png    (기존 #1, 유지)
├── KakaoTalk_20260521_165113695.png    (기존 #2, 유지)
├── 03-mark-the-start.png                (이번 결과물)
├── _bg/
│   └── 03-bg.png                        (FLUX 원본)
├── fonts/
│   └── Druk-Heavy-Trial.otf             (ZETO_insta에서 복사)
├── generate.py                          (배경 생성 + 텍스트 합성)
├── prompts.py                           (시안별 설정 dict)
├── requirements.txt                     (huggingface_hub, Pillow, python-dotenv)
├── .env                                 (HF_TOKEN — .gitignore)
└── .gitignore                           (.env, __pycache__/)
```

### 핵심 설계 — `prompts.py` 모듈화

후속 04, 05, 06… 시안 추가 시 코드 변경 없이 dict entry만 추가:

```python
# prompts.py
SPECS = {
    "03": {
        "slug": "mark-the-start",
        "flux_prompt": "...",
        "flux_negative": "...",
        "seed": 20260522,
        "texts": {
            "meta_top": "ZERO TO ART · Vol. 03",
            "wordmark": "ZETO ART",
            "headline": "MARK THE START.",
            "kor_line1": "미술이 시작되는 순간을",
            "kor_line2": "ZETO ART에서 그려갑니다.",
            "handle": "@ZETO_magok",
        },
    },
    # "04": { ... } ← 다음 시안 추가 위치
}
```

`generate.py` 실행: `python generate.py 03`

## 6. 의존성

```
# requirements.txt
huggingface_hub>=0.24.0
Pillow>=10.0.0
python-dotenv>=1.0.0
```

Python 3.10+. 가상환경(`venv`) 권장이지만 강제 아님.

## 7. 검증 방법

작업 완료 판정 기준:

1. **시각적 검증** — 결과 PNG를 직접 열어 확인:
   - 스플래시가 텍스트 영역(상단 메타/워드마크/하단 카피)을 침범하지 않음
   - 카피 5개 블록 모두 가독성 확보 (특히 한글 2행)
   - 기존 2장 옆에 두고 톤 일관성(브랜드 블루 채도·아이보리 배경 톤) 확인
2. **모바일 미리보기** — 320×320으로 다운스케일 후 작은 화면에서도 워드마크·한글 카피가 읽히는지 확인
3. **재현성** — `seed` 고정 상태에서 동일 결과 나오는지 한 번 더 실행

만약 스플래시 형상이 의도와 다르면(너무 작거나·텍스트 영역 침범·색이 어긋남) `seed` 값만 바꿔 재실행. 프롬프트는 마지막에 조정.

## 8. 스코프 밖 (이번 작업에서 하지 않음)

- 후속 시안 04, 05… 실제 생성 (이번엔 03 한 장만)
- `ZETO_insta/` 폴더 변경 (별도 트랙)
- 동영상·스토리·릴스 포맷 (정사각 피드 포스트만)
- 자동 카피 생성·다국어 변환
- HF Inference API 외 다른 백엔드(Replicate, fal 등) 지원

## 9. 사전 조건

- HF Inference API 토큰 (Read 권한) — 사용자가 huggingface.co에서 발급해 `.env`에 저장
- 인터넷 연결 (FLUX 호출 시)
- Python 3.10+ 설치
