# Artwork Extractor V2 — 설계 문서

> CPU 전용 · 모바일/태블릿 최적화 · 웹 게시용 1080×1080

## 1. 개요

폰으로 촬영한 회화 작품 사진(아이가 들고 있는)에서 작품 영역을 자동 검출하고, 원근 보정 크롭 후, 손 부분을 자동 인페인팅하여 깨끗한 작품 이미지를 출력하는 시스템.

### 현재 문제점
- 꼭지점 검출이 복잡한 배경/손 가림에서 실패 빈번
- 손 보정이 수동 브러시만 가능 (자동 감지 없음)
- UI가 데스크탑 중심, 모바일에서 사용 불편

### 목표
- Florence-2 기반 검출 정확도 향상
- MediaPipe Hands로 손 자동 감지 + 자동 인페인팅
- 3단계 위자드 모바일 UI 재설계
- 모델 레이어 분리로 향후 GPU 모델 교체 용이

## 2. 파이프라인 아키텍처

```
[사진 업로드]
    ↓
[1] Florence-2 회화 검출 → 바운딩박스
    ├── 실패 시 → rembg + Canny 폴백
    └── 폴백도 실패 → 수동 4점 지정
    ↓
[2] 바운딩박스 내 엣지 검출 → 정밀 4꼭지점
    ↓
[3] OpenCV getPerspectiveTransform → 원근 보정 크롭
    ↓
[4] MediaPipe Hands → 손 영역 마스크 자동 생성
    ├── 손 감지됨 → LaMa 자동 인페인팅
    └── 손 없음 → 스킵
    ↓
[5] 1080×1080 정사각형 피팅 (흰 배경, 5% 패딩)
    ↓
[결과 출력] → 확인 / 수동보정 / 재시도
```

### 2.1 검출 모듈 (detector.py)

**FloenceDetector** (기본):
- `microsoft/Florence-2-base` 모델 사용 (~0.5GB)
- `<OD>` 태스크로 "painting" 객체 검출 → 바운딩박스 좌표
- 바운딩박스 영역 내에서 Canny 엣지 → 정밀 4꼭지점 추출
- CPU에서 첫 로딩 ~10초, 이후 추론 ~2-3초

**LegacyDetector** (폴백):
- 현재 rembg + Canny 기반 로직 그대로 유지
- Florence-2 검출 실패 시 자동 전환

인터페이스:
```python
class Detector(ABC):
    def detect(self, image: np.ndarray) -> DetectionResult:
        """회화 영역 검출. 4꼭지점 + 신뢰도 반환."""
        pass

@dataclass
class DetectionResult:
    corners: Optional[np.ndarray]  # 4x2 array, 순서: TL, TR, BR, BL
    confidence: float              # 0.0~1.0
    method: str                    # "florence2" | "legacy" | "manual"
```

### 2.2 손 감지 모듈 (hand_detector.py)

**MediaPipeHandDetector**:
- MediaPipe Hands로 손 랜드마크 21개 검출
- 각 손의 랜드마크 → 컨벡스 헐 → 마스크 (팽창 적용)
- 그림 영역(크롭 후)과 겹치는 손 마스크만 반환
- CPU 실시간 동작 가능 (~50ms)

인터페이스:
```python
class HandDetector:
    def detect_hand_mask(self, image: np.ndarray) -> Optional[np.ndarray]:
        """손 영역 마스크 반환. 없으면 None.
        반환: grayscale 마스크 (255=손, 0=배경), 입력과 동일 크기"""
        pass
```

### 2.3 인페인팅 모듈 (inpainter.py)

**LamaInpainter** (현재):
- simple-lama-inpainting 사용
- CPU에서 ~1-2초 (1080px 기준)
- 검증된 품질, 그대로 유지

인터페이스:
```python
class Inpainter(ABC):
    def inpaint(self, image: PIL.Image, mask: PIL.Image) -> PIL.Image:
        """마스크 영역 인페인팅. RGB 이미지 반환."""
        pass
```

나중에 `PixelHackerInpainter` 등 추가 가능.

### 2.4 오케스트레이터 (extractor.py)

기존 `process_photo()` 리팩토링:

```python
def process_photo(img, output_size=1080) -> ProcessResult:
    # 1. 검출
    result = florence_detector.detect(img)
    if result.corners is None:
        result = legacy_detector.detect(img)

    if result.corners is None:
        return ProcessResult(success=False, needs_manual=True)

    # 2. 원근 보정
    cropped = perspective_transform(img, result.corners)

    # 3. 손 감지 + 자동 인페인팅
    hand_mask = hand_detector.detect_hand_mask(cropped)
    if hand_mask is not None:
        cropped = inpainter.inpaint(cropped, hand_mask)

    # 4. 정사각형 피팅
    final = fit_to_square(cropped, output_size)

    return ProcessResult(
        success=True,
        image=final,
        hand_detected=hand_mask is not None,
        detection_method=result.method,
        confidence=result.confidence
    )
```

## 3. 모바일 UI 설계

### 3.1 3단계 위자드

**STEP 1 — 촬영**
- 큰 촬영 버튼 (카메라 직접 or 갤러리 선택)
- `<input type="file" accept="image/*" capture="environment">` → 후면 카메라 기본
- 선택 즉시 업로드 + 자동 처리 시작
- 처리 중: 프로그레스 바 + 단계별 텍스트 ("회화 검출 중...", "손 보정 중...")

**STEP 2 — 확인/수정**
- 결과 이미지 크게 표시 (화면 폭 100%)
- 원본 ↔ 결과 스와이프 비교 (터치 좌우 슬라이드)
- 손 보정이 된 경우 "손 자동 보정됨" 배지 표시
- 검출 신뢰도 표시 (높음/중간/낮음)
- 버튼: "이대로 저장" (주 액션), "꼭지점 수정" / "추가 보정" (보조)

**STEP 3 — 수동 보정** (선택적)
- 탭 전환: 꼭지점 모드 / 브러시 모드
- 꼭지점 모드: 4개 핸들 터치 드래그 (44×44px 최소 터치 타겟)
- 브러시 모드: 기존 인페인팅 UI (터치 페인팅 + 사이즈 슬라이더)
- 되돌리기 버튼
- "적용" → STEP 2로 복귀

### 3.2 모바일 최적화

- 최소 터치 타겟: 44×44px
- 이미지 업로드 전 클라이언트에서 리사이즈 (max 2048px → 전송 속도)
- 세로 모드 기본, 이미지 영역 최대화
- 버튼 하단 고정 (position: sticky)
- 핀치 줌 지원 (STEP 3 캔버스)

### 3.3 API 변경

기존 API 유지 + 응답 필드 추가:

```
POST /upload 응답 추가 필드:
{
    "hand_detected": true,        // 손 자동 보정 여부
    "detection_method": "florence2", // 검출 방법
    "confidence": 0.92            // 검출 신뢰도
}
```

기존 `/manual-crop`, `/inpaint` 엔드포인트는 그대로 유지.

## 4. 파일 구조

```
gallery/
├── app.py                  # Flask 라우트 (변경 최소화)
├── extractor.py            # 오케스트레이터 (리팩토링)
├── detector.py             # NEW: 검출 모듈
│   ├── Detector (ABC)
│   ├── FlorenceDetector
│   └── LegacyDetector
├── hand_detector.py        # NEW: 손 감지 모듈
│   └── MediaPipeHandDetector
├── inpainter.py            # NEW: 인페인팅 모듈
│   └── LamaInpainter
├── templates/
│   └── index.html          # 재설계 (위자드 UI)
├── static/
│   ├── app.js              # 재작성 (위자드 로직)
│   └── style.css           # 재작성 (모바일 퍼스트)
├── requirements.txt        # 의존성 추가
└── tests/
    ├── test_extractor.py   # 업데이트
    ├── test_detector.py    # NEW
    └── test_hand_detector.py # NEW
```

## 5. 의존성 추가

```
transformers              # Florence-2
torch                     # Florence-2 백엔드 (CPU)
mediapipe                 # 손 감지
```

기존 의존성 유지: flask, opencv-python-headless, numpy, rembg[cpu], Pillow, simple-lama-inpainting

## 6. 성능 예상 (CPU)

| 단계 | 예상 시간 |
|------|-----------|
| Florence-2 초기 로딩 | ~10초 (앱 시작 시 1회) |
| Florence-2 추론 | ~2-3초 |
| 엣지 검출 + 꼭지점 | <0.5초 |
| 원근 보정 | <0.5초 |
| MediaPipe 손 감지 | <0.1초 |
| LaMa 인페인팅 | ~1-2초 |
| 정사각형 피팅 | <0.1초 |
| **전체 파이프라인** | **~4-6초** |

## 7. 범위 외 (YAGNI)

- 배치 처리 (여러 장 동시)
- 사용자 계정/인증
- 클라우드 GPU 연동
- 고해상도 출력 (2000px+)
- 동영상 처리
