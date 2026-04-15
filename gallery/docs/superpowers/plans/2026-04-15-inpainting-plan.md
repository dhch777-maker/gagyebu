# AI 부분 보정 (인페인팅) 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 추출된 그림에서 손가락 등 불필요한 부분을 드래그로 칠하면 LaMa AI가 주변 텍스처로 자연스럽게 채워주는 인페인팅 기능 추가

**Architecture:** Flask 백엔드에 `/inpaint` 엔드포인트 추가, `simple-lama-inpainting` 패키지로 마스크 영역 보정. 프론트엔드는 Canvas 기반 드래그 페인팅으로 마스크 생성 후 API 호출, Undo 스택으로 되돌리기 지원.

**Tech Stack:** simple-lama-inpainting (LaMa), Flask, Canvas API, Pillow

---

## 파일 구조

| 파일 | 작업 | 역할 |
|------|------|------|
| `requirements.txt` | 수정 | `simple-lama-inpainting` 추가 |
| `extractor.py` | 수정 | `inpaint_region()` 함수 추가 |
| `app.py` | 수정 | `/inpaint` 엔드포인트 추가 |
| `templates/index.html` | 수정 | 인페인팅 모드 UI 섹션 추가 |
| `static/app.js` | 수정 | 드래그 페인팅 + API 호출 로직 추가 |
| `static/style.css` | 수정 | 인페인팅 UI 스타일 추가 |
| `tests/test_extractor.py` | 수정 | 인페인팅 함수 테스트 추가 |

---

### Task 1: 의존성 추가 및 LaMa 설치 확인

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: requirements.txt에 simple-lama-inpainting 추가**

```
flask
opencv-python-headless
numpy
rembg[cpu]
Pillow
simple-lama-inpainting
```

- [ ] **Step 2: 패키지 설치**

Run: `pip install simple-lama-inpainting`
Expected: Successfully installed simple-lama-inpainting

- [ ] **Step 3: 설치 확인**

Run: `python -c "from simple_lama_inpainting import SimpleLama; print('OK')"`
Expected: `OK`

- [ ] **Step 4: Commit**

```bash
git add requirements.txt
git commit -m "feat: simple-lama-inpainting 의존성 추가"
```

---

### Task 2: 백엔드 인페인팅 함수 (테스트 먼저)

**Files:**
- Modify: `extractor.py:1-8` (import 및 모델 로드 추가)
- Modify: `extractor.py` (함수 추가)
- Modify: `tests/test_extractor.py`

- [ ] **Step 1: 인페인팅 테스트 작성**

`tests/test_extractor.py` 끝에 추가:

```python
def test_inpaint_region_fills_masked_area():
    from extractor import inpaint_region
    from PIL import Image
    import numpy as np

    # Create a 200x200 red image
    img = Image.new("RGB", (200, 200), (200, 100, 100))

    # Create mask: white circle in center (area to inpaint)
    mask = Image.new("L", (200, 200), 0)
    from PIL import ImageDraw
    draw = ImageDraw.Draw(mask)
    draw.ellipse([80, 80, 120, 120], fill=255)

    result = inpaint_region(img, mask)

    assert isinstance(result, Image.Image), "Should return PIL Image"
    assert result.size == (200, 200), "Should preserve original size, got {}".format(result.size)
    assert result.mode == "RGB", "Should return RGB image"
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `cd gallery && python -m pytest tests/test_extractor.py::test_inpaint_region_fills_masked_area -v`
Expected: FAIL with "cannot import name 'inpaint_region'"

- [ ] **Step 3: extractor.py에 LaMa 모델 로드 및 inpaint_region 함수 구현**

`extractor.py` 상단 import 영역에 추가:

```python
from simple_lama_inpainting import SimpleLama

# Pre-load LaMa inpainting model (loaded once at startup)
_lama_model = SimpleLama()
```

`extractor.py` 하단에 함수 추가:

```python
def inpaint_region(img: Image.Image, mask: Image.Image) -> Image.Image:
    """Inpaint masked region using LaMa model.

    Args:
        img: RGB PIL Image.
        mask: Grayscale PIL Image (white=area to inpaint).

    Returns:
        Inpainted RGB PIL Image.
    """
    result = _lama_model(img, mask)
    return result.convert("RGB")
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `cd gallery && python -m pytest tests/test_extractor.py::test_inpaint_region_fills_masked_area -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: LaMa 기반 inpaint_region 함수 추가 (테스트 포함)"
```

---

### Task 3: 백엔드 `/inpaint` API 엔드포인트

**Files:**
- Modify: `app.py:4` (import 추가)
- Modify: `app.py` (라우트 추가)

- [ ] **Step 1: app.py import에 inpaint_region 추가**

`app.py:4`를 수정:

```python
from extractor import process_photo, manual_process, inpaint_region
```

추가 import도 필요:

```python
from PIL import Image
import base64
from io import BytesIO
```

- [ ] **Step 2: `/inpaint` 엔드포인트 추가**

`app.py`의 `serve_processed` 함수 위에 추가:

```python
@app.route("/inpaint", methods=["POST"])
def inpaint():
    data = request.get_json()
    image_b64 = data.get("image")
    mask_b64 = data.get("mask")
    file_id = data.get("file_id")

    if not image_b64 or not mask_b64 or not file_id:
        return jsonify({"error": "image, mask, file_id가 필요합니다"}), 400

    # Decode base64 images
    img_data = base64.b64decode(image_b64.split(",")[1] if "," in image_b64 else image_b64)
    mask_data = base64.b64decode(mask_b64.split(",")[1] if "," in mask_b64 else mask_b64)

    img = Image.open(BytesIO(img_data)).convert("RGB")
    mask = Image.open(BytesIO(mask_data)).convert("L")

    # Run inpainting
    result = inpaint_region(img, mask)

    # Save to processed directory
    processed_name = f"{file_id}_cropped.jpg"
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    result.save(processed_path, "JPEG", quality=92)

    # Return as base64
    buf = BytesIO()
    result.save(buf, format="JPEG", quality=92)
    result_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

    return jsonify({
        "result_image": "data:image/jpeg;base64," + result_b64,
        "processed": f"/files/processed/{processed_name}",
        "message": "부분 보정 완료!",
    })
```

- [ ] **Step 3: 수동 테스트로 엔드포인트 확인**

Run: `cd gallery && python -c "
from app import app
import json, base64
from PIL import Image
from io import BytesIO

client = app.test_client()

# Create test image
img = Image.new('RGB', (100, 100), (200, 100, 100))
buf = BytesIO()
img.save(buf, 'JPEG')
img_b64 = base64.b64encode(buf.getvalue()).decode()

# Create test mask
mask = Image.new('L', (100, 100), 0)
buf2 = BytesIO()
mask.save(buf2, 'PNG')
mask_b64 = base64.b64encode(buf2.getvalue()).decode()

resp = client.post('/inpaint', json={'image': img_b64, 'mask': mask_b64, 'file_id': 'test123'})
data = json.loads(resp.data)
print('Status:', resp.status_code)
print('Has result:', 'result_image' in data)
print('Message:', data.get('message'))
"`
Expected: Status: 200, Has result: True, Message: 부분 보정 완료!

- [ ] **Step 4: Commit**

```bash
git add app.py
git commit -m "feat: /inpaint API 엔드포인트 추가 (base64 이미지+마스크 → LaMa 보정)"
```

---

### Task 4: 프론트엔드 — 인페인팅 모드 HTML 구조

**Files:**
- Modify: `templates/index.html`

- [ ] **Step 1: 미리보기 actions에 "부분 보정" 버튼 추가**

`templates/index.html`의 `.actions` div (line 35-39)를 수정:

```html
<div class="actions">
    <button id="confirmBtn" class="btn-confirm">확인 — 이대로 저장</button>
    <button id="inpaintBtn" class="btn-inpaint">부분 보정</button>
    <button id="manualBtn" class="btn-manual">수동 지정</button>
    <button id="retryBtn" class="btn-retry">다시하기</button>
</div>
```

- [ ] **Step 2: 인페인팅 섹션 HTML 추가**

`templates/index.html`의 `manualSection` div 뒤 (line 64 이후, `</div>` 닫기 전)에 추가:

```html
<div id="inpaintSection" class="inpaint-section" hidden>
    <p class="inpaint-hint">제거할 부분을 드래그로 칠해주세요 (마우스를 떼면 자동 보정됩니다)</p>
    <div class="inpaint-controls">
        <label for="brushSize">브러시 크기:</label>
        <input type="range" id="brushSize" min="5" max="100" value="30">
        <span id="brushSizeLabel">30px</span>
    </div>
    <div class="inpaint-canvas-area">
        <canvas id="inpaintCanvas"></canvas>
    </div>
    <div class="actions">
        <button id="inpaintUndoBtn" class="btn-retry" disabled>되돌리기 (Undo)</button>
        <button id="inpaintDoneBtn" class="btn-confirm">완료</button>
        <button id="inpaintCancelBtn" class="btn-retry">취소</button>
    </div>
</div>
```

- [ ] **Step 3: Commit**

```bash
git add templates/index.html
git commit -m "feat: 인페인팅 모드 HTML 구조 추가 (부분 보정 버튼, 캔버스, 브러시 컨트롤)"
```

---

### Task 5: 프론트엔드 — 인페인팅 CSS 스타일

**Files:**
- Modify: `static/style.css`

- [ ] **Step 1: 인페인팅 관련 스타일 추가**

`static/style.css` 끝에 추가:

```css
/* Inpainting */
.btn-inpaint {
    padding: 0.7rem 1.5rem;
    font-size: 1rem;
    background: #9b59b6;
    color: white;
    border: none;
    border-radius: 8px;
    cursor: pointer;
}
.inpaint-section { margin-top: 2rem; }
.inpaint-hint {
    margin-bottom: 1rem;
    color: #555;
    font-size: 1.05rem;
    font-weight: 500;
}
.inpaint-controls {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.8rem;
    margin-bottom: 1rem;
}
.inpaint-controls label { font-size: 0.95rem; color: #555; }
.inpaint-controls input[type="range"] { width: 200px; }
.inpaint-controls span { font-size: 0.9rem; color: #888; min-width: 45px; }
.inpaint-canvas-area {
    margin: 0 auto;
    max-width: 700px;
    display: inline-block;
    position: relative;
}
.inpaint-canvas-area canvas {
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.12);
    max-width: 100%;
}
.inpaint-spinner {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    background: rgba(0,0,0,0.5);
    color: white;
    padding: 1rem 2rem;
    border-radius: 8px;
    font-size: 1rem;
    z-index: 10;
}
```

- [ ] **Step 2: Commit**

```bash
git add static/style.css
git commit -m "feat: 인페인팅 모드 CSS 스타일 추가"
```

---

### Task 6: 프론트엔드 — 인페인팅 JavaScript 로직

**Files:**
- Modify: `static/app.js`

- [ ] **Step 1: 인페인팅 DOM 요소 참조 추가**

`static/app.js`의 기존 DOM 참조 영역 (line 13 `failRetryBtn` 뒤)에 추가:

```javascript
// Inpainting elements
var inpaintBtn = document.getElementById("inpaintBtn");
var inpaintSection = document.getElementById("inpaintSection");
var inpaintCanvas = document.getElementById("inpaintCanvas");
var inpaintCtx = inpaintCanvas.getContext("2d");
var brushSize = document.getElementById("brushSize");
var brushSizeLabel = document.getElementById("brushSizeLabel");
var inpaintUndoBtn = document.getElementById("inpaintUndoBtn");
var inpaintDoneBtn = document.getElementById("inpaintDoneBtn");
var inpaintCancelBtn = document.getElementById("inpaintCancelBtn");
```

- [ ] **Step 2: 인페인팅 상태 변수 추가**

State 섹션 (line 36-39 부근, `var manualScale = 1;` 뒤)에 추가:

```javascript
// Inpainting state
var inpaintImg = null;
var inpaintScale = 1;
var inpaintDrawing = false;
var inpaintMaskCanvas = null;  // offscreen canvas for mask
var inpaintMaskCtx = null;
var inpaintHistory = [];  // undo stack of {imageDataUrl, maskDataUrl}
```

- [ ] **Step 3: resetUI에 인페인팅 초기화 추가**

`resetUI` 함수의 `manualPoints = [];` 뒤에 추가:

```javascript
inpaintSection.hidden = true;
inpaintHistory = [];
inpaintUndoBtn.disabled = true;
```

그리고 `inpaintBtn.hidden = false;`를 `manualBtn.hidden = false;` 뒤에 추가.

`confirmBtn` 이벤트의 `manualBtn.hidden = true;` 뒤에 `inpaintBtn.hidden = true;` 추가.

- [ ] **Step 4: 인페인팅 모드 진입 함수 추가**

`app.js` 하단, `confirmBtn.addEventListener` 앞에 추가:

```javascript
// --- Inpainting Mode ---

function startInpaint() {
    preview.hidden = true;
    failMessage.hidden = true;
    manualSection.hidden = true;
    inpaintSection.hidden = false;
    inpaintHistory = [];
    inpaintUndoBtn.disabled = true;

    loadInpaintImage();
}

function loadInpaintImage() {
    var url = currentProcessedUrl || currentOriginalUrl;
    inpaintImg = new Image();
    inpaintImg.onload = function() {
        var maxW = Math.min(700, window.innerWidth - 40);
        inpaintScale = Math.min(maxW / inpaintImg.width, 1);
        inpaintCanvas.width = Math.round(inpaintImg.width * inpaintScale);
        inpaintCanvas.height = Math.round(inpaintImg.height * inpaintScale);

        // Create offscreen mask canvas (same size as display canvas)
        inpaintMaskCanvas = document.createElement("canvas");
        inpaintMaskCanvas.width = inpaintCanvas.width;
        inpaintMaskCanvas.height = inpaintCanvas.height;
        inpaintMaskCtx = inpaintMaskCanvas.getContext("2d");
        inpaintMaskCtx.fillStyle = "black";
        inpaintMaskCtx.fillRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);

        drawInpaintCanvas();
    };
    inpaintImg.src = url + "?t=" + Date.now();
}

function drawInpaintCanvas() {
    inpaintCtx.drawImage(inpaintImg, 0, 0, inpaintCanvas.width, inpaintCanvas.height);

    // Overlay mask with semi-transparent red
    inpaintCtx.save();
    inpaintCtx.globalAlpha = 0.4;
    inpaintCtx.drawImage(inpaintMaskCanvas, 0, 0);

    // Draw only red where mask is white: use composite
    inpaintCtx.restore();

    // Better approach: draw red overlay where mask is white
    var maskData = inpaintMaskCtx.getImageData(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
    var overlay = inpaintCtx.getImageData(0, 0, inpaintCanvas.width, inpaintCanvas.height);
    for (var i = 0; i < maskData.data.length; i += 4) {
        if (maskData.data[i] > 128) {  // white pixel in mask = area to inpaint
            overlay.data[i] = Math.min(255, overlay.data[i] + 100);     // R
            overlay.data[i + 1] = Math.max(0, overlay.data[i + 1] - 50); // G
            overlay.data[i + 2] = Math.max(0, overlay.data[i + 2] - 50); // B
            overlay.data[i + 3] = 200;  // semi-transparent
        }
    }
    inpaintCtx.putImageData(overlay, 0, 0);
}

function drawBrushCursor(x, y) {
    var r = parseInt(brushSize.value) * inpaintScale;
    inpaintCtx.beginPath();
    inpaintCtx.arc(x, y, r, 0, Math.PI * 2);
    inpaintCtx.strokeStyle = "rgba(255, 255, 255, 0.8)";
    inpaintCtx.lineWidth = 2;
    inpaintCtx.stroke();
}

function paintMask(x, y) {
    var r = parseInt(brushSize.value) * inpaintScale;
    inpaintMaskCtx.beginPath();
    inpaintMaskCtx.arc(x, y, r, 0, Math.PI * 2);
    inpaintMaskCtx.fillStyle = "white";
    inpaintMaskCtx.fill();
}

function sendInpaintRequest() {
    // Show spinner overlay
    var spinnerEl = document.createElement("div");
    spinnerEl.className = "inpaint-spinner";
    spinnerEl.textContent = "AI 보정 중...";
    inpaintCanvas.parentElement.appendChild(spinnerEl);

    // Save current state to undo history
    inpaintHistory.push(inpaintImg.src);
    inpaintUndoBtn.disabled = false;

    // Get current image as base64 (original resolution)
    var imgCanvas = document.createElement("canvas");
    imgCanvas.width = inpaintImg.naturalWidth;
    imgCanvas.height = inpaintImg.naturalHeight;
    var imgCtx = imgCanvas.getContext("2d");
    imgCtx.drawImage(inpaintImg, 0, 0);
    var imageB64 = imgCanvas.toDataURL("image/jpeg", 0.92);

    // Scale mask to original resolution
    var maskFullCanvas = document.createElement("canvas");
    maskFullCanvas.width = inpaintImg.naturalWidth;
    maskFullCanvas.height = inpaintImg.naturalHeight;
    var maskFullCtx = maskFullCanvas.getContext("2d");
    maskFullCtx.drawImage(inpaintMaskCanvas, 0, 0, maskFullCanvas.width, maskFullCanvas.height);
    var maskB64 = maskFullCanvas.toDataURL("image/png");

    fetch("/inpaint", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            image: imageB64,
            mask: maskB64,
            file_id: currentFileId
        })
    })
    .then(function(resp) {
        return resp.json().then(function(data) {
            return { ok: resp.ok, data: data };
        });
    })
    .then(function(result) {
        spinnerEl.remove();

        if (!result.ok) {
            alert(result.data.error || "보정 실패");
            return;
        }

        // Update image with result
        currentProcessedUrl = result.data.processed;
        inpaintImg = new Image();
        inpaintImg.onload = function() {
            // Clear mask
            inpaintMaskCtx.fillStyle = "black";
            inpaintMaskCtx.fillRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
            drawInpaintCanvas();
        };
        inpaintImg.src = result.data.result_image;
    })
    .catch(function() {
        spinnerEl.remove();
        alert("서버 연결에 실패했습니다.");
    });
}
```

- [ ] **Step 5: 캔버스 이벤트 리스너 추가**

위 함수들 바로 뒤에 추가:

```javascript
// Inpainting canvas events
inpaintCanvas.addEventListener("mousedown", function(e) {
    inpaintDrawing = true;
    var rect = inpaintCanvas.getBoundingClientRect();
    var x = e.clientX - rect.left;
    var y = e.clientY - rect.top;
    paintMask(x, y);
    drawInpaintCanvas();
    drawBrushCursor(x, y);
});

inpaintCanvas.addEventListener("mousemove", function(e) {
    var rect = inpaintCanvas.getBoundingClientRect();
    var x = e.clientX - rect.left;
    var y = e.clientY - rect.top;

    if (inpaintDrawing) {
        paintMask(x, y);
    }
    drawInpaintCanvas();
    drawBrushCursor(x, y);
});

inpaintCanvas.addEventListener("mouseup", function() {
    if (!inpaintDrawing) return;
    inpaintDrawing = false;

    // Check if any mask was painted
    var maskData = inpaintMaskCtx.getImageData(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
    var hasWhite = false;
    for (var i = 0; i < maskData.data.length; i += 4) {
        if (maskData.data[i] > 128) { hasWhite = true; break; }
    }
    if (hasWhite) {
        sendInpaintRequest();
    }
});

inpaintCanvas.addEventListener("mouseleave", function() {
    if (inpaintDrawing) {
        inpaintDrawing = false;
        sendInpaintRequest();
    }
    drawInpaintCanvas();
});

brushSize.addEventListener("input", function() {
    brushSizeLabel.textContent = brushSize.value + "px";
});

// Inpainting buttons
inpaintBtn.addEventListener("click", startInpaint);

inpaintUndoBtn.addEventListener("click", function() {
    if (inpaintHistory.length === 0) return;
    var prevSrc = inpaintHistory.pop();
    if (inpaintHistory.length === 0) {
        inpaintUndoBtn.disabled = true;
    }
    inpaintImg = new Image();
    inpaintImg.onload = function() {
        inpaintMaskCtx.fillStyle = "black";
        inpaintMaskCtx.fillRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
        drawInpaintCanvas();
    };
    inpaintImg.src = prevSrc;
});

inpaintDoneBtn.addEventListener("click", function() {
    inpaintSection.hidden = true;
    processedImg.src = (currentProcessedUrl || "") + "?t=" + Date.now();
    message.textContent = "부분 보정 완료!";
    message.style.color = "#2a7d2a";
    confirmBtn.hidden = false;
    manualBtn.hidden = false;
    inpaintBtn.hidden = false;
    preview.hidden = false;
});

inpaintCancelBtn.addEventListener("click", function() {
    inpaintSection.hidden = true;
    if (currentProcessedUrl) {
        preview.hidden = false;
    } else {
        failMessage.hidden = false;
    }
});
```

- [ ] **Step 6: Commit**

```bash
git add static/app.js
git commit -m "feat: 인페인팅 모드 JS 로직 (드래그 페인팅, LaMa API 호출, Undo)"
```

---

### Task 7: 통합 테스트 및 마무리

**Files:**
- All files

- [ ] **Step 1: 전체 테스트 실행**

Run: `cd gallery && python -m pytest tests/ -v`
Expected: 모든 테스트 PASS

- [ ] **Step 2: 개발 서버 실행 및 수동 확인**

Run: `cd gallery && python app.py`

브라우저에서 확인할 항목:
1. 사진 업로드 → 추출 결과 화면에 "부분 보정" 버튼이 "수동 지정" 옆에 표시되는지
2. "부분 보정" 클릭 → 인페인팅 모드 진입, 캔버스에 현재 이미지 표시
3. 브러시 크기 슬라이더 동작
4. 드래그로 영역 칠하기 → 빨간 마스크 표시
5. 마우스 떼면 "AI 보정 중..." 스피너 → 보정 결과 반영
6. Undo 버튼으로 되돌리기
7. 완료 버튼으로 미리보기 화면 복귀

- [ ] **Step 3: 최종 Commit**

```bash
git add -A
git commit -m "feat: AI 부분 보정(인페인팅) 기능 완성 — LaMa 기반 드래그 페인팅"
```
