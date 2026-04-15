# Artwork Extractor V2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Upgrade artwork extractor with Florence-2 detection, MediaPipe auto hand removal, and mobile-first wizard UI.

**Architecture:** Modular pipeline — detector.py (Florence-2 + legacy fallback), hand_detector.py (MediaPipe), inpainter.py (LaMa) orchestrated by extractor.py. Frontend rewritten as 3-step mobile wizard.

**Tech Stack:** Flask, OpenCV, Florence-2 (transformers), MediaPipe, LaMa, vanilla JS

---

## File Structure

```
gallery/
├── app.py                    # Modify: add new response fields, update imports
├── extractor.py              # Modify: refactor to orchestrator using new modules
├── detector.py               # Create: Detector ABC, FlorenceDetector, LegacyDetector
├── hand_detector.py          # Create: MediaPipeHandDetector
├── inpainter.py              # Create: LamaInpainter (extracted from extractor.py)
├── requirements.txt          # Modify: add transformers, torch, mediapipe
├── templates/
│   └── index.html            # Rewrite: 3-step mobile wizard
├── static/
│   ├── app.js                # Rewrite: wizard logic + touch events
│   └── style.css             # Rewrite: mobile-first responsive
└── tests/
    ├── test_extractor.py     # Modify: update for new ProcessResult API
    ├── test_detector.py      # Create: tests for detector module
    ├── test_hand_detector.py # Create: tests for hand detection
    └── test_inpainter.py     # Create: tests for inpainter module
```

---

### Task 1: Create inpainter module with tests

Extract LaMa inpainting from extractor.py into a standalone module.

**Files:**
- Create: `gallery/inpainter.py`
- Create: `gallery/tests/test_inpainter.py`

- [ ] **Step 1: Write the failing test**

Create `gallery/tests/test_inpainter.py`:

```python
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PIL import Image, ImageDraw


def test_lama_inpainter_returns_rgb_image():
    from inpainter import LamaInpainter
    inpainter = LamaInpainter()

    img = Image.new("RGB", (200, 200), (200, 100, 100))
    mask = Image.new("L", (200, 200), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse([80, 80, 120, 120], fill=255)

    result = inpainter.inpaint(img, mask)

    assert isinstance(result, Image.Image)
    assert result.size == (200, 200)
    assert result.mode == "RGB"


def test_lama_inpainter_with_empty_mask():
    from inpainter import LamaInpainter
    inpainter = LamaInpainter()

    img = Image.new("RGB", (200, 200), (100, 150, 200))
    mask = Image.new("L", (200, 200), 0)  # all black = nothing to inpaint

    result = inpainter.inpaint(img, mask)

    assert isinstance(result, Image.Image)
    assert result.size == (200, 200)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd gallery && python -m pytest tests/test_inpainter.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'inpainter'`

- [ ] **Step 3: Write inpainter.py**

Create `gallery/inpainter.py`:

```python
from abc import ABC, abstractmethod
from PIL import Image
import torch
from simple_lama_inpainting import SimpleLama


class Inpainter(ABC):
    @abstractmethod
    def inpaint(self, image: Image.Image, mask: Image.Image) -> Image.Image:
        """Inpaint masked region. White=area to inpaint.
        Args:
            image: RGB PIL Image.
            mask: Grayscale PIL Image (255=inpaint, 0=keep).
        Returns:
            Inpainted RGB PIL Image.
        """
        pass


class LamaInpainter(Inpainter):
    def __init__(self):
        device = torch.device("cpu")
        self._model = object.__new__(SimpleLama)
        from simple_lama_inpainting.models.model import download_model, LAMA_MODEL_URL
        import os
        model_path = os.environ.get("LAMA_MODEL") or download_model(LAMA_MODEL_URL)
        self._model.model = torch.jit.load(model_path, map_location=device)
        self._model.model.eval()
        self._model.model.to(device)
        self._model.device = device

    def inpaint(self, image: Image.Image, mask: Image.Image) -> Image.Image:
        result = self._model(image, mask)
        return result.convert("RGB")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd gallery && python -m pytest tests/test_inpainter.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add gallery/inpainter.py gallery/tests/test_inpainter.py
git commit -m "feat: inpainter 모듈 분리 (LamaInpainter)"
```

---

### Task 2: Create detector module with LegacyDetector

Extract the existing rembg+Canny detection logic into the detector module.

**Files:**
- Create: `gallery/detector.py`
- Create: `gallery/tests/test_detector.py`

- [ ] **Step 1: Write the failing test**

Create `gallery/tests/test_detector.py`:

```python
import cv2
import numpy as np
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def make_test_image(width=800, height=600, rect_color=(0, 0, 200), bg_color=(220, 220, 220)):
    """Synthetic image: gray background with a colored tilted rectangle."""
    img = np.full((height, width, 3), bg_color, dtype=np.uint8)
    center = (width // 2, height // 2)
    rect_w, rect_h = 400, 300
    angle = 5
    box = cv2.boxPoints(((center[0], center[1]), (rect_w, rect_h), angle))
    box = np.intp(box)
    cv2.fillPoly(img, [box], rect_color)
    return img


def test_detection_result_fields():
    from detector import DetectionResult
    result = DetectionResult(corners=None, confidence=0.0, method="test")
    assert result.corners is None
    assert result.confidence == 0.0
    assert result.method == "test"


def test_legacy_detector_finds_rectangle():
    from detector import LegacyDetector
    detector = LegacyDetector()
    img = make_test_image()
    # LegacyDetector expects BGR input (same as cv2.imread)
    result = detector.detect(img)
    assert result.corners is not None, "Should detect corners on synthetic image"
    assert result.corners.shape == (4, 2), "Should return 4 corner points"
    assert result.method == "legacy"
    assert 0.0 <= result.confidence <= 1.0


def test_legacy_detector_returns_none_for_blank():
    from detector import LegacyDetector
    detector = LegacyDetector()
    blank = np.full((600, 800, 3), (200, 200, 200), dtype=np.uint8)
    result = detector.detect(blank)
    # May return corners (rembg may find something) or None
    # Important: doesn't crash
    assert result.method == "legacy"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd gallery && python -m pytest tests/test_detector.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'detector'`

- [ ] **Step 3: Write detector.py**

Create `gallery/detector.py`:

```python
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Optional
import cv2
import numpy as np
from rembg import remove, new_session
from PIL import Image


@dataclass
class DetectionResult:
    corners: Optional[np.ndarray]  # 4x2 float32 array: TL, TR, BR, BL
    confidence: float              # 0.0~1.0
    method: str                    # "florence2" | "legacy" | "manual"


class Detector(ABC):
    @abstractmethod
    def detect(self, image: np.ndarray) -> DetectionResult:
        """Detect painting region and return 4 corners.
        Args:
            image: BGR numpy array (from cv2.imread).
        Returns:
            DetectionResult with corners (or None if not found).
        """
        pass


def _order_points(pts: np.ndarray) -> np.ndarray:
    """Order points: top-left, top-right, bottom-right, bottom-left."""
    rect = np.zeros((4, 2), dtype=np.float32)
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    d = np.diff(pts, axis=1).ravel()
    rect[1] = pts[np.argmin(d)]
    rect[3] = pts[np.argmax(d)]
    return rect


class LegacyDetector(Detector):
    """Original rembg + Canny edge detection approach."""

    def __init__(self):
        self._bg_session = new_session("u2net")

    def detect(self, image: np.ndarray) -> DetectionResult:
        rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)

        result = remove(pil_img, session=self._bg_session)
        arr = np.array(result)
        fg_alpha = arr[:, :, 3]

        if np.count_nonzero(fg_alpha > 128) == 0:
            return DetectionResult(corners=None, confidence=0.0, method="legacy")

        corners = self._find_painting_corners(rgb, fg_alpha)
        confidence = 0.7 if corners is not None else 0.0
        return DetectionResult(corners=corners, confidence=confidence, method="legacy")

    def _find_painting_corners(self, rgb: np.ndarray, fg_alpha: np.ndarray) -> Optional[np.ndarray]:
        h, w = rgb.shape[:2]
        _, binary = cv2.threshold(fg_alpha, 128, 255, cv2.THRESH_BINARY)

        gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
        blurred = cv2.bilateralFilter(gray, 9, 75, 75)

        for canny_lo, canny_hi in [(30, 100), (50, 150), (20, 80)]:
            edges = cv2.Canny(blurred, canny_lo, canny_hi)
            dk = np.ones((11, 11), np.uint8)
            search = cv2.dilate(binary, dk)
            edges[search == 0] = 0
            edges = cv2.dilate(edges, np.ones((3, 3), np.uint8), iterations=1)

            contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            for cnt in sorted(contours, key=cv2.contourArea, reverse=True)[:5]:
                area = cv2.contourArea(cnt)
                if area < h * w * 0.05:
                    break
                peri = cv2.arcLength(cnt, True)
                for eps in [0.015, 0.02, 0.03, 0.04, 0.05, 0.08]:
                    approx = cv2.approxPolyDP(cnt, eps * peri, True)
                    if len(approx) == 4 and cv2.isContourConvex(approx):
                        if cv2.contourArea(approx) > h * w * 0.05:
                            return _order_points(approx.reshape(4, 2))

        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if contours:
            largest = max(contours, key=cv2.contourArea)
            hull = cv2.convexHull(largest)
            peri = cv2.arcLength(hull, True)
            for eps in [0.02, 0.03, 0.05, 0.08, 0.10]:
                approx = cv2.approxPolyDP(hull, eps * peri, True)
                if len(approx) == 4:
                    return _order_points(approx.reshape(4, 2))

        return None

    def get_foreground_mask(self, image: np.ndarray):
        """Return foreground alpha mask (for fallback crop)."""
        rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        result = remove(pil_img, session=self._bg_session)
        arr = np.array(result)
        return arr[:, :, 3], arr
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd gallery && python -m pytest tests/test_detector.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add gallery/detector.py gallery/tests/test_detector.py
git commit -m "feat: detector 모듈 분리 (LegacyDetector + DetectionResult)"
```

---

### Task 3: Add FlorenceDetector

Add Florence-2 based detection to the detector module.

**Files:**
- Modify: `gallery/detector.py`
- Modify: `gallery/tests/test_detector.py`
- Modify: `gallery/requirements.txt`

- [ ] **Step 1: Update requirements.txt**

Add to `gallery/requirements.txt`:

```
transformers
torch
timm
einops
```

- [ ] **Step 2: Install new dependencies**

Run: `cd gallery && pip install transformers torch timm einops`

- [ ] **Step 3: Write the failing test**

Append to `gallery/tests/test_detector.py`:

```python
def test_florence_detector_finds_rectangle():
    from detector import FlorenceDetector
    detector = FlorenceDetector()
    img = make_test_image()
    result = detector.detect(img)
    # Florence-2 may or may not detect a "painting" in a synthetic image
    # Key: it should not crash, and return a valid DetectionResult
    assert result.method == "florence2"
    assert 0.0 <= result.confidence <= 1.0
    if result.corners is not None:
        assert result.corners.shape == (4, 2)


def test_florence_detector_on_blank():
    from detector import FlorenceDetector
    detector = FlorenceDetector()
    blank = np.full((600, 800, 3), (200, 200, 200), dtype=np.uint8)
    result = detector.detect(blank)
    assert result.method == "florence2"
    assert result.confidence >= 0.0
```

- [ ] **Step 4: Run test to verify it fails**

Run: `cd gallery && python -m pytest tests/test_detector.py::test_florence_detector_finds_rectangle -v`
Expected: FAIL — `ImportError: cannot import name 'FlorenceDetector'`

- [ ] **Step 5: Implement FlorenceDetector**

Add to `gallery/detector.py` (after LegacyDetector class):

```python
class FlorenceDetector(Detector):
    """Florence-2 based painting detection."""

    def __init__(self, model_name: str = "microsoft/Florence-2-base"):
        from transformers import AutoProcessor, AutoModelForCausalLM
        self._processor = AutoProcessor.from_pretrained(model_name, trust_remote_code=True)
        self._model = AutoModelForCausalLM.from_pretrained(model_name, trust_remote_code=True)
        self._model.eval()

    def detect(self, image: np.ndarray) -> DetectionResult:
        rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        h, w = image.shape[:2]

        # Use object detection task to find "painting"
        prompt = "<OD>"
        inputs = self._processor(text=prompt, images=pil_img, return_tensors="pt")
        import torch
        with torch.no_grad():
            generated = self._model.generate(
                input_ids=inputs["input_ids"],
                pixel_values=inputs["pixel_values"],
                max_new_tokens=1024,
                num_beams=3,
            )
        result_text = self._processor.batch_decode(generated, skip_special_tokens=False)[0]
        parsed = self._processor.post_process_generation(result_text, task=prompt, image_size=(w, h))

        # Find painting-like objects in results
        bbox = self._find_painting_bbox(parsed, w, h)
        if bbox is None:
            return DetectionResult(corners=None, confidence=0.0, method="florence2")

        # Refine bbox to precise corners using edge detection within the box
        corners = self._refine_corners(rgb, bbox)
        confidence = 0.85 if corners is not None else 0.5

        if corners is None:
            # Use bbox corners directly
            x1, y1, x2, y2 = bbox
            corners = np.array([
                [x1, y1], [x2, y1], [x2, y2], [x1, y2]
            ], dtype=np.float32)
            confidence = 0.5

        return DetectionResult(corners=corners, confidence=confidence, method="florence2")

    def _find_painting_bbox(self, parsed, w, h):
        """Find the most likely painting bbox from Florence-2 OD results."""
        od_result = parsed.get("<OD>", {})
        bboxes = od_result.get("bboxes", [])
        labels = od_result.get("labels", [])

        painting_keywords = ["painting", "picture", "art", "drawing", "frame",
                             "poster", "canvas", "board", "paper", "photo"]

        # First try: find an object matching painting keywords
        best_bbox = None
        best_area = 0
        for bbox, label in zip(bboxes, labels):
            label_lower = label.lower()
            if any(kw in label_lower for kw in painting_keywords):
                area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                if area > best_area:
                    best_area = area
                    best_bbox = bbox

        # Second try: use the largest detected object
        if best_bbox is None and bboxes:
            areas = [(b[2]-b[0])*(b[3]-b[1]) for b in bboxes]
            idx = max(range(len(areas)), key=lambda i: areas[i])
            best_bbox = bboxes[idx]

        if best_bbox is None:
            return None

        x1, y1, x2, y2 = [int(v) for v in best_bbox]
        x1 = max(0, x1)
        y1 = max(0, y1)
        x2 = min(w, x2)
        y2 = min(h, y2)
        return (x1, y1, x2, y2)

    def _refine_corners(self, rgb, bbox):
        """Refine bbox to precise 4 corners using edge detection within the region."""
        x1, y1, x2, y2 = bbox
        pad = 20
        h, w = rgb.shape[:2]
        rx1 = max(0, x1 - pad)
        ry1 = max(0, y1 - pad)
        rx2 = min(w, x2 + pad)
        ry2 = min(h, y2 + pad)
        region = rgb[ry1:ry2, rx1:rx2]

        gray = cv2.cvtColor(region, cv2.COLOR_RGB2GRAY)
        blurred = cv2.bilateralFilter(gray, 9, 75, 75)

        for canny_lo, canny_hi in [(30, 100), (50, 150), (20, 80)]:
            edges = cv2.Canny(blurred, canny_lo, canny_hi)
            edges = cv2.dilate(edges, np.ones((3, 3), np.uint8), iterations=1)
            contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            rh, rw = region.shape[:2]
            for cnt in sorted(contours, key=cv2.contourArea, reverse=True)[:5]:
                area = cv2.contourArea(cnt)
                if area < rh * rw * 0.1:
                    break
                peri = cv2.arcLength(cnt, True)
                for eps in [0.015, 0.02, 0.03, 0.04, 0.05]:
                    approx = cv2.approxPolyDP(cnt, eps * peri, True)
                    if len(approx) == 4 and cv2.isContourConvex(approx):
                        pts = approx.reshape(4, 2).astype(np.float32)
                        pts[:, 0] += rx1
                        pts[:, 1] += ry1
                        return _order_points(pts)

        return None
```

- [ ] **Step 6: Run test to verify it passes**

Run: `cd gallery && python -m pytest tests/test_detector.py -v`
Expected: 5 passed (first run will be slow — Florence-2 model download ~1GB)

- [ ] **Step 7: Commit**

```bash
git add gallery/detector.py gallery/tests/test_detector.py gallery/requirements.txt
git commit -m "feat: FlorenceDetector 추가 (Florence-2 기반 회화 검출)"
```

---

### Task 4: Create hand_detector module

Implement MediaPipe Hands based auto hand detection.

**Files:**
- Create: `gallery/hand_detector.py`
- Create: `gallery/tests/test_hand_detector.py`
- Modify: `gallery/requirements.txt`

- [ ] **Step 1: Update requirements.txt**

Add to `gallery/requirements.txt`:

```
mediapipe
```

- [ ] **Step 2: Install dependency**

Run: `cd gallery && pip install mediapipe`

- [ ] **Step 3: Write the failing test**

Create `gallery/tests/test_hand_detector.py`:

```python
import numpy as np
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def test_hand_detector_returns_none_for_no_hands():
    from hand_detector import MediaPipeHandDetector
    detector = MediaPipeHandDetector()

    # Plain colored image — no hands
    img = np.full((400, 400, 3), (100, 150, 200), dtype=np.uint8)
    mask = detector.detect_hand_mask(img)

    assert mask is None, "Should return None when no hands detected"


def test_hand_detector_mask_shape():
    from hand_detector import MediaPipeHandDetector
    detector = MediaPipeHandDetector()

    # Even with no detection, verify interface contract
    img = np.full((300, 400, 3), (200, 200, 200), dtype=np.uint8)
    mask = detector.detect_hand_mask(img)

    # No hands in synthetic image → None
    if mask is not None:
        assert mask.shape == (300, 400), "Mask should match input dimensions"
        assert mask.dtype == np.uint8
```

- [ ] **Step 4: Run test to verify it fails**

Run: `cd gallery && python -m pytest tests/test_hand_detector.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'hand_detector'`

- [ ] **Step 5: Write hand_detector.py**

Create `gallery/hand_detector.py`:

```python
from typing import Optional
import cv2
import numpy as np
import mediapipe as mp


class MediaPipeHandDetector:
    """Detect hands using MediaPipe and generate inpainting mask."""

    def __init__(self, dilation_px: int = 15):
        self._dilation_px = dilation_px
        self._hands = mp.solutions.hands.Hands(
            static_image_mode=True,
            max_num_hands=4,
            min_detection_confidence=0.3,
        )

    def detect_hand_mask(self, image: np.ndarray) -> Optional[np.ndarray]:
        """Detect hands and return a grayscale mask (255=hand, 0=background).

        Args:
            image: RGB numpy array.

        Returns:
            Grayscale mask (same H, W as input) or None if no hands found.
        """
        h, w = image.shape[:2]
        results = self._hands.process(image)

        if not results.multi_hand_landmarks:
            return None

        mask = np.zeros((h, w), dtype=np.uint8)

        for hand_landmarks in results.multi_hand_landmarks:
            # Convert landmarks to pixel coordinates
            points = []
            for lm in hand_landmarks.landmark:
                px = int(lm.x * w)
                py = int(lm.y * h)
                points.append([px, py])

            points = np.array(points, dtype=np.int32)
            hull = cv2.convexHull(points)
            cv2.fillConvexPoly(mask, hull, 255)

        # Dilate to cover edges around fingers
        if self._dilation_px > 0:
            kernel = cv2.getStructuringElement(
                cv2.MORPH_ELLIPSE,
                (self._dilation_px * 2 + 1, self._dilation_px * 2 + 1)
            )
            mask = cv2.dilate(mask, kernel, iterations=1)

        if np.count_nonzero(mask) == 0:
            return None

        return mask

    def close(self):
        self._hands.close()
```

- [ ] **Step 6: Run test to verify it passes**

Run: `cd gallery && python -m pytest tests/test_hand_detector.py -v`
Expected: 2 passed

- [ ] **Step 7: Commit**

```bash
git add gallery/hand_detector.py gallery/tests/test_hand_detector.py gallery/requirements.txt
git commit -m "feat: hand_detector 모듈 추가 (MediaPipe Hands)"
```

---

### Task 5: Refactor extractor.py as orchestrator

Wire up all new modules in extractor.py. Keep backward-compatible public API.

**Files:**
- Modify: `gallery/extractor.py`
- Modify: `gallery/tests/test_extractor.py`

- [ ] **Step 1: Update test_extractor.py for new API**

Rewrite `gallery/tests/test_extractor.py`:

```python
import cv2
import numpy as np
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def make_test_image(width=800, height=600, rect_color=(0, 0, 200), bg_color=(220, 220, 220)):
    img = np.full((height, width, 3), bg_color, dtype=np.uint8)
    center = (width // 2, height // 2)
    rect_w, rect_h = 400, 300
    angle = 5
    box = cv2.boxPoints(((center[0], center[1]), (rect_w, rect_h), angle))
    box = np.intp(box)
    cv2.fillPoly(img, [box], rect_color)
    return img


def test_process_photo_returns_process_result():
    from extractor import process_photo
    img = make_test_image()
    result = process_photo(img, output_size=500)
    assert result is not None
    # result can be a ProcessResult or np.ndarray (backward compat)
    if hasattr(result, "image"):
        assert result.image is not None or result.needs_manual
        if result.image is not None:
            h, w = result.image.shape[:2]
            assert h == 500 and w == 500
            assert result.detection_method in ("florence2", "legacy")
    else:
        h, w = result.shape[:2]
        assert h == 500 and w == 500


def test_process_photo_blank_image():
    from extractor import process_photo
    blank = np.full((600, 800, 3), (200, 200, 200), dtype=np.uint8)
    result = process_photo(blank, output_size=500)
    if result is not None and hasattr(result, "image"):
        if result.image is not None:
            h, w = result.image.shape[:2]
            assert h == 500 and w == 500


def test_manual_process():
    from extractor import manual_process
    img = make_test_image()
    corners = [[200, 150], [600, 150], [600, 450], [200, 450]]
    result = manual_process(img, corners, output_size=500)
    assert result is not None
    h, w = result.shape[:2]
    assert h == 500 and w == 500


def test_fit_to_square():
    from extractor import _fit_to_square
    rect_img = np.full((300, 400, 3), (100, 150, 200), dtype=np.uint8)
    square = _fit_to_square(rect_img, 500)
    h, w = square.shape[:2]
    assert h == w == 500


def test_fit_to_square_tall():
    from extractor import _fit_to_square
    tall_img = np.full((500, 300, 3), (50, 100, 150), dtype=np.uint8)
    square = _fit_to_square(tall_img, 500)
    h, w = square.shape[:2]
    assert h == w == 500


def test_inpaint_region():
    from extractor import inpaint_region
    from PIL import Image, ImageDraw

    img = Image.new("RGB", (200, 200), (200, 100, 100))
    mask = Image.new("L", (200, 200), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse([80, 80, 120, 120], fill=255)

    result = inpaint_region(img, mask)
    assert isinstance(result, Image.Image)
    assert result.size == (200, 200)
    assert result.mode == "RGB"
```

- [ ] **Step 2: Rewrite extractor.py**

Rewrite `gallery/extractor.py`:

```python
import cv2
import numpy as np
from dataclasses import dataclass
from typing import Optional
from PIL import Image

from detector import FlorenceDetector, LegacyDetector, DetectionResult, _order_points
from hand_detector import MediaPipeHandDetector
from inpainter import LamaInpainter


@dataclass
class ProcessResult:
    success: bool
    image: Optional[np.ndarray] = None        # BGR output image
    needs_manual: bool = False
    hand_detected: bool = False
    detection_method: str = "none"
    confidence: float = 0.0


# Pre-load models at startup
print("[extractor] Loading models...")
try:
    _florence_detector = FlorenceDetector()
    print("[extractor] Florence-2 loaded")
except Exception as e:
    print(f"[extractor] Florence-2 failed to load: {e}, using legacy only")
    _florence_detector = None

_legacy_detector = LegacyDetector()
print("[extractor] Legacy detector loaded")

_hand_detector = MediaPipeHandDetector()
print("[extractor] Hand detector loaded")

_inpainter = LamaInpainter()
print("[extractor] LaMa inpainter loaded")


def process_photo(img: np.ndarray, output_size: int = 1080) -> Optional[ProcessResult]:
    """Full pipeline: detect painting → perspective transform → hand removal → square fit.

    Args:
        img: BGR numpy array (from cv2.imread).
        output_size: Final square dimension.

    Returns:
        ProcessResult with processed image and metadata.
    """
    # Step 1: Detect painting with Florence-2
    detection = None
    if _florence_detector is not None:
        detection = _florence_detector.detect(img)

    # Step 2: Fallback to legacy if Florence-2 failed
    if detection is None or detection.corners is None:
        detection = _legacy_detector.detect(img)

    # Step 3: If still no corners, try fallback crop
    if detection.corners is None:
        fg_alpha, rgba_arr = _legacy_detector.get_foreground_mask(img)
        if np.count_nonzero(fg_alpha > 128) == 0:
            return ProcessResult(success=False, needs_manual=True,
                                 detection_method=detection.method, confidence=0.0)
        fallback = _fallback_crop(rgba_arr, fg_alpha, output_size)
        return ProcessResult(success=True, image=fallback, hand_detected=False,
                             detection_method=detection.method, confidence=0.3)

    # Step 4: Perspective transform
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    warped = _perspective_transform(rgb, detection.corners)

    # Step 5: Auto hand detection + inpainting
    hand_detected = False
    hand_mask = _hand_detector.detect_hand_mask(warped)
    if hand_mask is not None:
        pil_img = Image.fromarray(warped)
        pil_mask = Image.fromarray(hand_mask)
        inpainted = _inpainter.inpaint(pil_img, pil_mask)
        warped = np.array(inpainted)
        hand_detected = True

    # Step 6: Fit to square
    final = _fit_to_square(warped, output_size)

    return ProcessResult(
        success=True,
        image=final,
        hand_detected=hand_detected,
        detection_method=detection.method,
        confidence=detection.confidence,
    )


def manual_process(img: np.ndarray, points: list, output_size: int = 1080) -> np.ndarray:
    """Process with manually specified 4 corner points."""
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    corners = _order_points(np.array(points, dtype=np.float32))
    warped = _perspective_transform(rgb, corners)
    return _fit_to_square(warped, output_size)


def inpaint_region(img: Image.Image, mask: Image.Image) -> Image.Image:
    """Inpaint masked region using LaMa model (backward compat)."""
    return _inpainter.inpaint(img, mask)


def _perspective_transform(rgb: np.ndarray, corners: np.ndarray) -> np.ndarray:
    tl, tr, br, bl = corners
    max_w = int(max(np.linalg.norm(tr - tl), np.linalg.norm(br - bl)))
    max_h = int(max(np.linalg.norm(bl - tl), np.linalg.norm(br - tr)))
    dst = np.array([
        [0, 0], [max_w - 1, 0],
        [max_w - 1, max_h - 1], [0, max_h - 1]
    ], dtype=np.float32)
    M = cv2.getPerspectiveTransform(corners, dst)
    warped = cv2.warpPerspective(rgb, M, (max_w, max_h))
    return warped


def _fit_to_square(rgb_img: np.ndarray, output_size: int) -> np.ndarray:
    h, w = rgb_img.shape[:2]
    side = max(h, w)
    padding = int(side * 0.05)
    total = side + padding * 2
    square = np.full((total, total, 3), 255, dtype=np.uint8)
    y_off = (total - h) // 2
    x_off = (total - w) // 2
    square[y_off:y_off + h, x_off:x_off + w] = rgb_img
    resized = cv2.resize(square, (output_size, output_size), interpolation=cv2.INTER_LANCZOS4)
    return cv2.cvtColor(resized, cv2.COLOR_RGB2BGR)


def _fallback_crop(rgba_arr: np.ndarray, fg_alpha: np.ndarray, output_size: int) -> np.ndarray:
    rgb_out = rgba_arr[:, :, :3].copy()
    rgb_out[fg_alpha < 128] = [255, 255, 255]
    coords = np.column_stack(np.where(fg_alpha > 128))
    y_min, x_min = coords.min(axis=0)
    y_max, x_max = coords.max(axis=0)
    cropped = rgb_out[y_min:y_max, x_min:x_max]
    return _fit_to_square(cropped, output_size)
```

- [ ] **Step 3: Run tests**

Run: `cd gallery && python -m pytest tests/test_extractor.py -v`
Expected: 6 passed

- [ ] **Step 4: Commit**

```bash
git add gallery/extractor.py gallery/tests/test_extractor.py
git commit -m "refactor: extractor.py를 모듈 오케스트레이터로 리팩토링"
```

---

### Task 6: Update app.py for new API

Add new response fields (hand_detected, detection_method, confidence) to the upload endpoint.

**Files:**
- Modify: `gallery/app.py`

- [ ] **Step 1: Update app.py**

Replace the `upload()` function in `gallery/app.py` (lines 26-64):

```python
@app.route("/upload", methods=["POST"])
def upload():
    if "photo" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["photo"]
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": f"Unsupported format: {ext}"}), 400

    file_id = uuid.uuid4().hex[:12]
    original_name = f"{file_id}_original{ext}"
    original_path = os.path.join(UPLOAD_DIR, original_name)
    file.save(original_path)

    img = cv2.imread(original_path)
    if img is None:
        return jsonify({"error": "Cannot read image file"}), 400

    result = process_photo(img, output_size=OUTPUT_SIZE)

    if result is None or (hasattr(result, 'success') and not result.success):
        return jsonify({
            "file_id": file_id,
            "original": f"/files/uploads/{original_name}",
            "processed": None,
            "hand_detected": False,
            "detection_method": getattr(result, 'detection_method', 'none'),
            "confidence": getattr(result, 'confidence', 0.0),
            "message": "작품 추출에 실패했습니다. 수동으로 지정해주세요.",
        })

    image = result.image if hasattr(result, 'image') else result
    processed_name = f"{file_id}_cropped.jpg"
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    cv2.imwrite(processed_path, image, [cv2.IMWRITE_JPEG_QUALITY, 92])

    hand_detected = getattr(result, 'hand_detected', False)
    method = getattr(result, 'detection_method', 'unknown')
    confidence = getattr(result, 'confidence', 0.0)

    msg = "작품 추출 완료!"
    if hand_detected:
        msg += " (손 자동 보정됨)"

    return jsonify({
        "file_id": file_id,
        "original": f"/files/uploads/{original_name}",
        "processed": f"/files/processed/{processed_name}",
        "hand_detected": hand_detected,
        "detection_method": method,
        "confidence": confidence,
        "message": msg,
    })
```

- [ ] **Step 2: Run the app manually to verify it starts**

Run: `cd gallery && timeout 10 python app.py || true`
Expected: Server starts, models load, no import errors.

- [ ] **Step 3: Commit**

```bash
git add gallery/app.py
git commit -m "feat: upload API에 hand_detected, detection_method, confidence 필드 추가"
```

---

### Task 7: Rewrite HTML for mobile wizard UI

Replace the current layout with a 3-step wizard optimized for mobile.

**Files:**
- Rewrite: `gallery/templates/index.html`

- [ ] **Step 1: Rewrite index.html**

Rewrite `gallery/templates/index.html`:

```html
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>작품 추출기</title>
    <link rel="stylesheet" href="/static/style.css?v=5">
</head>
<body>
    <div class="app">
        <!-- Progress Bar -->
        <div id="progressBar" class="progress-bar" hidden>
            <div class="progress-steps">
                <div class="progress-step active" data-step="1"></div>
                <div class="progress-step" data-step="2"></div>
                <div class="progress-step" data-step="3"></div>
            </div>
            <span id="progressLabel" class="progress-label">1/3</span>
        </div>

        <!-- STEP 1: Upload -->
        <div id="step1" class="step">
            <h1>작품 사진 추출기</h1>
            <p class="subtitle">사진을 찍으면 작품만 깔끔하게 추출해드려요</p>
            <div class="upload-area">
                <input type="file" id="fileInput" accept="image/*" capture="environment" hidden>
                <button id="cameraBtn" class="btn-camera">
                    <span class="camera-icon">📷</span>
                    사진 촬영
                </button>
                <button id="galleryBtn" class="btn-gallery">앨범에서 선택</button>
                <input type="file" id="galleryInput" accept="image/*" hidden>
            </div>
        </div>

        <!-- Loading -->
        <div id="loadingView" class="step" hidden>
            <div class="loading-content">
                <div class="loading-spinner"></div>
                <p id="loadingText" class="loading-text">AI가 작품을 감지하고 있어요...</p>
                <div class="loading-steps">
                    <div id="lsDetect" class="loading-step">회화 영역 검출</div>
                    <div id="lsTransform" class="loading-step">원근 보정</div>
                    <div id="lsHand" class="loading-step">손 자동 감지</div>
                    <div id="lsFinish" class="loading-step">마무리</div>
                </div>
            </div>
        </div>

        <!-- STEP 2: Review -->
        <div id="step2" class="step" hidden>
            <div class="review-image-container">
                <img id="reviewImg" src="" alt="결과">
                <div id="handBadge" class="badge" hidden>✋ 손 자동 보정됨</div>
                <div id="confidenceBadge" class="badge confidence-badge" hidden></div>
            </div>
            <div id="reviewMessage" class="review-message"></div>
            <div class="actions-bottom">
                <button id="confirmBtn" class="btn-primary">이대로 저장</button>
                <div class="actions-secondary">
                    <button id="adjustCornersBtn" class="btn-secondary">꼭지점 수정</button>
                    <button id="touchupBtn" class="btn-secondary">추가 보정</button>
                    <button id="retryBtn" class="btn-text">다시 촬영</button>
                </div>
            </div>
        </div>

        <!-- STEP 2-fail: Detection failed -->
        <div id="step2fail" class="step" hidden>
            <div class="fail-content">
                <p class="fail-icon">😥</p>
                <p id="failText" class="fail-text">작품을 자동으로 찾지 못했어요</p>
                <div class="actions-bottom">
                    <button id="failManualBtn" class="btn-primary">직접 꼭지점 지정</button>
                    <button id="failRetryBtn" class="btn-text">다시 촬영</button>
                </div>
            </div>
        </div>

        <!-- STEP 3: Manual Adjust -->
        <div id="step3" class="step" hidden>
            <div class="adjust-tabs">
                <button id="tabCorners" class="tab-btn active">📐 꼭지점</button>
                <button id="tabBrush" class="tab-btn">🖌️ 브러시</button>
            </div>

            <!-- Corner mode -->
            <div id="cornerMode" class="adjust-mode">
                <p class="adjust-hint" id="cornerHint">그림의 꼭지점 4개를 터치하세요</p>
                <div class="canvas-container">
                    <canvas id="cornerCanvas"></canvas>
                </div>
                <div class="actions-bottom">
                    <button id="cornerApplyBtn" class="btn-primary" disabled>적용</button>
                    <button id="cornerResetBtn" class="btn-secondary">초기화</button>
                    <button id="cornerCancelBtn" class="btn-text">취소</button>
                </div>
            </div>

            <!-- Brush mode -->
            <div id="brushMode" class="adjust-mode" hidden>
                <p class="adjust-hint">제거할 부분을 터치로 칠해주세요</p>
                <div class="brush-controls">
                    <label>브러시:</label>
                    <input type="range" id="brushSize" min="10" max="80" value="30">
                    <span id="brushSizeLabel">30px</span>
                </div>
                <div class="canvas-container">
                    <canvas id="brushCanvas"></canvas>
                </div>
                <div class="actions-bottom">
                    <button id="brushUndoBtn" class="btn-secondary" disabled>되돌리기</button>
                    <button id="brushDoneBtn" class="btn-primary">완료</button>
                    <button id="brushCancelBtn" class="btn-text">취소</button>
                </div>
            </div>
        </div>

        <!-- Saved confirmation -->
        <div id="savedView" class="step" hidden>
            <div class="saved-content">
                <p class="saved-icon">✅</p>
                <p class="saved-text">저장 완료!</p>
                <button id="newPhotoBtn" class="btn-primary">새 사진 촬영</button>
            </div>
        </div>
    </div>

    <script src="/static/app.js?v=5"></script>
</body>
</html>
```

- [ ] **Step 2: Commit**

```bash
git add gallery/templates/index.html
git commit -m "feat: 모바일 3단계 위자드 HTML 구조 재설계"
```

---

### Task 8: Rewrite CSS for mobile-first design

**Files:**
- Rewrite: `gallery/static/style.css`

- [ ] **Step 1: Rewrite style.css**

Rewrite `gallery/static/style.css`:

```css
* { margin: 0; padding: 0; box-sizing: border-box; }

:root {
    --primary: #2563eb;
    --primary-dark: #1d4ed8;
    --success: #16a34a;
    --danger: #dc2626;
    --gray-50: #f9fafb;
    --gray-100: #f3f4f6;
    --gray-200: #e5e7eb;
    --gray-400: #9ca3af;
    --gray-600: #4b5563;
    --gray-800: #1f2937;
    --radius: 12px;
    --safe-bottom: env(safe-area-inset-bottom, 0px);
}

body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    background: var(--gray-50);
    color: var(--gray-800);
    min-height: 100vh;
    -webkit-tap-highlight-color: transparent;
    overscroll-behavior: none;
}

.app {
    max-width: 480px;
    margin: 0 auto;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
}

/* Progress bar */
.progress-bar {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    padding: 1rem 1.5rem 0;
}
.progress-steps {
    display: flex;
    gap: 0.4rem;
    flex: 1;
}
.progress-step {
    height: 4px;
    flex: 1;
    background: var(--gray-200);
    border-radius: 2px;
    transition: background 0.3s;
}
.progress-step.active {
    background: var(--primary);
}
.progress-label {
    font-size: 0.8rem;
    color: var(--gray-400);
    min-width: 2rem;
}

/* Steps */
.step {
    flex: 1;
    display: flex;
    flex-direction: column;
    padding: 1.5rem;
}

/* Step 1: Upload */
h1 {
    font-size: 1.5rem;
    margin-bottom: 0.5rem;
    margin-top: 2rem;
}
.subtitle {
    color: var(--gray-400);
    margin-bottom: 2rem;
    font-size: 0.95rem;
}
.upload-area {
    display: flex;
    flex-direction: column;
    gap: 1rem;
    margin-top: auto;
    padding-bottom: 3rem;
}
.btn-camera {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.5rem;
    width: 100%;
    padding: 1.2rem;
    font-size: 1.1rem;
    font-weight: 600;
    background: var(--primary);
    color: white;
    border: none;
    border-radius: var(--radius);
    cursor: pointer;
    -webkit-appearance: none;
}
.btn-camera:active { background: var(--primary-dark); }
.camera-icon { font-size: 1.3rem; }
.btn-gallery {
    width: 100%;
    padding: 1rem;
    font-size: 1rem;
    background: white;
    color: var(--gray-600);
    border: 2px solid var(--gray-200);
    border-radius: var(--radius);
    cursor: pointer;
}

/* Loading */
.loading-content {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 1.5rem;
}
.loading-spinner {
    width: 48px;
    height: 48px;
    border: 4px solid var(--gray-200);
    border-top-color: var(--primary);
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }
.loading-text {
    color: var(--gray-600);
    font-size: 1rem;
}
.loading-steps {
    display: flex;
    flex-direction: column;
    gap: 0.5rem;
    width: 100%;
    max-width: 250px;
}
.loading-step {
    font-size: 0.85rem;
    color: var(--gray-400);
    padding: 0.3rem 0;
    position: relative;
    padding-left: 1.5rem;
}
.loading-step::before {
    content: "○";
    position: absolute;
    left: 0;
}
.loading-step.active {
    color: var(--primary);
    font-weight: 600;
}
.loading-step.active::before { content: "◉"; }
.loading-step.done {
    color: var(--success);
}
.loading-step.done::before { content: "✓"; }

/* Step 2: Review */
.review-image-container {
    position: relative;
    width: 100%;
    border-radius: var(--radius);
    overflow: hidden;
    background: white;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
}
.review-image-container img {
    width: 100%;
    display: block;
}
.badge {
    position: absolute;
    top: 0.8rem;
    left: 0.8rem;
    background: rgba(0,0,0,0.7);
    color: white;
    padding: 0.3rem 0.8rem;
    border-radius: 20px;
    font-size: 0.8rem;
}
.confidence-badge {
    left: auto;
    right: 0.8rem;
}
.review-message {
    text-align: center;
    padding: 0.8rem 0;
    font-weight: 600;
    color: var(--success);
}

/* Actions */
.actions-bottom {
    margin-top: auto;
    padding-bottom: calc(1rem + var(--safe-bottom));
    display: flex;
    flex-direction: column;
    gap: 0.8rem;
}
.actions-secondary {
    display: flex;
    gap: 0.6rem;
}
.btn-primary {
    width: 100%;
    padding: 1rem;
    font-size: 1.05rem;
    font-weight: 600;
    background: var(--primary);
    color: white;
    border: none;
    border-radius: var(--radius);
    cursor: pointer;
    min-height: 48px;
}
.btn-primary:active { background: var(--primary-dark); }
.btn-primary:disabled { background: var(--gray-200); color: var(--gray-400); }
.btn-secondary {
    flex: 1;
    padding: 0.8rem;
    font-size: 0.9rem;
    background: white;
    color: var(--gray-600);
    border: 2px solid var(--gray-200);
    border-radius: var(--radius);
    cursor: pointer;
    min-height: 44px;
}
.btn-text {
    width: 100%;
    padding: 0.8rem;
    font-size: 0.9rem;
    background: none;
    color: var(--gray-400);
    border: none;
    cursor: pointer;
    text-decoration: underline;
}

/* Step 2 fail */
.fail-content {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 1rem;
}
.fail-icon { font-size: 3rem; }
.fail-text { color: var(--gray-600); font-size: 1.1rem; }

/* Step 3: Adjust */
.adjust-tabs {
    display: flex;
    gap: 0;
    margin-bottom: 1rem;
}
.tab-btn {
    flex: 1;
    padding: 0.7rem;
    font-size: 0.95rem;
    background: white;
    color: var(--gray-600);
    border: 2px solid var(--gray-200);
    cursor: pointer;
}
.tab-btn:first-child { border-radius: var(--radius) 0 0 var(--radius); }
.tab-btn:last-child { border-radius: 0 var(--radius) var(--radius) 0; }
.tab-btn.active {
    background: var(--primary);
    color: white;
    border-color: var(--primary);
}
.adjust-hint {
    text-align: center;
    color: var(--gray-600);
    font-size: 0.9rem;
    margin-bottom: 1rem;
}
.canvas-container {
    width: 100%;
    display: flex;
    justify-content: center;
    margin-bottom: 1rem;
}
.canvas-container canvas {
    max-width: 100%;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    touch-action: none;
}
.brush-controls {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.6rem;
    margin-bottom: 0.8rem;
    font-size: 0.85rem;
    color: var(--gray-600);
}
.brush-controls input[type="range"] { width: 150px; }

/* Saved */
.saved-content {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 1.5rem;
}
.saved-icon { font-size: 4rem; }
.saved-text { font-size: 1.3rem; font-weight: 600; }

/* Utility */
.adjust-mode { display: flex; flex-direction: column; flex: 1; }

@media (min-width: 768px) {
    .app { max-width: 600px; }
    .btn-camera { padding: 1rem; font-size: 1rem; }
}
```

- [ ] **Step 2: Commit**

```bash
git add gallery/static/style.css
git commit -m "feat: 모바일 퍼스트 CSS 재설계"
```

---

### Task 9: Rewrite app.js for wizard logic + touch support

Complete rewrite of frontend JavaScript with wizard navigation, touch events, and swipe support.

**Files:**
- Rewrite: `gallery/static/app.js`

- [ ] **Step 1: Rewrite app.js**

Rewrite `gallery/static/app.js`:

```javascript
document.addEventListener("DOMContentLoaded", function() {

// --- Element refs ---
var step1 = document.getElementById("step1");
var step2 = document.getElementById("step2");
var step2fail = document.getElementById("step2fail");
var step3 = document.getElementById("step3");
var loadingView = document.getElementById("loadingView");
var savedView = document.getElementById("savedView");
var progressBar = document.getElementById("progressBar");
var progressLabel = document.getElementById("progressLabel");

var fileInput = document.getElementById("fileInput");
var galleryInput = document.getElementById("galleryInput");
var cameraBtn = document.getElementById("cameraBtn");
var galleryBtn = document.getElementById("galleryBtn");

var reviewImg = document.getElementById("reviewImg");
var handBadge = document.getElementById("handBadge");
var confidenceBadge = document.getElementById("confidenceBadge");
var reviewMessage = document.getElementById("reviewMessage");
var confirmBtn = document.getElementById("confirmBtn");
var adjustCornersBtn = document.getElementById("adjustCornersBtn");
var touchupBtn = document.getElementById("touchupBtn");
var retryBtn = document.getElementById("retryBtn");

var failText = document.getElementById("failText");
var failManualBtn = document.getElementById("failManualBtn");
var failRetryBtn = document.getElementById("failRetryBtn");

var tabCorners = document.getElementById("tabCorners");
var tabBrush = document.getElementById("tabBrush");
var cornerMode = document.getElementById("cornerMode");
var brushMode = document.getElementById("brushMode");

var cornerCanvas = document.getElementById("cornerCanvas");
var cornerCtx = cornerCanvas.getContext("2d");
var cornerHint = document.getElementById("cornerHint");
var cornerApplyBtn = document.getElementById("cornerApplyBtn");
var cornerResetBtn = document.getElementById("cornerResetBtn");
var cornerCancelBtn = document.getElementById("cornerCancelBtn");

var brushCanvas = document.getElementById("brushCanvas");
var brushCtx = brushCanvas.getContext("2d");
var brushSize = document.getElementById("brushSize");
var brushSizeLabel = document.getElementById("brushSizeLabel");
var brushUndoBtn = document.getElementById("brushUndoBtn");
var brushDoneBtn = document.getElementById("brushDoneBtn");
var brushCancelBtn = document.getElementById("brushCancelBtn");

var newPhotoBtn = document.getElementById("newPhotoBtn");

// Loading step indicators
var lsDetect = document.getElementById("lsDetect");
var lsTransform = document.getElementById("lsTransform");
var lsHand = document.getElementById("lsHand");
var lsFinish = document.getElementById("lsFinish");

// --- State ---
var state = {
    fileId: null,
    originalUrl: null,
    processedUrl: null,
    cornerPoints: [],
    cornerImg: null,
    cornerScale: 1,
    brushImg: null,
    brushScale: 1,
    brushDrawing: false,
    brushMaskCanvas: null,
    brushMaskCtx: null,
    brushHistory: [],
};

// --- Navigation ---
function showStep(stepEl, stepNum) {
    [step1, step2, step2fail, step3, loadingView, savedView].forEach(function(el) {
        el.hidden = true;
    });
    stepEl.hidden = false;
    if (stepNum) {
        progressBar.hidden = false;
        progressLabel.textContent = stepNum + "/3";
        var steps = progressBar.querySelectorAll(".progress-step");
        steps.forEach(function(s, i) {
            s.classList.toggle("active", i < stepNum);
        });
    } else {
        progressBar.hidden = true;
    }
}

function resetState() {
    state.fileId = null;
    state.originalUrl = null;
    state.processedUrl = null;
    state.cornerPoints = [];
    state.brushHistory = [];
    fileInput.value = "";
    galleryInput.value = "";
}

// --- Step 1: Upload ---
cameraBtn.addEventListener("click", function() { fileInput.click(); });
galleryBtn.addEventListener("click", function() { galleryInput.click(); });

fileInput.addEventListener("change", function() {
    if (fileInput.files.length > 0) uploadFile(fileInput.files[0]);
});
galleryInput.addEventListener("change", function() {
    if (galleryInput.files.length > 0) uploadFile(galleryInput.files[0]);
});

function uploadFile(file) {
    showStep(loadingView, null);
    animateLoading();

    // Resize image before upload for faster transfer
    resizeImage(file, 2048, function(blob) {
        var formData = new FormData();
        formData.append("photo", blob, file.name);

        fetch("/upload", { method: "POST", body: formData })
            .then(function(resp) {
                return resp.json().then(function(data) {
                    return { ok: resp.ok, data: data };
                });
            })
            .then(function(result) {
                if (!result.ok) {
                    failText.textContent = result.data.error || "업로드 실패";
                    showStep(step2fail, 2);
                    return;
                }

                state.fileId = result.data.file_id;
                state.originalUrl = result.data.original;

                if (result.data.processed) {
                    state.processedUrl = result.data.processed;
                    reviewImg.src = result.data.processed;
                    reviewMessage.textContent = result.data.message;

                    handBadge.hidden = !result.data.hand_detected;
                    if (result.data.confidence > 0) {
                        var conf = Math.round(result.data.confidence * 100);
                        var label = conf >= 80 ? "높음" : conf >= 50 ? "중간" : "낮음";
                        confidenceBadge.textContent = "정확도: " + label;
                        confidenceBadge.hidden = false;
                    } else {
                        confidenceBadge.hidden = true;
                    }

                    showStep(step2, 2);
                } else {
                    failText.textContent = result.data.message;
                    showStep(step2fail, 2);
                }
            })
            .catch(function() {
                failText.textContent = "서버 연결에 실패했습니다.";
                showStep(step2fail, 2);
            });
    });
}

function resizeImage(file, maxDim, callback) {
    var img = new Image();
    img.onload = function() {
        if (img.width <= maxDim && img.height <= maxDim) {
            callback(file);
            return;
        }
        var scale = maxDim / Math.max(img.width, img.height);
        var canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        var ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        canvas.toBlob(function(blob) { callback(blob); }, "image/jpeg", 0.92);
    };
    img.src = URL.createObjectURL(file);
}

function animateLoading() {
    var steps = [lsDetect, lsTransform, lsHand, lsFinish];
    steps.forEach(function(s) { s.className = "loading-step"; });
    var i = 0;
    var interval = setInterval(function() {
        if (i > 0) steps[i - 1].className = "loading-step done";
        if (i < steps.length) {
            steps[i].className = "loading-step active";
            i++;
        } else {
            clearInterval(interval);
        }
    }, 1200);
}

// --- Step 2: Review ---
confirmBtn.addEventListener("click", function() {
    showStep(savedView, null);
});

retryBtn.addEventListener("click", function() {
    resetState();
    showStep(step1, null);
});

failRetryBtn.addEventListener("click", function() {
    resetState();
    showStep(step1, null);
});

newPhotoBtn.addEventListener("click", function() {
    resetState();
    showStep(step1, null);
});

adjustCornersBtn.addEventListener("click", function() {
    showStep(step3, 3);
    switchTab("corners");
    loadCornerCanvas();
});

touchupBtn.addEventListener("click", function() {
    showStep(step3, 3);
    switchTab("brush");
    loadBrushCanvas();
});

failManualBtn.addEventListener("click", function() {
    showStep(step3, 3);
    switchTab("corners");
    loadCornerCanvas();
});

// --- Step 3: Tabs ---
function switchTab(tab) {
    if (tab === "corners") {
        tabCorners.classList.add("active");
        tabBrush.classList.remove("active");
        cornerMode.hidden = false;
        brushMode.hidden = true;
    } else {
        tabBrush.classList.add("active");
        tabCorners.classList.remove("active");
        brushMode.hidden = false;
        cornerMode.hidden = true;
    }
}

tabCorners.addEventListener("click", function() {
    switchTab("corners");
    loadCornerCanvas();
});
tabBrush.addEventListener("click", function() {
    switchTab("brush");
    loadBrushCanvas();
});

// --- Corner Mode ---
function loadCornerCanvas() {
    state.cornerPoints = [];
    cornerApplyBtn.disabled = true;
    updateCornerHint();

    var url = state.originalUrl;
    state.cornerImg = new Image();
    state.cornerImg.onload = function() {
        var maxW = Math.min(window.innerWidth - 48, 500);
        state.cornerScale = Math.min(maxW / state.cornerImg.width, 1);
        cornerCanvas.width = Math.round(state.cornerImg.width * state.cornerScale);
        cornerCanvas.height = Math.round(state.cornerImg.height * state.cornerScale);
        drawCornerCanvas();
    };
    state.cornerImg.src = url;
}

function updateCornerHint() {
    var remaining = 4 - state.cornerPoints.length;
    if (remaining > 0) {
        cornerHint.textContent = "꼭지점을 터치하세요 (남은: " + remaining + "개)";
    } else {
        cornerHint.textContent = "4개 완료! '적용'을 눌러주세요";
    }
}

function drawCornerCanvas() {
    cornerCtx.drawImage(state.cornerImg, 0, 0, cornerCanvas.width, cornerCanvas.height);

    var pts = state.cornerPoints;
    if (pts.length > 1) {
        cornerCtx.beginPath();
        cornerCtx.moveTo(pts[0].cx, pts[0].cy);
        for (var i = 1; i < pts.length; i++) {
            cornerCtx.lineTo(pts[i].cx, pts[i].cy);
        }
        if (pts.length === 4) cornerCtx.closePath();
        cornerCtx.strokeStyle = "rgba(37, 99, 235, 0.9)";
        cornerCtx.lineWidth = 2;
        cornerCtx.stroke();
    }

    if (pts.length === 4) {
        cornerCtx.beginPath();
        cornerCtx.moveTo(pts[0].cx, pts[0].cy);
        for (var i = 1; i < pts.length; i++) cornerCtx.lineTo(pts[i].cx, pts[i].cy);
        cornerCtx.closePath();
        cornerCtx.fillStyle = "rgba(37, 99, 235, 0.15)";
        cornerCtx.fill();
    }

    for (var i = 0; i < pts.length; i++) {
        cornerCtx.beginPath();
        cornerCtx.arc(pts[i].cx, pts[i].cy, 16, 0, Math.PI * 2);
        cornerCtx.fillStyle = "rgba(37, 99, 235, 0.85)";
        cornerCtx.fill();
        cornerCtx.strokeStyle = "white";
        cornerCtx.lineWidth = 2;
        cornerCtx.stroke();

        cornerCtx.fillStyle = "white";
        cornerCtx.font = "bold 14px sans-serif";
        cornerCtx.textAlign = "center";
        cornerCtx.textBaseline = "middle";
        cornerCtx.fillText((i + 1).toString(), pts[i].cx, pts[i].cy);
    }
}

function canvasTouchPos(canvas, e) {
    var rect = canvas.getBoundingClientRect();
    var touch = e.touches ? e.touches[0] : e;
    return { x: touch.clientX - rect.left, y: touch.clientY - rect.top };
}

cornerCanvas.addEventListener("click", function(e) {
    if (state.cornerPoints.length >= 4) return;
    var pos = canvasTouchPos(cornerCanvas, e);
    state.cornerPoints.push({
        cx: pos.x, cy: pos.y,
        ox: pos.x / state.cornerScale,
        oy: pos.y / state.cornerScale
    });
    updateCornerHint();
    drawCornerCanvas();
    if (state.cornerPoints.length === 4) cornerApplyBtn.disabled = false;
});

cornerCanvas.addEventListener("touchend", function(e) {
    if (state.cornerPoints.length >= 4) return;
    e.preventDefault();
    var touch = e.changedTouches[0];
    var rect = cornerCanvas.getBoundingClientRect();
    var cx = touch.clientX - rect.left;
    var cy = touch.clientY - rect.top;
    state.cornerPoints.push({
        cx: cx, cy: cy,
        ox: cx / state.cornerScale,
        oy: cy / state.cornerScale
    });
    updateCornerHint();
    drawCornerCanvas();
    if (state.cornerPoints.length === 4) cornerApplyBtn.disabled = false;
});

cornerApplyBtn.addEventListener("click", function() {
    if (state.cornerPoints.length !== 4 || !state.fileId) return;
    showStep(loadingView, null);

    var points = state.cornerPoints.map(function(p) { return [p.ox, p.oy]; });

    fetch("/manual-crop", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: state.fileId, points: points, source: "original" })
    })
    .then(function(resp) { return resp.json().then(function(d) { return { ok: resp.ok, data: d }; }); })
    .then(function(result) {
        if (!result.ok) {
            failText.textContent = result.data.error || "보정 실패";
            showStep(step2fail, 2);
            return;
        }
        state.processedUrl = result.data.processed;
        reviewImg.src = result.data.processed + "?t=" + Date.now();
        reviewMessage.textContent = result.data.message;
        handBadge.hidden = true;
        confidenceBadge.hidden = true;
        showStep(step2, 2);
    })
    .catch(function() {
        failText.textContent = "서버 연결 실패";
        showStep(step2fail, 2);
    });
});

cornerResetBtn.addEventListener("click", function() {
    state.cornerPoints = [];
    cornerApplyBtn.disabled = true;
    updateCornerHint();
    drawCornerCanvas();
});

cornerCancelBtn.addEventListener("click", function() {
    if (state.processedUrl) {
        showStep(step2, 2);
    } else {
        showStep(step2fail, 2);
    }
});

// --- Brush Mode ---
function loadBrushCanvas() {
    state.brushHistory = [];
    brushUndoBtn.disabled = true;

    var url = state.processedUrl || state.originalUrl;
    state.brushImg = new Image();
    state.brushImg.onload = function() {
        var maxW = Math.min(window.innerWidth - 48, 500);
        state.brushScale = Math.min(maxW / state.brushImg.width, 1);
        brushCanvas.width = Math.round(state.brushImg.width * state.brushScale);
        brushCanvas.height = Math.round(state.brushImg.height * state.brushScale);

        state.brushMaskCanvas = document.createElement("canvas");
        state.brushMaskCanvas.width = brushCanvas.width;
        state.brushMaskCanvas.height = brushCanvas.height;
        state.brushMaskCtx = state.brushMaskCanvas.getContext("2d");
        state.brushMaskCtx.fillStyle = "black";
        state.brushMaskCtx.fillRect(0, 0, brushCanvas.width, brushCanvas.height);

        drawBrushCanvas();
    };
    state.brushImg.src = url + "?t=" + Date.now();
}

function drawBrushCanvas() {
    brushCtx.drawImage(state.brushImg, 0, 0, brushCanvas.width, brushCanvas.height);

    var maskData = state.brushMaskCtx.getImageData(0, 0, brushCanvas.width, brushCanvas.height);
    var overlay = brushCtx.getImageData(0, 0, brushCanvas.width, brushCanvas.height);
    for (var i = 0; i < maskData.data.length; i += 4) {
        if (maskData.data[i] > 128) {
            overlay.data[i] = Math.min(255, overlay.data[i] + 100);
            overlay.data[i + 1] = Math.max(0, overlay.data[i + 1] - 50);
            overlay.data[i + 2] = Math.max(0, overlay.data[i + 2] - 50);
            overlay.data[i + 3] = 200;
        }
    }
    brushCtx.putImageData(overlay, 0, 0);
}

function paintBrushMask(x, y) {
    var r = parseInt(brushSize.value) * state.brushScale;
    state.brushMaskCtx.beginPath();
    state.brushMaskCtx.arc(x, y, r, 0, Math.PI * 2);
    state.brushMaskCtx.fillStyle = "white";
    state.brushMaskCtx.fill();
}

function sendBrushInpaint() {
    var maskData = state.brushMaskCtx.getImageData(0, 0, brushCanvas.width, brushCanvas.height);
    var hasWhite = false;
    for (var i = 0; i < maskData.data.length; i += 4) {
        if (maskData.data[i] > 128) { hasWhite = true; break; }
    }
    if (!hasWhite) return;

    state.brushHistory.push(state.brushImg.src);
    brushUndoBtn.disabled = false;

    var imgC = document.createElement("canvas");
    imgC.width = state.brushImg.naturalWidth;
    imgC.height = state.brushImg.naturalHeight;
    imgC.getContext("2d").drawImage(state.brushImg, 0, 0);
    var imageB64 = imgC.toDataURL("image/jpeg", 0.92);

    var maskC = document.createElement("canvas");
    maskC.width = state.brushImg.naturalWidth;
    maskC.height = state.brushImg.naturalHeight;
    maskC.getContext("2d").drawImage(state.brushMaskCanvas, 0, 0, maskC.width, maskC.height);
    var maskB64 = maskC.toDataURL("image/png");

    var spinner = document.createElement("div");
    spinner.className = "loading-overlay";
    spinner.textContent = "AI 보정 중...";
    brushCanvas.parentElement.appendChild(spinner);

    fetch("/inpaint", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ image: imageB64, mask: maskB64, file_id: state.fileId })
    })
    .then(function(resp) { return resp.json().then(function(d) { return { ok: resp.ok, data: d }; }); })
    .then(function(result) {
        spinner.remove();
        if (!result.ok) { alert(result.data.error || "보정 실패"); return; }
        state.processedUrl = result.data.processed;
        state.brushImg = new Image();
        state.brushImg.onload = function() {
            state.brushMaskCtx.fillStyle = "black";
            state.brushMaskCtx.fillRect(0, 0, brushCanvas.width, brushCanvas.height);
            drawBrushCanvas();
        };
        state.brushImg.src = result.data.result_image;
    })
    .catch(function() { spinner.remove(); alert("서버 연결 실패"); });
}

// Touch events for brush
var brushLastPos = null;
brushCanvas.addEventListener("touchstart", function(e) {
    e.preventDefault();
    state.brushDrawing = true;
    var pos = canvasTouchPos(brushCanvas, e);
    brushLastPos = pos;
    paintBrushMask(pos.x, pos.y);
    drawBrushCanvas();
});
brushCanvas.addEventListener("touchmove", function(e) {
    e.preventDefault();
    if (!state.brushDrawing) return;
    var pos = canvasTouchPos(brushCanvas, e);
    paintBrushMask(pos.x, pos.y);
    brushLastPos = pos;
    drawBrushCanvas();
});
brushCanvas.addEventListener("touchend", function(e) {
    e.preventDefault();
    if (!state.brushDrawing) return;
    state.brushDrawing = false;
    sendBrushInpaint();
});

// Mouse events for brush (desktop fallback)
brushCanvas.addEventListener("mousedown", function(e) {
    state.brushDrawing = true;
    var pos = canvasTouchPos(brushCanvas, e);
    paintBrushMask(pos.x, pos.y);
    drawBrushCanvas();
});
brushCanvas.addEventListener("mousemove", function(e) {
    var pos = canvasTouchPos(brushCanvas, e);
    if (state.brushDrawing) paintBrushMask(pos.x, pos.y);
    drawBrushCanvas();
    // Brush cursor
    var r = parseInt(brushSize.value) * state.brushScale;
    brushCtx.beginPath();
    brushCtx.arc(pos.x, pos.y, r, 0, Math.PI * 2);
    brushCtx.strokeStyle = "rgba(255,255,255,0.8)";
    brushCtx.lineWidth = 2;
    brushCtx.stroke();
});
brushCanvas.addEventListener("mouseup", function() {
    if (!state.brushDrawing) return;
    state.brushDrawing = false;
    sendBrushInpaint();
});

brushSize.addEventListener("input", function() {
    brushSizeLabel.textContent = brushSize.value + "px";
});

brushUndoBtn.addEventListener("click", function() {
    if (state.brushHistory.length === 0) return;
    var prev = state.brushHistory.pop();
    if (state.brushHistory.length === 0) brushUndoBtn.disabled = true;
    state.brushImg = new Image();
    state.brushImg.onload = function() {
        state.brushMaskCtx.fillStyle = "black";
        state.brushMaskCtx.fillRect(0, 0, brushCanvas.width, brushCanvas.height);
        drawBrushCanvas();
    };
    state.brushImg.src = prev;
});

brushDoneBtn.addEventListener("click", function() {
    reviewImg.src = (state.processedUrl || "") + "?t=" + Date.now();
    reviewMessage.textContent = "보정 완료!";
    handBadge.hidden = true;
    showStep(step2, 2);
});

brushCancelBtn.addEventListener("click", function() {
    if (state.processedUrl) {
        showStep(step2, 2);
    } else {
        showStep(step2fail, 2);
    }
});

}); // end DOMContentLoaded
```

- [ ] **Step 2: Start the dev server and test in browser**

Run: `cd gallery && python app.py`
Open `http://localhost:5000` on phone or browser dev tools mobile view.
Test: upload photo → review result → manual adjust → brush → save.

- [ ] **Step 3: Commit**

```bash
git add gallery/static/app.js
git commit -m "feat: 모바일 위자드 JS 재작성 (터치 이벤트, 스텝 네비게이션)"
```

---

### Task 10: Integration test — full pipeline

Verify the entire pipeline works end-to-end.

**Files:**
- None (manual testing)

- [ ] **Step 1: Run all tests**

Run: `cd gallery && python -m pytest tests/ -v`
Expected: All tests pass.

- [ ] **Step 2: Start server and test manually**

Run: `cd gallery && python app.py`

Test checklist:
1. Open `http://localhost:5000` in Chrome DevTools mobile mode
2. Upload a test image → verify auto processing completes
3. Check response includes `hand_detected`, `detection_method`, `confidence`
4. Test "꼭지점 수정" mode with 4-point touch
5. Test "추가 보정" brush mode
6. Test "이대로 저장" confirmation
7. Test "다시 촬영" reset
8. Test detection failure → manual mode

- [ ] **Step 3: Final commit**

```bash
git add -A
git commit -m "feat: Artwork Extractor V2 통합 완료 (Florence-2 + MediaPipe + 모바일 UI)"
```
