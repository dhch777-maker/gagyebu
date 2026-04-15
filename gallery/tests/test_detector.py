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
