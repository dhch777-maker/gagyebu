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
