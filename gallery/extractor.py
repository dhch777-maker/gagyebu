import cv2
import numpy as np
from dataclasses import dataclass
from typing import Optional
from PIL import Image

from detector import LegacyDetector, DetectionResult, _order_points
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

_legacy_detector = LegacyDetector()
print("[extractor] Legacy detector loaded")

_hand_detector = MediaPipeHandDetector()
print("[extractor] Hand detector loaded")

_inpainter = LamaInpainter()
print("[extractor] LaMa inpainter loaded")


def process_photo(img: np.ndarray, output_size: int = 1080) -> Optional[ProcessResult]:
    """Full pipeline: detect painting → perspective transform → hand removal → square fit."""
    # Step 1: Detect painting
    detection = _legacy_detector.detect(img)

    # Step 2: If no corners, try fallback crop
    if detection.corners is None:
        fg_alpha, rgba_arr = _legacy_detector.get_foreground_mask(img)
        if np.count_nonzero(fg_alpha > 128) == 0:
            return ProcessResult(success=False, needs_manual=True,
                                 detection_method=detection.method, confidence=0.0)
        fallback = _fallback_crop(rgba_arr, fg_alpha, output_size)
        return ProcessResult(success=True, image=fallback, hand_detected=False,
                             detection_method=detection.method, confidence=0.3)

    # Step 3: Perspective transform
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    warped = _perspective_transform(rgb, detection.corners)

    # Step 4: Auto hand detection + inpainting
    hand_detected = False
    hand_mask = _hand_detector.detect_hand_mask(warped)
    if hand_mask is not None:
        pil_img = Image.fromarray(warped)
        pil_mask = Image.fromarray(hand_mask)
        inpainted = _inpainter.inpaint(pil_img, pil_mask)
        warped = np.array(inpainted)
        hand_detected = True

    # Step 5: Fit to square
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
