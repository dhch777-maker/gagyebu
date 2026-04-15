import cv2
import numpy as np
from typing import Optional


def find_artwork_contour(img, min_area_ratio=0.05):
    # type: (np.ndarray, float) -> Optional[np.ndarray]
    """Detect the largest quadrilateral in the image (the artwork).

    Args:
        img: BGR image as NumPy array.
        min_area_ratio: Minimum ratio of contour area to image area to be considered.

    Returns:
        4x2 NumPy array of corner points, ordered [top-left, top-right, bottom-right, bottom-left],
        or None if no suitable contour found.
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 50, 150)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    edges = cv2.dilate(edges, kernel, iterations=2)

    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    img_area = img.shape[0] * img.shape[1]
    min_area = img_area * min_area_ratio

    best = None
    best_area = 0

    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < min_area:
            continue
        peri = cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, 0.02 * peri, True)
        if len(approx) == 4 and area > best_area:
            best = approx
            best_area = area

    if best is None:
        return None

    return _order_points(best.reshape(4, 2))


def _order_points(pts):
    # type: (np.ndarray) -> np.ndarray
    """Order points as: top-left, top-right, bottom-right, bottom-left."""
    rect = np.zeros((4, 2), dtype=np.float32)
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    d = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(d)]
    rect[3] = pts[np.argmax(d)]
    return rect


def correct_perspective(img, corners):
    # type: (np.ndarray, np.ndarray) -> np.ndarray
    """Apply perspective transform to straighten the artwork."""
    tl, tr, br, bl = corners

    width_top = np.linalg.norm(tr - tl)
    width_bottom = np.linalg.norm(br - bl)
    max_width = int(max(width_top, width_bottom))

    height_left = np.linalg.norm(bl - tl)
    height_right = np.linalg.norm(br - tr)
    max_height = int(max(height_left, height_right))

    dst = np.array([
        [0, 0],
        [max_width - 1, 0],
        [max_width - 1, max_height - 1],
        [0, max_height - 1],
    ], dtype=np.float32)

    matrix = cv2.getPerspectiveTransform(corners.astype(np.float32), dst)
    warped = cv2.warpPerspective(img, matrix, (max_width, max_height))
    return warped


def crop_square(img):
    # type: (np.ndarray) -> np.ndarray
    """Crop the image to a centered square (no stretch, no distortion)."""
    h, w = img.shape[:2]
    side = min(h, w)
    y_offset = (h - side) // 2
    x_offset = (w - side) // 2
    return img[y_offset:y_offset + side, x_offset:x_offset + side]


def process_photo(img, output_size=1080):
    # type: (np.ndarray, int) -> Optional[np.ndarray]
    """Full pipeline: detect artwork -> correct perspective -> square crop -> resize."""
    contour = find_artwork_contour(img)
    if contour is None:
        return None

    corrected = correct_perspective(img, contour)
    square = crop_square(corrected)
    resized = cv2.resize(square, (output_size, output_size), interpolation=cv2.INTER_LANCZOS4)
    return resized
