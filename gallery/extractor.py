import cv2
import numpy as np
from typing import Optional
from rembg import remove, new_session
from PIL import Image
import torch
from simple_lama_inpainting import SimpleLama

# Pre-load AI model (loaded once at startup)
_bg_session = new_session("u2net")

# Pre-load LaMa inpainting model (loaded once at startup, forced CPU)
def _load_lama():
    device = torch.device("cpu")
    lama = object.__new__(SimpleLama)
    from simple_lama_inpainting.models.model import download_model, LAMA_MODEL_URL
    import os
    model_path = os.environ.get("LAMA_MODEL") or download_model(LAMA_MODEL_URL)
    lama.model = torch.jit.load(model_path, map_location=device)
    lama.model.eval()
    lama.model.to(device)
    lama.device = device
    return lama

_lama_model = _load_lama()


def process_photo(img: np.ndarray, output_size: int = 1080) -> Optional[np.ndarray]:
    """Extract painting: rembg background removal -> find rectangle -> perspective transform.

    1. rembg removes wall/floor background
    2. Edge detection finds the painting's 4 corners
    3. Perspective transform corrects tilt -> front-facing view
    4. Fit into square with white padding
    """
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(rgb)

    # Step 1: Remove background
    result = remove(pil_img, session=_bg_session)
    arr = np.array(result)
    fg_alpha = arr[:, :, 3]

    if np.count_nonzero(fg_alpha > 128) == 0:
        return None

    # Step 2: Find painting rectangle (4 corners)
    corners = _find_painting_corners(rgb, fg_alpha)

    if corners is not None:
        # Step 3: Perspective transform from ORIGINAL image
        warped = _perspective_transform(rgb, corners)
        return _fit_to_square(warped, output_size)
    else:
        # Fallback: bounding box of foreground
        return _fallback_crop(arr, fg_alpha, output_size)


def _find_painting_corners(rgb: np.ndarray, fg_alpha: np.ndarray) -> Optional[np.ndarray]:
    """Find 4 corners of the painting using edge detection + contour analysis."""
    h, w = rgb.shape[:2]
    _, binary = cv2.threshold(fg_alpha, 128, 255, cv2.THRESH_BINARY)

    # --- Method 1: Edge detection on original image, masked to foreground ---
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    blurred = cv2.bilateralFilter(gray, 9, 75, 75)

    for canny_lo, canny_hi in [(30, 100), (50, 150), (20, 80)]:
        edges = cv2.Canny(blurred, canny_lo, canny_hi)

        # Mask to foreground area (slightly dilated to include painting edges)
        dk = np.ones((11, 11), np.uint8)
        search = cv2.dilate(binary, dk)
        edges[search == 0] = 0

        # Connect nearby edge fragments
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

    # --- Method 2: Foreground mask contour + convex hull ---
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


def _order_points(pts: np.ndarray) -> np.ndarray:
    """Order points: top-left, top-right, bottom-right, bottom-left."""
    rect = np.zeros((4, 2), dtype=np.float32)
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]   # top-left: smallest x+y
    rect[2] = pts[np.argmax(s)]   # bottom-right: largest x+y
    d = np.diff(pts, axis=1).ravel()
    rect[1] = pts[np.argmin(d)]   # top-right: smallest y-x
    rect[3] = pts[np.argmax(d)]   # bottom-left: largest y-x
    return rect


def _perspective_transform(rgb: np.ndarray, corners: np.ndarray) -> np.ndarray:
    """Warp painting to front-facing view using 4 corner points."""
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
    """Fit image into square with white padding, resize, convert to BGR."""
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


def manual_process(img: np.ndarray, points: list, output_size: int = 1080) -> np.ndarray:
    """Process with manually specified 4 corner points.

    Args:
        img: BGR image (from cv2.imread).
        points: List of 4 [x, y] points in original image coordinates.
        output_size: Final square dimension.
    """
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    corners = _order_points(np.array(points, dtype=np.float32))
    warped = _perspective_transform(rgb, corners)
    return _fit_to_square(warped, output_size)


def _fallback_crop(rgba_arr: np.ndarray, fg_alpha: np.ndarray, output_size: int) -> np.ndarray:
    """Fallback: bounding box of foreground with white background."""
    rgb_out = rgba_arr[:, :, :3].copy()
    rgb_out[fg_alpha < 128] = [255, 255, 255]
    coords = np.column_stack(np.where(fg_alpha > 128))
    y_min, x_min = coords.min(axis=0)
    y_max, x_max = coords.max(axis=0)
    cropped = rgb_out[y_min:y_max, x_min:x_max]
    return _fit_to_square(cropped, output_size)


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
