import cv2
import numpy as np
from typing import Optional
from rembg import remove
from PIL import Image


def process_photo(img: np.ndarray, output_size: int = 1080) -> Optional[np.ndarray]:
    """Full pipeline: rembg background removal -> remove hands/body -> white bg -> square.

    Args:
        img: BGR image (from cv2.imread).
        output_size: Final square dimension in pixels.

    Returns:
        Square BGR image with white background, or None if processing failed.
    """
    # BGR -> RGB -> PIL
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(rgb)

    # Remove background (returns RGBA)
    result = remove(pil_img)
    arr = np.array(result)

    alpha = arr[:, :, 3]
    if np.count_nonzero(alpha > 128) == 0:
        return None

    # Replace transparent with white
    rgb_out = arr[:, :, :3].copy()
    rgb_out[alpha < 128] = [255, 255, 255]

    # --- Remove hands/body using morphological opening ---
    # Opening (erosion + dilation) removes thin protrusions (fingers, arms, feet)
    # while preserving the large rectangular painting area.
    _, binary = cv2.threshold(alpha, 128, 255, cv2.THRESH_BINARY)

    h, w = binary.shape
    kernel_size = max(h, w) // 20  # ~5% of image dimension
    if kernel_size < 5:
        kernel_size = 5
    if kernel_size % 2 == 0:
        kernel_size += 1
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (kernel_size, kernel_size))
    opened = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)

    # Check we didn't remove too much (keep at least 30% of original foreground)
    if np.count_nonzero(opened) > np.count_nonzero(binary) * 0.3:
        coords = np.column_stack(np.where(opened > 0))
    else:
        # Fallback: use original mask
        coords = np.column_stack(np.where(alpha > 128))

    y_min, x_min = coords.min(axis=0)
    y_max, x_max = coords.max(axis=0)

    cropped = rgb_out[y_min:y_max, x_min:x_max]
    ch, cw = cropped.shape[:2]

    # Fit into square with white padding (no crop)
    side = max(ch, cw)
    padding = int(side * 0.05)
    total = side + padding * 2
    square = np.full((total, total, 3), 255, dtype=np.uint8)
    y_off = (total - ch) // 2
    x_off = (total - cw) // 2
    square[y_off:y_off + ch, x_off:x_off + cw] = cropped

    # Resize and convert back to BGR
    resized = cv2.resize(square, (output_size, output_size), interpolation=cv2.INTER_LANCZOS4)
    bgr_out = cv2.cvtColor(resized, cv2.COLOR_RGB2BGR)
    return bgr_out
