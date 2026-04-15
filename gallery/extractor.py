import cv2
import numpy as np
from typing import Optional
from rembg import remove
from PIL import Image


def process_photo(img: np.ndarray, output_size: int = 1080) -> Optional[np.ndarray]:
    """Full pipeline: rembg background removal -> white bg -> fit into square -> resize.

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

    # Find bounding box of foreground
    coords = np.column_stack(np.where(alpha > 128))
    y_min, x_min = coords.min(axis=0)
    y_max, x_max = coords.max(axis=0)

    cropped = rgb_out[y_min:y_max, x_min:x_max]
    h, w = cropped.shape[:2]

    # Fit into square with white padding (no crop)
    side = max(h, w)
    padding = int(side * 0.05)
    total = side + padding * 2
    square = np.full((total, total, 3), 255, dtype=np.uint8)
    y_off = (total - h) // 2
    x_off = (total - w) // 2
    square[y_off:y_off + h, x_off:x_off + w] = cropped

    # Resize and convert back to BGR
    resized = cv2.resize(square, (output_size, output_size), interpolation=cv2.INTER_LANCZOS4)
    bgr_out = cv2.cvtColor(resized, cv2.COLOR_RGB2BGR)
    return bgr_out
