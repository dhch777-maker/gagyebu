import cv2
import numpy as np
from typing import Optional
from rembg import remove, new_session
from PIL import Image

# Pre-load AI models (loaded once at startup)
_bg_session = new_session("u2net")
_human_session = new_session("u2net_human_seg")


def process_photo(img: np.ndarray, output_size: int = 1080) -> Optional[np.ndarray]:
    """Extract artwork from photo: remove background + human body parts.

    Uses two AI models:
    1. u2net: removes wall/floor background
    2. u2net_human_seg: detects human body (hands, face, feet)

    Painting = foreground - human body
    """
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(rgb)

    # Step 1: Remove background (wall, floor, etc.)
    result_fg = remove(pil_img, session=_bg_session)
    fg_arr = np.array(result_fg)
    fg_alpha = fg_arr[:, :, 3]

    if np.count_nonzero(fg_alpha > 128) == 0:
        return None

    # Step 2: Detect human body (hands, face, feet, legs)
    result_human = remove(pil_img, session=_human_session)
    human_alpha = np.array(result_human)[:, :, 3]

    # Step 3: Painting = foreground AND NOT human
    painting_mask = np.zeros_like(fg_alpha, dtype=np.uint8)
    painting_mask[(fg_alpha > 128) & (human_alpha < 128)] = 255

    # If subtraction removed too much (>80%), fall back to full foreground
    fg_count = np.count_nonzero(fg_alpha > 128)
    paint_count = np.count_nonzero(painting_mask > 0)
    if paint_count < fg_count * 0.2:
        painting_mask = np.where(fg_alpha > 128, 255, 0).astype(np.uint8)

    # Close small gaps at painting edges (where hands overlapped)
    h, w = painting_mask.shape
    k = max(h, w) // 60
    if k >= 3:
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k, k))
        painting_mask = cv2.morphologyEx(painting_mask, cv2.MORPH_CLOSE, kernel)

    # Set non-painting areas to white
    rgb_out = rgb.copy()
    rgb_out[painting_mask == 0] = [255, 255, 255]

    # Find bounding box of painting area
    coords = np.column_stack(np.where(painting_mask > 0))
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
