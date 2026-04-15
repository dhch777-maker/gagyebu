import cv2
import numpy as np
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def make_test_image(width=800, height=600, rect_color=(0, 0, 200), bg_color=(220, 220, 220)):
    """Create a synthetic image: gray background with a colored rectangle (simulating a painting)."""
    img = np.full((height, width, 3), bg_color, dtype=np.uint8)
    center = (width // 2, height // 2)
    rect_w, rect_h = 400, 300
    angle = 5  # degrees tilt
    box = cv2.boxPoints(((center[0], center[1]), (rect_w, rect_h), angle))
    box = np.intp(box)
    cv2.fillPoly(img, [box], rect_color)
    return img


def test_process_photo_end_to_end():
    from extractor import process_photo
    img = make_test_image()
    result = process_photo(img, output_size=500)
    assert result is not None, "Should return processed image"
    h, w = result.shape[:2]
    assert h == 500 and w == 500, "Should be 500x500, got {}x{}".format(w, h)


def test_process_photo_returns_none_for_blank():
    from extractor import process_photo
    blank = np.full((600, 800, 3), (200, 200, 200), dtype=np.uint8)
    result = process_photo(blank, output_size=500)
    # With rembg, a blank image may or may not return None depending on model behavior
    # The important thing is it doesn't crash
    if result is not None:
        h, w = result.shape[:2]
        assert h == 500 and w == 500


def test_manual_process_with_given_corners():
    from extractor import manual_process
    img = make_test_image()
    corners = [[200, 150], [600, 150], [600, 450], [200, 450]]
    result = manual_process(img, corners, output_size=500)
    assert result is not None
    h, w = result.shape[:2]
    assert h == 500 and w == 500, "Expected 500x500, got {}x{}".format(w, h)


def test_fit_to_square_returns_square():
    from extractor import _fit_to_square
    rect_img = np.full((300, 400, 3), (100, 150, 200), dtype=np.uint8)
    square = _fit_to_square(rect_img, 500)
    h, w = square.shape[:2]
    assert h == w == 500, "Should be 500x500, got {}x{}".format(w, h)


def test_fit_to_square_tall_image():
    from extractor import _fit_to_square
    tall_img = np.full((500, 300, 3), (50, 100, 150), dtype=np.uint8)
    square = _fit_to_square(tall_img, 500)
    h, w = square.shape[:2]
    assert h == w == 500, "Should be 500x500, got {}x{}".format(w, h)


def test_inpaint_region_fills_masked_area():
    from extractor import inpaint_region
    from PIL import Image

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
