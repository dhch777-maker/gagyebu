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
    from extractor import process_photo, ProcessResult
    img = make_test_image()
    result = process_photo(img, output_size=500)
    assert isinstance(result, ProcessResult)
    assert result.success
    assert result.image is not None
    h, w = result.image.shape[:2]
    assert h == 500 and w == 500
    assert result.detection_method == "legacy"


def test_process_photo_blank_image():
    from extractor import process_photo, ProcessResult
    blank = np.full((600, 800, 3), (200, 200, 200), dtype=np.uint8)
    result = process_photo(blank, output_size=500)
    assert isinstance(result, ProcessResult)
    # May succeed (rembg finds something) or fail
    if result.success and result.image is not None:
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
