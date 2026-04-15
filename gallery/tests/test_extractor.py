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


def test_find_artwork_contour_returns_four_points():
    from extractor import find_artwork_contour
    img = make_test_image()
    contour = find_artwork_contour(img)
    assert contour is not None, "Should detect a contour"
    assert contour.shape == (4, 2), "Should return 4 corner points, got shape {}".format(contour.shape)


def test_find_artwork_contour_returns_none_for_blank():
    from extractor import find_artwork_contour
    blank = np.full((600, 800, 3), (200, 200, 200), dtype=np.uint8)
    contour = find_artwork_contour(blank)
    assert contour is None, "Should return None when no artwork detected"


def test_correct_perspective_returns_straightened_image():
    from extractor import find_artwork_contour, correct_perspective
    img = make_test_image()
    contour = find_artwork_contour(img)
    assert contour is not None
    corrected = correct_perspective(img, contour)
    assert corrected is not None, "Should return corrected image"
    assert len(corrected.shape) == 3, "Should be a color image"
    h, w = corrected.shape[:2]
    assert 250 < h < 400, "Height {} out of expected range".format(h)
    assert 350 < w < 500, "Width {} out of expected range".format(w)


def test_crop_square_returns_square():
    from extractor import crop_square
    rect_img = np.full((300, 400, 3), (100, 150, 200), dtype=np.uint8)
    square = crop_square(rect_img)
    h, w = square.shape[:2]
    assert h == w, "Should be square, got {}x{}".format(w, h)


def test_crop_square_preserves_content():
    from extractor import crop_square
    tall_img = np.full((500, 300, 3), (50, 100, 150), dtype=np.uint8)
    square = crop_square(tall_img)
    h, w = square.shape[:2]
    assert h == w, "Should be square, got {}x{}".format(w, h)


def test_process_photo_end_to_end():
    from extractor import process_photo
    img = make_test_image()
    result = process_photo(img, output_size=500)
    assert result is not None, "Should return processed image"
    h, w = result.shape[:2]
    assert h == 500 and w == 500, "Should be 500x500, got {}x{}".format(w, h)
