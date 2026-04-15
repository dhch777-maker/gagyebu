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
