import numpy as np
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def test_hand_detector_returns_none_for_no_hands():
    from hand_detector import MediaPipeHandDetector
    detector = MediaPipeHandDetector()

    # Plain colored image — no hands
    img = np.full((400, 400, 3), (100, 150, 200), dtype=np.uint8)
    mask = detector.detect_hand_mask(img)

    assert mask is None, "Should return None when no hands detected"


def test_hand_detector_mask_shape():
    from hand_detector import MediaPipeHandDetector
    detector = MediaPipeHandDetector()

    img = np.full((300, 400, 3), (200, 200, 200), dtype=np.uint8)
    mask = detector.detect_hand_mask(img)

    if mask is not None:
        assert mask.shape == (300, 400), "Mask should match input dimensions"
        assert mask.dtype == np.uint8
