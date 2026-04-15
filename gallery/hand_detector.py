from typing import Optional
import os
import urllib.request
import cv2
import numpy as np
import mediapipe as mp
import mediapipe.tasks as mp_tasks
from mediapipe.tasks.python.vision.hand_landmarker import (
    HandLandmarker,
    HandLandmarkerOptions,
)

_MODEL_URL = (
    "https://storage.googleapis.com/mediapipe-models/"
    "hand_landmarker/hand_landmarker/float16/1/hand_landmarker.task"
)
_DEFAULT_MODEL_PATH = os.path.join(os.path.dirname(__file__), "hand_landmarker.task")


def _ensure_model(model_path: str) -> None:
    """Download the model file if it doesn't exist yet."""
    if not os.path.exists(model_path):
        print(f"[hand_detector] Downloading model to {model_path} ...")
        urllib.request.urlretrieve(_MODEL_URL, model_path)
        print("[hand_detector] Download complete.")


class MediaPipeHandDetector:
    """Detect hands using MediaPipe and generate inpainting mask."""

    def __init__(self, dilation_px: int = 15, model_path: Optional[str] = None):
        self._dilation_px = dilation_px
        if model_path is None:
            model_path = _DEFAULT_MODEL_PATH
        _ensure_model(model_path)

        options = HandLandmarkerOptions(
            base_options=mp_tasks.BaseOptions(model_asset_path=model_path),
            num_hands=4,
            min_hand_detection_confidence=0.3,
        )
        self._detector = HandLandmarker.create_from_options(options)

    def detect_hand_mask(self, image: np.ndarray) -> Optional[np.ndarray]:
        """Detect hands and return a grayscale mask (255=hand, 0=background).

        Args:
            image: RGB numpy array.

        Returns:
            Grayscale mask (same H, W as input) or None if no hands found.
        """
        h, w = image.shape[:2]
        mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=image)
        results = self._detector.detect(mp_image)

        if not results.hand_landmarks:
            return None

        mask = np.zeros((h, w), dtype=np.uint8)

        for hand_landmarks in results.hand_landmarks:
            points = []
            for lm in hand_landmarks:
                px = int(lm.x * w)
                py = int(lm.y * h)
                points.append([px, py])

            points = np.array(points, dtype=np.int32)
            hull = cv2.convexHull(points)
            cv2.fillConvexPoly(mask, hull, 255)

        # Dilate to cover edges around fingers
        if self._dilation_px > 0:
            kernel = cv2.getStructuringElement(
                cv2.MORPH_ELLIPSE,
                (self._dilation_px * 2 + 1, self._dilation_px * 2 + 1),
            )
            mask = cv2.dilate(mask, kernel, iterations=1)

        if np.count_nonzero(mask) == 0:
            return None

        return mask

    def close(self):
        self._detector.close()
