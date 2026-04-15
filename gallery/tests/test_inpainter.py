import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PIL import Image, ImageDraw


def test_lama_inpainter_returns_rgb_image():
    from inpainter import LamaInpainter
    inpainter = LamaInpainter()

    img = Image.new("RGB", (200, 200), (200, 100, 100))
    mask = Image.new("L", (200, 200), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse([80, 80, 120, 120], fill=255)

    result = inpainter.inpaint(img, mask)

    assert isinstance(result, Image.Image)
    assert result.size == (200, 200)
    assert result.mode == "RGB"


def test_lama_inpainter_with_empty_mask():
    from inpainter import LamaInpainter
    inpainter = LamaInpainter()

    img = Image.new("RGB", (200, 200), (100, 150, 200))
    mask = Image.new("L", (200, 200), 0)  # all black = nothing to inpaint

    result = inpainter.inpaint(img, mask)

    assert isinstance(result, Image.Image)
    assert result.size == (200, 200)
