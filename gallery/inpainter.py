from abc import ABC, abstractmethod
from PIL import Image
import torch
from simple_lama_inpainting import SimpleLama


class Inpainter(ABC):
    @abstractmethod
    def inpaint(self, image: Image.Image, mask: Image.Image) -> Image.Image:
        """Inpaint masked region. White=area to inpaint.
        Args:
            image: RGB PIL Image.
            mask: Grayscale PIL Image (255=inpaint, 0=keep).
        Returns:
            Inpainted RGB PIL Image.
        """
        pass


class LamaInpainter(Inpainter):
    def __init__(self):
        device = torch.device("cpu")
        self._model = object.__new__(SimpleLama)
        from simple_lama_inpainting.models.model import download_model, LAMA_MODEL_URL
        import os
        model_path = os.environ.get("LAMA_MODEL") or download_model(LAMA_MODEL_URL)
        self._model.model = torch.jit.load(model_path, map_location=device)
        self._model.model.eval()
        self._model.model.to(device)
        self._model.device = device

    def inpaint(self, image: Image.Image, mask: Image.Image) -> Image.Image:
        result = self._model(image, mask)
        return result.convert("RGB")
