"""Hugging Face Inference API로 FLUX.1-dev 호출. 텍스트 없는 배경 이미지만 생성."""

import os
from pathlib import Path
from dotenv import load_dotenv
from huggingface_hub import InferenceClient
from PIL import Image

_HERE = Path(__file__).parent
_ENV_PATH = _HERE / ".env"


def _load_token() -> str:
    load_dotenv(_ENV_PATH)
    token = os.getenv("HF_TOKEN")
    if not token or not token.startswith("hf_"):
        raise RuntimeError(
            f"HF_TOKEN not found or invalid in {_ENV_PATH}. "
            "huggingface.co/settings/tokens 에서 발급해 .env 에 HF_TOKEN=hf_xxx 형식으로 저장."
        )
    return token


def generate_background(
    prompt: str,
    negative_prompt: str,
    seed: int,
    width: int = 1024,
    height: int = 1024,
    guidance_scale: float = 3.5,
    num_inference_steps: int = 28,
    model: str = "black-forest-labs/FLUX.1-dev",
) -> Image.Image:
    """FLUX 호출 → PIL Image 반환. 1024² 출력."""
    token = _load_token()
    client = InferenceClient(model=model, token=token)
    image = client.text_to_image(
        prompt=prompt,
        negative_prompt=negative_prompt,
        width=width,
        height=height,
        guidance_scale=guidance_scale,
        num_inference_steps=num_inference_steps,
        seed=seed,
    )
    return image
