"""시안 생성 CLI. 사용법: python generate.py 03"""

import argparse
import sys
from pathlib import Path
from PIL import Image

from prompts import get_spec
from flux_client import generate_background
from compositor import compose


_HERE = Path(__file__).parent
_BG_DIR = _HERE / "_bg"


def _generate(spec_id: str) -> Path:
    spec = get_spec(spec_id)

    # 1. 배경 생성 (캐시 활용 — 같은 seed면 재실행 시 FLUX 안 부르고 디스크 재사용)
    bg_path = _BG_DIR / f"{spec_id}-bg.png"
    if bg_path.exists():
        print(f"[bg] cached → {bg_path}")
        bg = Image.open(bg_path).convert("RGB")
    else:
        print(f"[bg] calling FLUX (seed={spec['seed']})...")
        bg = generate_background(
            prompt=spec["flux_prompt"],
            negative_prompt=spec["flux_negative"],
            seed=spec["seed"],
        )
        bg.save(bg_path)
        print(f"[bg] saved → {bg_path}")

    # 2. 1024 → 1080 업스케일 (Lanczos)
    if bg.size != (1080, 1080):
        bg = bg.resize((1080, 1080), Image.Resampling.LANCZOS)

    # 3. 텍스트 합성
    final = compose(bg, spec["texts"])

    # 4. 저장
    out_path = _HERE / f"{spec_id}-{spec['slug']}.png"
    final.save(out_path)
    print(f"[done] {out_path}")
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(description="ZETO 마케팅 시안 생성")
    parser.add_argument("spec_id", help="시안 ID (예: 03)")
    args = parser.parse_args()

    try:
        _generate(args.spec_id)
    except Exception as e:
        print(f"[error] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
