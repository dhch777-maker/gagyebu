"""ZETO 인스타 그리드 슬롯 HTML을 1080x1080 PNG로 캡처."""
from pathlib import Path
import sys
from playwright.sync_api import sync_playwright

POSTS_DIR = Path(__file__).resolve().parent
OUT_DIR = POSTS_DIR / "output"
CANVAS = 1080

SLOTS = [
    "01-wordmark",
    "02-slogan",
    "03-why",
    "04-keyword-1",
    "05-curriculum-01",
    "06-keyword-2",
    "07-space",
    "08-opening",
    "09-cta",
]


def capture_all() -> list[Path]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(
            viewport={"width": CANVAS, "height": CANVAS},
            device_scale_factor=1,
        )
        page = ctx.new_page()
        for slot in SLOTS:
            src = POSTS_DIR / f"{slot}.html"
            if not src.exists():
                print(f"[skip] {src.name} not found", file=sys.stderr)
                continue
            page.goto(src.as_uri())
            page.wait_for_load_state("networkidle")
            page.evaluate("document.fonts.ready")
            canvas = page.locator(".canvas").first
            out = OUT_DIR / f"{slot}.png"
            canvas.screenshot(path=str(out), omit_background=False)
            print(f"[ok] {out.relative_to(POSTS_DIR)}")
            written.append(out)
        browser.close()
    return written


if __name__ == "__main__":
    capture_all()
