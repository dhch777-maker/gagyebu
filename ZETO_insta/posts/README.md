# ZETO 인스타 그리드 — Posts

스펙: [`../../docs/superpowers/specs/2026-05-17-zeto-instagram-launch-design.md`](../../docs/superpowers/specs/2026-05-17-zeto-instagram-launch-design.md)
플랜: [`../../docs/superpowers/plans/2026-05-17-zeto-instagram-launch.md`](../../docs/superpowers/plans/2026-05-17-zeto-instagram-launch.md)

## 빌드

PowerShell에서:

```powershell
cd ZETO_insta\posts
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
python capture.py
```

결과: `output\01-wordmark.png` … `output\09-cta.png` (모두 1080×1080).

## 확인

`preview.html`을 브라우저로 열면 캡처된 9장이 3×3 그리드로 렌더된다. 캡처 전이면 셀이 "(아직 캡처 안됨)" 플레이스홀더로 표시된다.

## 폰트 폴백 수동 검증

`Druk-Heavy-Trial.otf` 로드 실패 시 폴백(Anton/Impact/Arial Black)으로 자연스럽게 렌더되는지 한 번 확인한다. 약식 절차:

1. `_common/fonts.css`의 `@font-face { font-family: "Druk Heavy"; … }` 블록 전체를 임시로 주석 처리.
2. `python capture.py` 재실행.
3. `output\01-wordmark.png`를 연다. 워드마크가 Impact/Arial Black류로 렌더되며 블루 채움 + 흰 외곽선 + 글로우는 그대로 유지되는지 확인.
4. `@font-face` 주석 해제. 다시 `python capture.py` 실행하여 Druk Heavy 렌더로 복원.

폴백이 부자연스러우면 `tokens.css`의 `--font-druk` 체인에서 우선순위를 조정한다 (별도 커밋).

## 다음 사이클 (carry-over)

본 디렉터리는 9장 단일 그리드와 캐러셀 03·05의 **첫 장**만 포함한다. 후속 작업:

1. **카피 채우기** — 슬로건(02), 철학 본문(03 카루셀 2~5장), 단계 설명(05 카루셀 2~8장), 키워드(04/06), 공간(07), 일자(08), 모집 조건(09).
2. **캐러셀 후속 슬라이드 빌드** — 03은 4–5장, 05는 6–8장 추가.
3. **릴스/스토리/하이라이트 커버** — 본 그리드 톤을 그대로 확장.
