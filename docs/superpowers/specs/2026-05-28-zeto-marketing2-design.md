# ZETO ART v2 — Mobile Promotional Page

**Date:** 2026-05-28
**Status:** Approved (pending spec-file review)
**Location:** `ZETO_marketing2/web/`
**Source assets:** `ZETO_marketing2/KakaoTalk_20260528_223450846*.png` (6 frames)

## Overview

A six-frame mobile promotional page for ZETO ART, reproducing the v2 marketing concept (black canvas, large colored letters cropped at frame edges) as a vertically scrolling Next.js page with curtain-style letter reveals and per-letter heading cascades.

This is a **new project** built alongside (not replacing) the existing `zeto/` v1 codebase, which targets a different visual direction (full-color backgrounds, centered letters).

## Tech Stack

- **Next.js 16** (App Router) + **React 19** + **TypeScript**
- **Tailwind CSS v4** (PostCSS plugin)
- **framer-motion** — curtain mask, cascade, fade-up variants
- **Fonts**
  - `Bebas Neue` — large letters and numeric labels
  - `Barlow` — chip/label tracking text
  - `Pretendard` (preferred) or `Noto Sans KR` — Korean body
- **Test runner**: Vitest + @testing-library/react (data integrity only)

> Per `zeto/AGENTS.md`, Next.js 16 has breaking changes from prior versions. Before writing code, consult `node_modules/next/dist/docs/` for the current App Router / Server Component conventions. Reuse the same pinned `next` version as `zeto/` to avoid version drift.

## Card Specifications

All cards: 100dvh × 100vw, black canvas (`#000000`).

### Card 0 — Cover

**Visual** (from `KakaoTalk_20260528_223450846.png`):
- Left-cropped row of four colored letters `Z E T O` filling top-left ~60% of frame
  - Z white `#F5F5F2`, E yellow `#EBD89A`, T blue `#6BA2BD`, O orange `#C97A4A`
  - The Z is cropped — its left half bleeds off the frame
- Top-right vertical nav list: `Zero ·` `Explore ·` `Thinking ·` `Output ·`
  - Each label colored to match its letter (Zero=white, Explore=yellow, Thinking=blue, Output=orange)
- Bottom-right stack (right-aligned, bold white): `THINK` / `TO` / `ART`

**Animation order on enter:** (see Animation System §E for letter mechanics — drop-in, not curtain)
1. Letters drop in left-to-right (Z → E → T → O), 120 ms between each
2. Nav list fades in from right, line-by-line, 50 ms stagger
3. THINK / TO / ART cascades per-letter

### Card 1 — 01 / ZERO

**Visual** (from `_01.png`):
- Top-right: large white `Z`, cropped — only the bottom-left quadrant visible
- Top-left: heading `01 / ZERO` (white, bold, Barlow)
- Mid-left (below heading): subhead `시작` (white, normal weight)
- Below subhead: 3-line body
  ```
  모든 창작은 백지에서 시작됩니다
  아무것도 없는 그 순간이
  가장 많은 가능성을 품고 있습니다
  ```
- Bottom-left edge: tall narrow **yellow** bar (`#EBD89A`) — hint of next letter (E)

### Card 2 — 02 / EXPLORE

**Visual** (from `_02.png`):
- Top-left: large yellow `E`, cropped — only top-left quadrant visible
- Mid-left (below E): heading `02 / EXPLORE`
- Lower-left: subhead `탐구` + 3-line body
  ```
  방향 없이 걸어보는 것
  낯선 길에서 발견하는
  예상치 못한 영감들
  ```
- Bottom-right edge: tall **blue** pill (`#6BA2BD`) — hint of next letter (T)

### Card 3 — 03 / THINKING

**Visual** (from `_03.png`):
- Bottom-left: large blue `T`, cropped — only top-right quadrant visible
- Top-left edge: thin **orange** bar (`#C97A4A`) — hint of next letter (O)
- Mid-right: heading `03 / THINKING`
- Lower-right: subhead `생각` + 3-line body (right-aligned)
  ```
  조용히 앉아 천천히 생각하는 시간
  머릿속의 안개가 걷히고
  하나의 형태가 떠오릅니다
  ```

### Card 4 — 04 / OUTPUT

**Visual** (from `_04.png`):
- Bottom-right: large orange `O`, cropped — only top-left quadrant visible
- Top-left edge: short **white** bar — hint of return-to-brand
- Top-right: heading `04 / OUTPUT`
- Mid-right: subhead `작품` + 3-line body (right-aligned)
  ```
  생각이 손끝을 통해 세상으로 나옵니다
  이것이 당신만의 예술
  오직 당신이 만들 수 있는 것
  ```

### Card 5 — Closing

**Visual** (from `_05.png`):
- Pure black, center-aligned text block at vertical 40%:
  ```
  ZERO 에서 시작해
  깊이 보고,
  골똘히 생각하고,
  자기만의 방식으로 표현하는 우리 아이.

  제토는 그 과정을
  가장 가까이에서 함께합니다.

  Think to Art ,
  생각이 예술이 되는 시간
  ```
- Bottom-center (~75% down): bold `ZETO ART.` wordmark

## Animation System

Implemented as framer-motion `Variants` in `components/motion.ts`. Triggered when a card enters the viewport at ≥50% visibility (IntersectionObserver). **Play-once** per card to prevent flicker on scroll-back.

### A. Curtain Reveal — large letters (Z/E/T/O)

Direction matches the side the letter is cropped from:

| Letter | Card | Crop side | Curtain direction |
|---|---|---|---|
| Z | 1 | top-right | down |
| E | 2 | top-left | down |
| T | 3 | bottom-left | up |
| O | 4 | bottom-right | up |

- `clip-path: inset(100% 0 0 0)` → `inset(0 0 0 0)` (or inverse for bottom)
- Duration: **750 ms**, ease `cubic-bezier(0.22, 1, 0.36, 1)`
- Concurrent: `transform: translate(0, 8px)` → `0`, opacity `0.4 → 1`

### B. Heading Cascade — `01 / ZERO` style

- Split per-letter (preserve spaces and slash)
- Each char: `y: 12 → 0`, opacity `0 → 1`
- Stagger: **35 ms**
- Start delay: **250 ms** (after curtain is ~50% open)
- Duration per char: **380 ms**

### C. Body Fade-up — 3-line block

- Split per-line (not per-char)
- Each line: `y: 8 → 0`, opacity `0 → 0.85`
- Stagger: **90 ms**
- Start delay: **550 ms**

### D. Hint Bar — next-letter cue

- `transform-origin` matches anchor edge (e.g., `top` for top-anchored bars)
- `scaleY: 0 → 1`, duration **600 ms**
- Start delay: **900 ms** (last to appear, signals "next")

### E. Cover-specific letter sequence

The cover does not use curtain (letters are the full composition, not cropped reveals). Instead each letter does a vertical drop-in (`y: -24 → 0`, opacity `0 → 1`, 400 ms, stagger 120 ms).

## Color Palette

```ts
export const palette = {
  canvas: '#000000',
  ink:    '#FFFFFF',
  z:      '#F5F5F2', // off-white
  e:      '#EBD89A', // butter yellow
  t:      '#6BA2BD', // dusty blue
  o:      '#C97A4A', // burnt orange
} as const
```

Body text uses `ink` at `opacity: 0.85`; nav labels match their letter color.

## Typography

| Role | Family | Weight | Size (mobile) |
|---|---|---|---|
| Large letter | Bebas Neue | 700 | `clamp(220px, 65vw, 360px)`, letter-spacing −10 |
| Heading `NN / NAME` | Barlow | 700 | `clamp(28px, 7vw, 36px)`, letter-spacing 1 |
| Subhead (시작 etc.) | Pretendard | 500 | `clamp(20px, 5vw, 24px)`, letter-spacing 4 |
| Body 3-line | Pretendard | 300 | `clamp(16px, 4.2vw, 18px)`, line-height 1.85 |
| Wordmark `ZETO ART.` | Bebas Neue | 700 | `clamp(40px, 10vw, 56px)` |
| Closing body | Pretendard | 400 | `clamp(17px, 4.4vw, 19px)`, line-height 2 |

## Navigation & Behavior

- Container: `h-[100dvh] overflow-y-scroll snap-y snap-mandatory`
- Each card: `h-[100dvh] snap-start`
- Keyboard: ↓/PageDown → next, ↑/PageUp → previous (preventDefault on the document while focused)
- IntersectionObserver tracks active card index for keyboard nav and (optional) progress indicator
- Optional: top-fixed 6-dot progress indicator (small, low opacity); flag `showProgress` in `lib/cards.ts` config — default off

## Directory Structure

```
ZETO_marketing2/web/
├── app/
│   ├── layout.tsx        # font loading, html lang="ko", bg-black
│   ├── page.tsx          # mounts <Deck/>
│   └── globals.css       # tailwind import, font faces
├── components/
│   ├── Deck.tsx          # snap container, IO tracking, keyboard nav
│   ├── CardFrame.tsx     # shared 100dvh wrapper, triggers variants on view
│   ├── CardCover.tsx
│   ├── CardLetter.tsx    # one component for Z/E/T/O, data-driven
│   ├── CardClosing.tsx
│   └── motion.ts         # all framer-motion Variants
├── lib/
│   └── cards.ts          # card data: text, color, letter position, hint bar config
├── tests/
│   └── cards.test.ts     # data integrity (4 letters, palette completeness)
├── public/
│   └── fonts/            # self-hosted Pretendard/Bebas/Barlow if not via next/font
├── package.json
├── tsconfig.json
├── next.config.ts
├── postcss.config.mjs
├── tailwind.config.ts    # if needed for v4
└── vitest.config.ts
```

## Data Model (`lib/cards.ts`)

```ts
type LetterId = 'z' | 'e' | 't' | 'o'
type Anchor = 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right'
type CurtainDir = 'down' | 'up'

type LetterCard = {
  kind: 'letter'
  id: LetterId
  index: 1 | 2 | 3 | 4
  letter: 'Z' | 'E' | 'T' | 'O'
  color: string                    // letter fill
  anchor: Anchor                   // where the letter sits (cropped from this edge)
  curtainDir: CurtainDir
  heading: string                  // e.g. '01 / ZERO'
  subhead: string                  // e.g. '시작'
  textAlign: 'left' | 'right'
  body: readonly [string, string, string]
  hint: {                          // next-letter cue bar
    color: string
    anchor: Anchor
    shape: 'bar' | 'pill'
    sizeVw: { width: number; height: number }
  }
}
```

Cover and Closing are separate `kind: 'cover' | 'closing'` variants with their own component.

## Accessibility

- `aria-label` on each card: `"카드 N: <title>"`
- Large decorative letter: `aria-hidden="true"` (information is duplicated in the heading)
- Respect `prefers-reduced-motion`: skip curtain & cascade, use plain `opacity 0 → 1` over 200 ms
- Focus ring on the scroll container; keyboard nav documented in an off-screen skip-link

## Testing

Minimal scope — visual work is verified by `npm run dev` + manual browser walkthrough.

- `tests/cards.test.ts`: assert `cards` array has exactly 6 entries (1 cover, 4 letters, 1 closing), letter indices 1..4 distinct, all letter colors are valid hex.
- No animation tests (framer-motion runtime is not assertable headlessly without flake).
- Manual verification checklist documented in the implementation plan.

## Out of Scope (YAGNI)

- Internationalization / language toggle
- Dark/light mode toggle (canvas is permanently black)
- Horizontal swipe / Instagram-carousel mode
- Audio / haptics
- Share buttons, social meta beyond basic OG
- Analytics
- CMS integration — content is hard-coded in `lib/cards.ts`
- Image export back to PNG carousel

## Open Questions Resolved

1. **Build location** → new `ZETO_marketing2/web/` (user choice)
2. **Reveal style** → curtain mask + per-letter cascade (user choice)
3. **Navigation** → vertical snap-scroll (user choice)
4. **Font for Korean body** → Pretendard preferred, Noto Sans KR fallback (designer default, not asked but reversible)
5. **Cover letter animation** → drop-in (curtain only fits the cropped letter cards, not the full Z-E-T-O row)

## Next Step

After user approves this written spec, invoke the `superpowers:writing-plans` skill to produce a step-by-step implementation plan.
