# ZETO ART v2 — Mobile Promo Page Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a six-frame Next.js 16 mobile promo page at `ZETO_marketing2/web/` that reproduces the v2 marketing concept (black canvas, large cropped colored letters peeking from frame edges) with curtain-mask letter reveals and per-letter heading cascades.

**Architecture:** Vertical snap-scroll deck of six `<section>`s (cover, four letter cards Z/E/T/O, closing). One shared `<CardFrame>` wraps each section, handles 100dvh sizing, snap behavior, and triggers framer-motion variants via `whileInView`. Card content is data-driven from `lib/cards.ts`. Reused config files from existing `zeto/` give a known-good Next.js 16 baseline (sidesteps Next 16 breaking-change risk).

**Tech Stack:** Next.js 16 (App Router) · React 19 · TypeScript · Tailwind CSS v4 · framer-motion 12 · Vitest 4 + @testing-library/react · next/font/google (Bebas Neue, Barlow Condensed, Noto Sans KR)

**Spec:** `docs/superpowers/specs/2026-05-28-zeto-marketing2-design.md`

---

## File Structure

```
ZETO_marketing2/web/
├── app/
│   ├── layout.tsx          # fonts, html lang=ko, body bg-black
│   ├── page.tsx            # mounts <Deck/>
│   └── globals.css         # tailwind + theme tokens + html/body reset
├── components/
│   ├── Deck.tsx            # snap container, IntersectionObserver, keyboard nav
│   ├── CardFrame.tsx       # shared section wrapper + variants trigger
│   ├── CardCover.tsx       # card 0
│   ├── CardLetter.tsx      # cards 1–4 (data-driven for Z/E/T/O)
│   ├── CardClosing.tsx     # card 5
│   └── motion.ts           # all framer-motion Variants
├── lib/
│   └── cards.ts            # card data + types
├── tests/
│   ├── setup.ts            # @testing-library/jest-dom
│   └── cards.test.ts       # data integrity
├── public/                 # (empty for now)
├── package.json
├── tsconfig.json
├── next.config.ts
├── postcss.config.mjs
├── vitest.config.ts
├── eslint.config.mjs
├── next-env.d.ts
└── .gitignore
```

**Boundaries:**
- `lib/cards.ts` — pure data, no React imports. Single source of truth for copy, colors, positions.
- `components/motion.ts` — pure framer-motion `Variants` exports, no JSX.
- `CardFrame` — knows nothing about content, only viewport + trigger mechanics.
- `CardLetter` — one component drives all four Z/E/T/O cards via the data record.
- `Deck` — knows nothing about individual card rendering; switches on `card.kind`.

---

## Task 1: Scaffold project from zeto/ baseline

**Files:**
- Create: `ZETO_marketing2/web/package.json`
- Create: `ZETO_marketing2/web/tsconfig.json`
- Create: `ZETO_marketing2/web/next.config.ts`
- Create: `ZETO_marketing2/web/postcss.config.mjs`
- Create: `ZETO_marketing2/web/vitest.config.ts`
- Create: `ZETO_marketing2/web/eslint.config.mjs`
- Create: `ZETO_marketing2/web/next-env.d.ts`
- Create: `ZETO_marketing2/web/.gitignore`
- Create: `ZETO_marketing2/web/tests/setup.ts`
- Create: `ZETO_marketing2/web/public/.gitkeep`

- [ ] **Step 1: Create directory tree**

```bash
mkdir -p ZETO_marketing2/web/app
mkdir -p ZETO_marketing2/web/components
mkdir -p ZETO_marketing2/web/lib
mkdir -p ZETO_marketing2/web/tests
mkdir -p ZETO_marketing2/web/public
```

- [ ] **Step 2: Copy known-good config files from zeto/**

```bash
cp zeto/tsconfig.json              ZETO_marketing2/web/tsconfig.json
cp zeto/next.config.ts             ZETO_marketing2/web/next.config.ts
cp zeto/postcss.config.mjs         ZETO_marketing2/web/postcss.config.mjs
cp zeto/vitest.config.ts           ZETO_marketing2/web/vitest.config.ts
cp zeto/eslint.config.mjs          ZETO_marketing2/web/eslint.config.mjs
cp zeto/next-env.d.ts              ZETO_marketing2/web/next-env.d.ts
```

- [ ] **Step 3: Write package.json**

`ZETO_marketing2/web/package.json`:

```json
{
  "name": "zeto-marketing2-web",
  "version": "0.1.0",
  "private": true,
  "scripts": {
    "dev": "next dev",
    "build": "next build",
    "start": "next start",
    "lint": "eslint",
    "test": "vitest run",
    "test:watch": "vitest"
  },
  "dependencies": {
    "framer-motion": "^12.40.0",
    "next": "16.2.6",
    "react": "19.2.4",
    "react-dom": "19.2.4"
  },
  "devDependencies": {
    "@tailwindcss/postcss": "^4",
    "@testing-library/jest-dom": "^6.9.1",
    "@testing-library/react": "^16.3.2",
    "@types/node": "^20",
    "@types/react": "^19.2.15",
    "@types/react-dom": "^19.2.3",
    "@vitejs/plugin-react": "^6.0.2",
    "eslint": "^9",
    "eslint-config-next": "16.2.6",
    "jsdom": "^29.1.1",
    "tailwindcss": "^4",
    "typescript": "^5",
    "vitest": "^4.1.7"
  }
}
```

- [ ] **Step 4: Write .gitignore**

`ZETO_marketing2/web/.gitignore`:

```gitignore
node_modules/
.next/
out/
*.tsbuildinfo
.env*.local
.DS_Store
```

- [ ] **Step 5: Write tests/setup.ts**

`ZETO_marketing2/web/tests/setup.ts`:

```ts
import '@testing-library/jest-dom/vitest'
```

- [ ] **Step 6: Write public/.gitkeep**

```bash
touch ZETO_marketing2/web/public/.gitkeep
```

- [ ] **Step 7: Install dependencies**

```bash
cd ZETO_marketing2/web && npm install
```

Expected: install completes without peer-dep errors.

- [ ] **Step 8: Sanity check — TypeScript and lint**

```bash
cd ZETO_marketing2/web && npx tsc --noEmit
```

Expected: PASS (no app code yet, just config). The empty `app/` directory may cause `next build` to fail — that's expected and gets fixed in Task 4.

- [ ] **Step 9: Commit**

```bash
git add ZETO_marketing2/web/package.json ZETO_marketing2/web/package-lock.json \
        ZETO_marketing2/web/tsconfig.json ZETO_marketing2/web/next.config.ts \
        ZETO_marketing2/web/postcss.config.mjs ZETO_marketing2/web/vitest.config.ts \
        ZETO_marketing2/web/eslint.config.mjs ZETO_marketing2/web/next-env.d.ts \
        ZETO_marketing2/web/.gitignore ZETO_marketing2/web/tests/setup.ts \
        ZETO_marketing2/web/public/.gitkeep
git commit -m "feat(zeto-mkt2): scaffold Next.js 16 project from zeto/ baseline"
```

---

## Task 2: Card data model (TDD)

**Files:**
- Create: `ZETO_marketing2/web/tests/cards.test.ts`
- Create: `ZETO_marketing2/web/lib/cards.ts`

- [ ] **Step 1: Write the failing test**

`ZETO_marketing2/web/tests/cards.test.ts`:

```ts
import { describe, it, expect } from 'vitest'
import { cards, LETTER_TOTAL, palette } from '@/lib/cards'

describe('cards data', () => {
  it('has exactly 6 cards: cover, 4 letters, closing', () => {
    expect(cards).toHaveLength(6)
    expect(cards[0].kind).toBe('cover')
    expect(cards[5].kind).toBe('closing')
    const letters = cards.filter((c) => c.kind === 'letter')
    expect(letters).toHaveLength(LETTER_TOTAL)
    expect(LETTER_TOTAL).toBe(4)
  })

  it('letter cards have ids z, e, t, o with indices 1..4', () => {
    const letters = cards.filter((c) => c.kind === 'letter')
    expect(letters.map((c) => c.id)).toEqual(['z', 'e', 't', 'o'])
    expect(letters.map((c) => c.index)).toEqual([1, 2, 3, 4])
  })

  it('every letter card has heading, subhead, 3-line body, color, anchor, curtainDir, textAlign, hint', () => {
    const letters = cards.filter((c) => c.kind === 'letter')
    for (const card of letters) {
      expect(card.heading).toMatch(/^0[1-4] \/ [A-Z]+$/)
      expect(card.subhead).toBeTruthy()
      expect(card.body).toHaveLength(3)
      expect(card.color).toMatch(/^#[0-9A-F]{6}$/i)
      expect(['top-left', 'top-right', 'bottom-left', 'bottom-right']).toContain(card.anchor)
      expect(['up', 'down']).toContain(card.curtainDir)
      expect(['left', 'right']).toContain(card.textAlign)
      expect(card.hint.color).toMatch(/^#[0-9A-F]{6}$/i)
      expect(['bar', 'pill']).toContain(card.hint.shape)
    }
  })

  it('preserves v2 copy exactly for Z (ZERO)', () => {
    const z = cards.find((c) => c.kind === 'letter' && c.id === 'z')
    if (z?.kind !== 'letter') throw new Error('z not letter')
    expect(z.heading).toBe('01 / ZERO')
    expect(z.subhead).toBe('시작')
    expect(z.body).toEqual([
      '모든 창작은 백지에서 시작됩니다',
      '아무것도 없는 그 순간이',
      '가장 많은 가능성을 품고 있습니다',
    ])
  })

  it('preserves v2 copy exactly for E (EXPLORE)', () => {
    const e = cards.find((c) => c.kind === 'letter' && c.id === 'e')
    if (e?.kind !== 'letter') throw new Error('e not letter')
    expect(e.heading).toBe('02 / EXPLORE')
    expect(e.subhead).toBe('탐구')
    expect(e.body).toEqual([
      '방향 없이 걸어보는 것',
      '낯선 길에서 발견하는',
      '예상치 못한 영감들',
    ])
  })

  it('preserves v2 copy exactly for T (THINKING)', () => {
    const t = cards.find((c) => c.kind === 'letter' && c.id === 't')
    if (t?.kind !== 'letter') throw new Error('t not letter')
    expect(t.heading).toBe('03 / THINKING')
    expect(t.subhead).toBe('생각')
    expect(t.body).toEqual([
      '조용히 앉아 천천히 생각하는 시간',
      '머릿속의 안개가 걷히고',
      '하나의 형태가 떠오릅니다',
    ])
  })

  it('preserves v2 copy exactly for O (OUTPUT)', () => {
    const o = cards.find((c) => c.kind === 'letter' && c.id === 'o')
    if (o?.kind !== 'letter') throw new Error('o not letter')
    expect(o.heading).toBe('04 / OUTPUT')
    expect(o.subhead).toBe('작품')
    expect(o.body).toEqual([
      '생각이 손끝을 통해 세상으로 나옵니다',
      '이것이 당신만의 예술',
      '오직 당신이 만들 수 있는 것',
    ])
  })

  it('exposes a palette with canvas, ink, z, e, t, o', () => {
    expect(palette.canvas).toBe('#000000')
    expect(palette.ink).toBe('#FFFFFF')
    expect(palette.z).toMatch(/^#[0-9A-F]{6}$/i)
    expect(palette.e).toMatch(/^#[0-9A-F]{6}$/i)
    expect(palette.t).toMatch(/^#[0-9A-F]{6}$/i)
    expect(palette.o).toMatch(/^#[0-9A-F]{6}$/i)
  })

  it('hint anchors do not collide with the letter anchor on the same card', () => {
    const letters = cards.filter((c) => c.kind === 'letter')
    for (const card of letters) {
      if (card.kind !== 'letter') continue
      expect(card.hint.anchor).not.toBe(card.anchor)
    }
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

```bash
cd ZETO_marketing2/web && npm test -- tests/cards.test.ts
```

Expected: FAIL with "Cannot find module '@/lib/cards'".

- [ ] **Step 3: Write minimal implementation**

`ZETO_marketing2/web/lib/cards.ts`:

```ts
export const LETTER_TOTAL = 4

export const palette = {
  canvas: '#000000',
  ink: '#FFFFFF',
  z: '#F5F5F2',
  e: '#EBD89A',
  t: '#6BA2BD',
  o: '#C97A4A',
} as const

export type LetterId = 'z' | 'e' | 't' | 'o'
export type Anchor = 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right'
export type CurtainDir = 'up' | 'down'
export type HintShape = 'bar' | 'pill'

export type HintCue = {
  color: string
  anchor: Anchor
  shape: HintShape
  widthVw: number
  heightVh: number
}

export type LetterCard = {
  kind: 'letter'
  id: LetterId
  index: 1 | 2 | 3 | 4
  letter: 'Z' | 'E' | 'T' | 'O'
  color: string
  anchor: Anchor
  curtainDir: CurtainDir
  heading: string
  subhead: string
  textAlign: 'left' | 'right'
  body: readonly [string, string, string]
  hint: HintCue
}

export type CoverCard = { kind: 'cover' }
export type ClosingCard = { kind: 'closing' }
export type Card = CoverCard | LetterCard | ClosingCard

export const cards: readonly Card[] = [
  { kind: 'cover' },
  {
    kind: 'letter',
    id: 'z',
    index: 1,
    letter: 'Z',
    color: palette.z,
    anchor: 'top-right',
    curtainDir: 'down',
    heading: '01 / ZERO',
    subhead: '시작',
    textAlign: 'left',
    body: [
      '모든 창작은 백지에서 시작됩니다',
      '아무것도 없는 그 순간이',
      '가장 많은 가능성을 품고 있습니다',
    ],
    hint: { color: palette.e, anchor: 'bottom-left', shape: 'bar', widthVw: 10, heightVh: 28 },
  },
  {
    kind: 'letter',
    id: 'e',
    index: 2,
    letter: 'E',
    color: palette.e,
    anchor: 'top-left',
    curtainDir: 'down',
    heading: '02 / EXPLORE',
    subhead: '탐구',
    textAlign: 'left',
    body: [
      '방향 없이 걸어보는 것',
      '낯선 길에서 발견하는',
      '예상치 못한 영감들',
    ],
    hint: { color: palette.t, anchor: 'bottom-right', shape: 'pill', widthVw: 8, heightVh: 24 },
  },
  {
    kind: 'letter',
    id: 't',
    index: 3,
    letter: 'T',
    color: palette.t,
    anchor: 'bottom-left',
    curtainDir: 'up',
    heading: '03 / THINKING',
    subhead: '생각',
    textAlign: 'right',
    body: [
      '조용히 앉아 천천히 생각하는 시간',
      '머릿속의 안개가 걷히고',
      '하나의 형태가 떠오릅니다',
    ],
    hint: { color: palette.o, anchor: 'top-left', shape: 'bar', widthVw: 6, heightVh: 22 },
  },
  {
    kind: 'letter',
    id: 'o',
    index: 4,
    letter: 'O',
    color: palette.o,
    anchor: 'bottom-right',
    curtainDir: 'up',
    heading: '04 / OUTPUT',
    subhead: '작품',
    textAlign: 'right',
    body: [
      '생각이 손끝을 통해 세상으로 나옵니다',
      '이것이 당신만의 예술',
      '오직 당신이 만들 수 있는 것',
    ],
    hint: { color: palette.ink, anchor: 'top-left', shape: 'bar', widthVw: 5, heightVh: 18 },
  },
  { kind: 'closing' },
] as const
```

- [ ] **Step 4: Run test to verify it passes**

```bash
cd ZETO_marketing2/web && npm test -- tests/cards.test.ts
```

Expected: all 9 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add ZETO_marketing2/web/lib/cards.ts ZETO_marketing2/web/tests/cards.test.ts
git commit -m "feat(zeto-mkt2): add card data model with v2 copy and palette"
```

---

## Task 3: Motion variants

**Files:**
- Create: `ZETO_marketing2/web/components/motion.ts`

- [ ] **Step 1: Write motion variants**

`ZETO_marketing2/web/components/motion.ts`:

```ts
import type { Variants } from 'framer-motion'

const EASE = [0.22, 1, 0.36, 1] as const

// Top-level container — controls when child variants fire.
export const staggerContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0, delayChildren: 0 } },
}

// Curtain reveal for the large cropped letter (Z/E/T/O letter cards).
// Direction is applied at call-site via `custom` prop ('down' or 'up').
export const curtainReveal: Variants = {
  hidden: (dir: 'down' | 'up' = 'down') => ({
    clipPath: dir === 'down' ? 'inset(100% 0 0 0)' : 'inset(0 0 100% 0)',
    opacity: 0.4,
    y: 8,
  }),
  show: {
    clipPath: 'inset(0 0 0 0)',
    opacity: 1,
    y: 0,
    transition: { duration: 0.75, ease: EASE },
  },
}

// Per-letter cascade for headings like "01 / ZERO".
// Used as a child variant — parent provides the stagger.
export const cascadeChar: Variants = {
  hidden: { opacity: 0, y: 12 },
  show: { opacity: 1, y: 0, transition: { duration: 0.38, ease: EASE } },
}

export const cascadeContainer: Variants = {
  hidden: {},
  show: {
    transition: { staggerChildren: 0.035, delayChildren: 0.25 },
  },
}

// Per-line fade-up for 3-line body blocks.
export const bodyLine: Variants = {
  hidden: { opacity: 0, y: 8 },
  show: { opacity: 0.85, y: 0, transition: { duration: 0.45, ease: EASE } },
}

export const bodyContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.09, delayChildren: 0.55 } },
}

// Hint bar — scales from anchor edge. transform-origin is applied via style at call-site.
export const hintBar: Variants = {
  hidden: { scaleY: 0, opacity: 0 },
  show: {
    scaleY: 1,
    opacity: 1,
    transition: { duration: 0.6, ease: EASE, delay: 0.9 },
  },
}

// Cover-specific: vertical drop-in for each of Z/E/T/O in the cover row.
export const coverLetter: Variants = {
  hidden: { opacity: 0, y: -24 },
  show: { opacity: 1, y: 0, transition: { duration: 0.4, ease: EASE } },
}

export const coverLetterContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.12, delayChildren: 0.1 } },
}

// Simple fade-in (subhead, nav labels, etc.)
export const fadeUp: Variants = {
  hidden: { opacity: 0, y: 6 },
  show: { opacity: 1, y: 0, transition: { duration: 0.5, ease: EASE } },
}
```

- [ ] **Step 2: TypeScript check**

```bash
cd ZETO_marketing2/web && npx tsc --noEmit
```

Expected: PASS.

- [ ] **Step 3: Commit**

```bash
git add ZETO_marketing2/web/components/motion.ts
git commit -m "feat(zeto-mkt2): add framer-motion variants (curtain, cascade, hint bar)"
```

---

## Task 4: Global styles, layout, and minimal page (boot the dev server)

**Files:**
- Create: `ZETO_marketing2/web/app/globals.css`
- Create: `ZETO_marketing2/web/app/layout.tsx`
- Create: `ZETO_marketing2/web/app/page.tsx`

- [ ] **Step 1: Write globals.css**

`ZETO_marketing2/web/app/globals.css`:

```css
@import "tailwindcss";

@theme {
  --color-zeto-black: #000000;
  --color-zeto-ink: #FFFFFF;
  --color-zeto-z: #F5F5F2;
  --color-zeto-e: #EBD89A;
  --color-zeto-t: #6BA2BD;
  --color-zeto-o: #C97A4A;

  --font-bebas: var(--font-bebas);
  --font-barlow: var(--font-barlow);
  --font-noto: var(--font-noto);
}

html, body {
  height: 100%;
  margin: 0;
  overflow: hidden;
  background: #000000;
}

body {
  overscroll-behavior: none;
  color: #FFFFFF;
}

@media (prefers-reduced-motion: reduce) {
  * {
    animation-duration: 0.001ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: 0.001ms !important;
  }
}
```

- [ ] **Step 2: Write layout.tsx**

`ZETO_marketing2/web/app/layout.tsx`:

```tsx
import type { Metadata, Viewport } from 'next'
import { Bebas_Neue, Barlow_Condensed, Noto_Sans_KR } from 'next/font/google'
import './globals.css'

const bebas = Bebas_Neue({
  weight: '400',
  subsets: ['latin'],
  variable: '--font-bebas',
  display: 'swap',
})

const barlow = Barlow_Condensed({
  weight: ['600', '700'],
  subsets: ['latin'],
  variable: '--font-barlow',
  display: 'swap',
})

const notoKR = Noto_Sans_KR({
  weight: ['300', '400', '500', '700'],
  subsets: ['latin'],
  variable: '--font-noto',
  display: 'swap',
  preload: false,
})

export const metadata: Metadata = {
  title: 'ZETO ART — Think to Art',
  description: '생각이 예술이 되는 시간, ZETO ART',
}

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  themeColor: '#000000',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko" className={`${bebas.variable} ${barlow.variable} ${notoKR.variable}`}>
      <body className="antialiased bg-black text-white">{children}</body>
    </html>
  )
}
```

- [ ] **Step 3: Write a placeholder page.tsx so dev server can boot**

`ZETO_marketing2/web/app/page.tsx`:

```tsx
export default function Page() {
  return (
    <main className="grid h-[100dvh] place-items-center text-white">
      <span className="font-noto opacity-50">ZETO ART · loading deck…</span>
    </main>
  )
}
```

- [ ] **Step 4: Boot dev server**

```bash
cd ZETO_marketing2/web && npm run dev
```

Expected: `Local: http://localhost:3000` printed. Open in browser → see placeholder text on black background. Kill server (Ctrl+C).

- [ ] **Step 5: Commit**

```bash
git add ZETO_marketing2/web/app/globals.css ZETO_marketing2/web/app/layout.tsx ZETO_marketing2/web/app/page.tsx
git commit -m "feat(zeto-mkt2): wire layout, fonts, and tailwind v4 theme on black canvas"
```

---

## Task 5: CardFrame wrapper

**Files:**
- Create: `ZETO_marketing2/web/components/CardFrame.tsx`

- [ ] **Step 1: Write CardFrame**

`ZETO_marketing2/web/components/CardFrame.tsx`:

```tsx
'use client'

import { motion } from 'framer-motion'
import type { ReactNode } from 'react'
import { staggerContainer } from './motion'

type Props = {
  index: number
  ariaLabel: string
  children: ReactNode
}

export function CardFrame({ index, ariaLabel, children }: Props) {
  return (
    <section
      data-index={index}
      aria-label={ariaLabel}
      className="relative h-[100dvh] w-full snap-start snap-always overflow-hidden bg-black"
    >
      <motion.div
        className="relative h-full w-full"
        variants={staggerContainer}
        initial="hidden"
        whileInView="show"
        viewport={{ once: true, amount: 0.5 }}
      >
        {children}
      </motion.div>
    </section>
  )
}
```

- [ ] **Step 2: TypeScript check**

```bash
cd ZETO_marketing2/web && npx tsc --noEmit
```

Expected: PASS.

- [ ] **Step 3: Commit**

```bash
git add ZETO_marketing2/web/components/CardFrame.tsx
git commit -m "feat(zeto-mkt2): add CardFrame wrapper with play-once viewport trigger"
```

---

## Task 6: CardCover

**Files:**
- Create: `ZETO_marketing2/web/components/CardCover.tsx`

- [ ] **Step 1: Write CardCover**

`ZETO_marketing2/web/components/CardCover.tsx`:

```tsx
'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { coverLetter, coverLetterContainer, fadeUp } from './motion'
import { palette } from '@/lib/cards'

type Props = { indexInDeck: number }

const LETTERS = [
  { ch: 'Z', color: palette.z },
  { ch: 'E', color: palette.e },
  { ch: 'T', color: palette.t },
  { ch: 'O', color: palette.o },
] as const

const NAV = [
  { en: 'Zero', color: palette.ink },
  { en: 'Explore', color: palette.e },
  { en: 'Thinking', color: palette.t },
  { en: 'Output', color: palette.o },
] as const

export function CardCover({ indexInDeck }: Props) {
  return (
    <CardFrame index={indexInDeck} ariaLabel="ZETO — Think to Art">
      {/* Top-right navigation list */}
      <motion.ul
        variants={coverLetterContainer}
        className="absolute right-5 top-8 flex flex-col items-end gap-1 font-noto text-[13px] font-medium tracking-wide"
      >
        {NAV.map((item) => (
          <motion.li
            key={item.en}
            variants={fadeUp}
            style={{ color: item.color }}
          >
            {item.en} ·
          </motion.li>
        ))}
      </motion.ul>

      {/* Big ZETO letters, top-left cropped */}
      <motion.div
        aria-hidden
        variants={coverLetterContainer}
        className="absolute left-[-12vw] top-[16vh] flex font-bebas leading-[0.78]"
        style={{ fontSize: 'clamp(180px, 32vw, 300px)', letterSpacing: '-6px' }}
      >
        {LETTERS.map(({ ch, color }) => (
          <motion.span key={ch} variants={coverLetter} style={{ color }}>
            {ch}
          </motion.span>
        ))}
      </motion.div>

      {/* Bottom-right THINK / TO / ART */}
      <motion.div
        variants={coverLetterContainer}
        className="absolute bottom-10 right-6 flex flex-col items-end font-noto font-bold leading-[0.95] text-white"
        style={{ fontSize: 'clamp(56px, 14vw, 86px)', letterSpacing: '-1px' }}
      >
        {['THINK', 'TO', 'ART'].map((w) => (
          <motion.span key={w} variants={coverLetter}>
            {w}
          </motion.span>
        ))}
      </motion.div>

      <span className="sr-only">ZETO — Think to Art</span>
    </CardFrame>
  )
}
```

- [ ] **Step 2: Mount cover on the page**

`ZETO_marketing2/web/app/page.tsx`:

```tsx
import { CardCover } from '@/components/CardCover'

export default function Page() {
  return (
    <div className="h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory bg-black">
      <CardCover indexInDeck={0} />
    </div>
  )
}
```

- [ ] **Step 3: Visually verify in browser**

```bash
cd ZETO_marketing2/web && npm run dev
```

Open `http://localhost:3000`. Expected: black screen, large `ZETO` row cropped at left bleeding off-screen, top-right four-item color-coded nav list, bottom-right THINK / TO / ART stack in white bold. Letters drop in sequentially on first paint. Kill server.

- [ ] **Step 4: Commit**

```bash
git add ZETO_marketing2/web/components/CardCover.tsx ZETO_marketing2/web/app/page.tsx
git commit -m "feat(zeto-mkt2): add cover card with ZETO drop-in row and Think to Art mark"
```

---

## Task 7: CardLetter (Z/E/T/O, data-driven)

**Files:**
- Create: `ZETO_marketing2/web/components/CardLetter.tsx`

- [ ] **Step 1: Write CardLetter**

`ZETO_marketing2/web/components/CardLetter.tsx`:

```tsx
'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import {
  curtainReveal,
  cascadeChar,
  cascadeContainer,
  bodyLine,
  bodyContainer,
  hintBar,
} from './motion'
import type { LetterCard, Anchor } from '@/lib/cards'

type Props = { card: LetterCard; indexInDeck: number }

function anchorClass(anchor: Anchor): string {
  switch (anchor) {
    case 'top-left':     return 'top-0 left-0'
    case 'top-right':    return 'top-0 right-0'
    case 'bottom-left':  return 'bottom-0 left-0'
    case 'bottom-right': return 'bottom-0 right-0'
  }
}

function letterPositionStyle(anchor: Anchor): React.CSSProperties {
  // The big letter sits at the named corner, half-cropped off the frame edges.
  // We push it off-screen along the anchor edges using negative offsets equal to ~half its box.
  const base: React.CSSProperties = {
    position: 'absolute',
    fontSize: 'clamp(360px, 95vw, 560px)',
    letterSpacing: '-12px',
    lineHeight: 0.78,
    fontWeight: 700,
  }
  switch (anchor) {
    case 'top-left':     return { ...base, top: '-22vh', left: '-18vw' }
    case 'top-right':    return { ...base, top: '-22vh', right: '-18vw' }
    case 'bottom-left':  return { ...base, bottom: '-22vh', left: '-18vw' }
    case 'bottom-right': return { ...base, bottom: '-22vh', right: '-18vw' }
  }
}

function hintStyle(card: LetterCard): React.CSSProperties {
  const { hint } = card
  const radius = hint.shape === 'pill' ? '999px' : '0px'
  const transformOrigin =
    hint.anchor.startsWith('top') ? 'top' : 'bottom'
  return {
    position: 'absolute',
    width: `${hint.widthVw}vw`,
    height: `${hint.heightVh}vh`,
    backgroundColor: hint.color,
    borderRadius: radius,
    transformOrigin,
  }
}

export function CardLetter({ card, indexInDeck }: Props) {
  const ariaLabel = `${card.heading} — ${card.subhead}`

  // Heading and subhead positions follow textAlign and stay clear of the letter corner.
  const textBlockClass =
    card.textAlign === 'left'
      ? 'absolute left-6 top-[44%] max-w-[80%]'
      : 'absolute right-6 top-[44%] max-w-[80%] text-right'

  const headingChars = Array.from(card.heading)

  return (
    <CardFrame index={indexInDeck} ariaLabel={ariaLabel}>
      {/* Big cropped letter with curtain reveal */}
      <motion.span
        aria-hidden
        custom={card.curtainDir}
        variants={curtainReveal}
        className="font-bebas"
        style={{ ...letterPositionStyle(card.anchor), color: card.color }}
      >
        {card.letter}
      </motion.span>

      {/* Hint bar for next letter */}
      <motion.span
        aria-hidden
        variants={hintBar}
        className={anchorClass(card.hint.anchor)}
        style={hintStyle(card)}
      />

      {/* Text block: heading + subhead + body */}
      <div className={textBlockClass}>
        <motion.h2
          variants={cascadeContainer}
          className="font-barlow font-bold text-white"
          style={{ fontSize: 'clamp(28px, 7vw, 36px)', letterSpacing: '1px' }}
        >
          {headingChars.map((c, i) => (
            <motion.span key={`${c}-${i}`} variants={cascadeChar} className="inline-block">
              {c === ' ' ? ' ' : c}
            </motion.span>
          ))}
        </motion.h2>

        <motion.p
          variants={cascadeContainer}
          className="mt-10 font-noto font-medium text-white/90"
          style={{ fontSize: 'clamp(20px, 5vw, 24px)', letterSpacing: '4px' }}
        >
          {Array.from(card.subhead).map((c, i) => (
            <motion.span key={`${c}-${i}`} variants={cascadeChar} className="inline-block">
              {c}
            </motion.span>
          ))}
        </motion.p>

        <motion.div
          variants={bodyContainer}
          className="mt-6 font-noto font-light leading-[1.85] text-white"
          style={{ fontSize: 'clamp(16px, 4.2vw, 18px)' }}
        >
          {card.body.map((line, i) => (
            <motion.div key={i} variants={bodyLine}>
              {line}
            </motion.div>
          ))}
        </motion.div>
      </div>
    </CardFrame>
  )
}
```

- [ ] **Step 2: Mount all 4 letter cards on the page (alongside cover)**

`ZETO_marketing2/web/app/page.tsx`:

```tsx
import { CardCover } from '@/components/CardCover'
import { CardLetter } from '@/components/CardLetter'
import { cards } from '@/lib/cards'

export default function Page() {
  return (
    <div className="h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory bg-black">
      {cards.map((card, i) => {
        if (card.kind === 'cover') return <CardCover key="cover" indexInDeck={i} />
        if (card.kind === 'letter') return <CardLetter key={card.id} card={card} indexInDeck={i} />
        return null
      })}
    </div>
  )
}
```

- [ ] **Step 3: Visually verify all 4 letter cards**

```bash
cd ZETO_marketing2/web && npm run dev
```

Open the page. Scroll down through cover → Z → E → T → O. For each card check:
- Letter cropped at the correct corner (Z top-right, E top-left, T bottom-left, O bottom-right)
- Letter curtain-reveals on entry (top cards drop down, bottom cards rise up)
- Heading "NN / NAME" cascades letter-by-letter
- Subhead and 3-line body appear after a short delay
- Hint bar appears last, at the opposite corner
- Snap-scroll locks each card to viewport

Kill server.

- [ ] **Step 4: Commit**

```bash
git add ZETO_marketing2/web/components/CardLetter.tsx ZETO_marketing2/web/app/page.tsx
git commit -m "feat(zeto-mkt2): add CardLetter with curtain reveal, cascade heading, hint bar"
```

---

## Task 8: CardClosing

**Files:**
- Create: `ZETO_marketing2/web/components/CardClosing.tsx`

- [ ] **Step 1: Write CardClosing**

`ZETO_marketing2/web/components/CardClosing.tsx`:

```tsx
'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { bodyContainer, bodyLine, fadeUp } from './motion'

type Props = { indexInDeck: number }

const LINES = [
  'ZERO 에서 시작해',
  '깊이 보고,',
  '골똘히 생각하고,',
  '자기만의 방식으로 표현하는 우리 아이.',
  '',
  '제토는 그 과정을',
  '가장 가까이에서 함께합니다.',
  '',
  'Think to Art ,',
  '생각이 예술이 되는 시간',
] as const

export function CardClosing({ indexInDeck }: Props) {
  return (
    <CardFrame index={indexInDeck} ariaLabel="ZETO ART — closing">
      <motion.div
        variants={bodyContainer}
        className="absolute left-0 right-0 top-[28%] flex flex-col items-center text-center font-noto font-normal leading-[2] text-white"
        style={{ fontSize: 'clamp(17px, 4.4vw, 19px)' }}
      >
        {LINES.map((line, i) => (
          <motion.div key={i} variants={bodyLine} className="min-h-[1.4em]">
            {line || ' '}
          </motion.div>
        ))}
      </motion.div>

      <motion.div
        variants={fadeUp}
        className="absolute bottom-[18%] left-0 right-0 text-center font-bebas font-bold text-white"
        style={{ fontSize: 'clamp(40px, 10vw, 56px)', letterSpacing: '0px' }}
      >
        ZETO ART.
      </motion.div>
    </CardFrame>
  )
}
```

- [ ] **Step 2: Wire CardClosing into the page**

Edit `ZETO_marketing2/web/app/page.tsx`:

```tsx
import { CardCover } from '@/components/CardCover'
import { CardLetter } from '@/components/CardLetter'
import { CardClosing } from '@/components/CardClosing'
import { cards } from '@/lib/cards'

export default function Page() {
  return (
    <div className="h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory bg-black">
      {cards.map((card, i) => {
        if (card.kind === 'cover') return <CardCover key="cover" indexInDeck={i} />
        if (card.kind === 'closing') return <CardClosing key="closing" indexInDeck={i} />
        return <CardLetter key={card.id} card={card} indexInDeck={i} />
      })}
    </div>
  )
}
```

- [ ] **Step 3: Visually verify**

```bash
cd ZETO_marketing2/web && npm run dev
```

Scroll to the last card. Expected: centered multi-paragraph text fading in line-by-line, then `ZETO ART.` wordmark fading in at the bottom. Kill server.

- [ ] **Step 4: Commit**

```bash
git add ZETO_marketing2/web/components/CardClosing.tsx ZETO_marketing2/web/app/page.tsx
git commit -m "feat(zeto-mkt2): add closing card with manifesto block and ZETO ART wordmark"
```

---

## Task 9: Deck (snap container, IO tracking, keyboard nav)

**Files:**
- Create: `ZETO_marketing2/web/components/Deck.tsx`
- Modify: `ZETO_marketing2/web/app/page.tsx`

- [ ] **Step 1: Write Deck**

`ZETO_marketing2/web/components/Deck.tsx`:

```tsx
'use client'

import { useEffect, useRef, useState } from 'react'
import { MotionConfig } from 'framer-motion'
import { cards } from '@/lib/cards'
import { CardCover } from './CardCover'
import { CardLetter } from './CardLetter'
import { CardClosing } from './CardClosing'

export function Deck() {
  const containerRef = useRef<HTMLDivElement>(null)
  const [activeIndex, setActiveIndex] = useState(0)

  // Track active card via IntersectionObserver (used for keyboard nav).
  useEffect(() => {
    const root = containerRef.current
    if (!root) return
    const sections = root.querySelectorAll<HTMLElement>('section[data-index]')
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting && entry.intersectionRatio >= 0.5) {
            const idx = Number((entry.target as HTMLElement).dataset.index ?? 0)
            setActiveIndex(idx)
          }
        }
      },
      { root, threshold: [0.5] },
    )
    sections.forEach((s) => observer.observe(s))
    return () => observer.disconnect()
  }, [])

  // Keyboard navigation.
  useEffect(() => {
    const root = containerRef.current
    if (!root) return
    const handler = (e: KeyboardEvent) => {
      const isNext = e.key === 'ArrowDown' || e.key === 'PageDown'
      const isPrev = e.key === 'ArrowUp' || e.key === 'PageUp'
      if (!isNext && !isPrev) return
      e.preventDefault()
      const sections = root.querySelectorAll<HTMLElement>('section[data-index]')
      const target = isNext
        ? Math.min(activeIndex + 1, sections.length - 1)
        : Math.max(activeIndex - 1, 0)
      sections[target]?.scrollIntoView({ behavior: 'smooth' })
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [activeIndex])

  return (
    <MotionConfig reducedMotion="user">
      <div
        ref={containerRef}
        tabIndex={0}
        className="h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory bg-black outline-none focus:outline-none"
      >
        {cards.map((card, i) => {
          if (card.kind === 'cover') return <CardCover key="cover" indexInDeck={i} />
          if (card.kind === 'closing') return <CardClosing key="closing" indexInDeck={i} />
          return <CardLetter key={card.id} card={card} indexInDeck={i} />
        })}
      </div>
    </MotionConfig>
  )
}
```

`reducedMotion="user"` tells framer-motion to honor the OS-level `prefers-reduced-motion` setting: transforms collapse, only opacity transitions remain.

- [ ] **Step 2: Simplify page.tsx to mount Deck**

`ZETO_marketing2/web/app/page.tsx`:

```tsx
import { Deck } from '@/components/Deck'

export default function Page() {
  return <Deck />
}
```

- [ ] **Step 3: Run tests + tsc + lint**

```bash
cd ZETO_marketing2/web && npm test && npx tsc --noEmit && npm run lint
```

Expected: tests PASS (9/9), tsc PASS, lint PASS.

- [ ] **Step 4: Manual verification — full deck walkthrough**

```bash
cd ZETO_marketing2/web && npm run dev
```

Open `http://localhost:3000` in mobile-emulation viewport (e.g., 390×844 iPhone 14). Walk through the whole deck twice — once by touch/scroll, once with ArrowDown. Verify:

- [ ] All 6 cards reachable in order: Cover → Z → E → T → O → Closing
- [ ] Each card snaps to viewport (no half-card states)
- [ ] First entry into each letter card triggers: curtain → heading cascade → body → hint bar
- [ ] Scrolling back up does NOT re-trigger animations (play-once via `once: true`)
- [ ] ArrowDown / ArrowUp navigates between cards
- [ ] Background stays pure black across all transitions
- [ ] Letters anchor to correct corners (Z top-right, E top-left, T bottom-left, O bottom-right)
- [ ] Hint bars appear at opposite corner from their host letter
- [ ] No console errors

Kill server.

- [ ] **Step 5: Commit**

```bash
git add ZETO_marketing2/web/components/Deck.tsx ZETO_marketing2/web/app/page.tsx
git commit -m "feat(zeto-mkt2): mount full Deck on root page with snap and keyboard nav"
```

---

## Task 10: Production build smoke test

**Files:** (none new — verification only)

- [ ] **Step 1: Build**

```bash
cd ZETO_marketing2/web && npm run build
```

Expected: build succeeds, no type errors, no warnings about deprecated APIs. If a warning about `next/font/google` preconnect appears with `preload: false` on Noto KR, it is expected and safe.

- [ ] **Step 2: Run production server**

```bash
cd ZETO_marketing2/web && npm start
```

Open `http://localhost:3000`. Same walkthrough as Task 9 Step 4. Production builds disable Fast Refresh — animations should look identical to dev. Kill server.

- [ ] **Step 3: Final commit (only if build produced lockfile changes)**

```bash
git status
# If package-lock.json changed:
git add ZETO_marketing2/web/package-lock.json
git commit -m "chore(zeto-mkt2): lockfile updates after first prod build"
# If nothing to commit, skip this step.
```

---

## Done Criteria

- All 10 tasks committed.
- `npm test` passes in `ZETO_marketing2/web/`.
- `npm run build` succeeds.
- Manual walkthrough in mobile-emulation viewport shows the 6-card v2 deck with curtain reveals, heading cascades, and snap-scroll behavior.
- No console errors in dev or production mode.
