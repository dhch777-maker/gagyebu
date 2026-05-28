# ZETO Mobile Brand Page — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** ZETO 학원 브랜드 카드뉴스 6장을 Next.js + Tailwind + Framer Motion으로 구현해 Vercel에 배포. 풀스크린 스냅 스크롤 + 프리미엄 미니멀 톤 + 스태거드 reveal.

**Architecture:** App Router 기반 단일 페이지. `Deck` 컴포넌트가 스냅 컨테이너로 6 카드를 마운트. `lib/cards.ts` 가 모든 카피/색상의 단일 진실. `CardFrame` 공통 외피가 스냅 섹션 + Framer variants를 캡슐화. 카드별 컴포넌트(`CardCover`, `CardLetter`, `CardClosing`)는 콘텐츠만 책임.

**Tech Stack:** Next.js 15+ (App Router), TypeScript, Tailwind CSS v4, Framer Motion, Vitest + Testing Library.

**Spec:** [`docs/superpowers/specs/2026-05-27-zeto-mobile-design.md`](../specs/2026-05-27-zeto-mobile-design.md)

---

## File structure (target)

```
zeto/
├── app/
│   ├── layout.tsx
│   ├── page.tsx
│   └── globals.css
├── components/
│   ├── Deck.tsx
│   ├── CardFrame.tsx
│   ├── CardCover.tsx
│   ├── CardLetter.tsx
│   ├── CardClosing.tsx
│   └── motion.ts
├── lib/
│   └── cards.ts
├── tests/
│   ├── cards.test.ts
│   └── CardLetter.test.tsx
├── tailwind.config.ts (if Tailwind v3) or @theme in globals.css (v4)
├── vitest.config.ts
├── package.json
├── tsconfig.json
└── README.md
```

---

### Task 1: Scaffold Next.js project

**Files:**
- Create: `c:\Users\dhchd\work\zeto\` (full Next.js project)

- [ ] **Step 1: Run create-next-app non-interactively**

From `c:\Users\dhchd\work\`:

```bash
npx create-next-app@latest zeto --typescript --tailwind --app --no-src-dir --eslint --import-alias "@/*" --use-npm
```

Expected: `zeto/` directory created with package.json, app/layout.tsx, app/page.tsx, app/globals.css, tailwind config, tsconfig, etc.

If the CLI prompts despite flags (Turbopack question, etc.), accept defaults (Yes for Turbopack is fine).

- [ ] **Step 2: Verify project boots**

```bash
cd zeto
npm run dev
```

Expected: dev server starts on http://localhost:3000 with default Next.js welcome page. Stop with Ctrl+C.

- [ ] **Step 3: Add `.gitignore` entries if missing**

Open `zeto/.gitignore` — verify it includes `node_modules/`, `.next/`, `.env*.local`. create-next-app generates these by default; no action needed if already present.

- [ ] **Step 4: Commit initial scaffold**

```bash
cd c:\Users\dhchd\work
git add zeto/
git commit -m "feat(zeto): scaffold Next.js project for ZETO mobile brand page"
```

---

### Task 2: Install Framer Motion

**Files:**
- Modify: `zeto/package.json`

- [ ] **Step 1: Install dependency**

```bash
cd c:\Users\dhchd\work\zeto
npm install framer-motion
```

Expected: `framer-motion` added to `dependencies` in package.json. No errors.

- [ ] **Step 2: Verify install**

```bash
npm list framer-motion
```

Expected: version printed (e.g., `framer-motion@11.x.x` or later).

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/package.json zeto/package-lock.json
git commit -m "feat(zeto): add framer-motion for staggered reveal animations"
```

---

### Task 3: Install Vitest + Testing Library

**Files:**
- Modify: `zeto/package.json`
- Create: `zeto/vitest.config.ts`

- [ ] **Step 1: Install dev dependencies**

```bash
cd c:\Users\dhchd\work\zeto
npm install -D vitest @vitejs/plugin-react @testing-library/react @testing-library/jest-dom jsdom @types/react @types/react-dom
```

Expected: packages added under `devDependencies`. No errors.

- [ ] **Step 2: Create `vitest.config.ts`**

```ts
import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import path from 'path'

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./tests/setup.ts'],
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './'),
    },
  },
})
```

- [ ] **Step 3: Create `tests/setup.ts`**

```ts
import '@testing-library/jest-dom/vitest'
```

- [ ] **Step 4: Add test scripts to `package.json`**

Open `zeto/package.json` and add to `scripts`:

```json
"test": "vitest run",
"test:watch": "vitest"
```

- [ ] **Step 5: Verify Vitest runs (no tests yet)**

```bash
npm run test
```

Expected: "No test files found" — not an error. Vitest is wired correctly.

- [ ] **Step 6: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/package.json zeto/package-lock.json zeto/vitest.config.ts zeto/tests/setup.ts
git commit -m "test(zeto): wire up Vitest + Testing Library"
```

---

### Task 4: Define card data with tests (TDD)

**Files:**
- Create: `zeto/tests/cards.test.ts`
- Create: `zeto/lib/cards.ts`

- [ ] **Step 1: Write failing tests first**

Create `zeto/tests/cards.test.ts`:

```ts
import { describe, it, expect } from 'vitest'
import { cards, LETTER_TOTAL } from '@/lib/cards'

describe('cards data', () => {
  it('has exactly 6 cards: cover, 4 letters, closing', () => {
    expect(cards).toHaveLength(6)
    expect(cards[0].kind).toBe('cover')
    expect(cards[5].kind).toBe('closing')
    const letters = cards.filter((c) => c.kind === 'letter')
    expect(letters).toHaveLength(LETTER_TOTAL)
    expect(LETTER_TOTAL).toBe(4)
  })

  it('letter cards have ids z, e, t, o in order with sequential indices', () => {
    const letters = cards.filter((c) => c.kind === 'letter')
    expect(letters.map((c) => c.id)).toEqual(['z', 'e', 't', 'o'])
    expect(letters.map((c) => c.index)).toEqual([1, 2, 3, 4])
  })

  it('letter cards expose en, ko, desc (3 lines), and palette', () => {
    const letters = cards.filter((c) => c.kind === 'letter')
    for (const card of letters) {
      expect(card.en).toMatch(/^[A-Z][a-z]+$/)
      expect(card.ko).toHaveLength(2)
      expect(card.desc).toHaveLength(3)
      expect(card.palette.screen).toMatch(/^#[0-9A-F]{6}$/i)
      expect(card.palette.letter).toMatch(/^#[0-9A-F]{6}$/i)
      expect(card.palette.bottomBg).toMatch(/^#[0-9A-F]{6}$/i)
      expect(card.palette.bottomFg).toMatch(/^#[0-9A-F]{6}$/i)
      expect(card.palette.chipHex).toBeTruthy()
    }
  })

  it('preserves v3 copy exactly for Z', () => {
    const z = cards.find((c) => c.kind === 'letter' && c.id === 'z')
    expect(z).toBeDefined()
    if (z?.kind !== 'letter') throw new Error('z not letter')
    expect(z.en).toBe('Zero')
    expect(z.ko).toBe('시작')
    expect(z.desc).toEqual([
      '모든 창작은 백지에서 시작됩니다.',
      '아무것도 없는 그 순간이',
      '가장 많은 가능성을 품고 있습니다.',
    ])
  })
})
```

- [ ] **Step 2: Run tests, verify failure**

```bash
cd c:\Users\dhchd\work\zeto
npm run test
```

Expected: FAIL — `Cannot find module '@/lib/cards'` or similar.

- [ ] **Step 3: Implement `lib/cards.ts`**

```ts
export const LETTER_TOTAL = 4

export type LetterId = 'z' | 'e' | 't' | 'o'

export type LetterPalette = {
  screen: string      // full screen bg
  letter: string      // large letter color
  bottomBg: string    // dark block bg
  bottomFg: string    // primary text color in dark block
  chipHex: string     // chip label text
  topInk: string      // top corner (number + bars) ink
}

export type LetterCard = {
  kind: 'letter'
  id: LetterId
  index: number
  letter: 'Z' | 'E' | 'T' | 'O'
  en: string
  ko: string
  desc: readonly [string, string, string]
  palette: LetterPalette
}

export type CoverCard = { kind: 'cover' }
export type ClosingCard = { kind: 'closing' }
export type Card = CoverCard | LetterCard | ClosingCard

const BLACK = '#000000'
const WHITE = '#FFFFFF'

export const cards: readonly Card[] = [
  { kind: 'cover' },
  {
    kind: 'letter',
    id: 'z',
    index: 1,
    letter: 'Z',
    en: 'Zero',
    ko: '시작',
    desc: [
      '모든 창작은 백지에서 시작됩니다.',
      '아무것도 없는 그 순간이',
      '가장 많은 가능성을 품고 있습니다.',
    ],
    palette: {
      screen: WHITE,
      letter: BLACK,
      bottomBg: BLACK,
      bottomFg: WHITE,
      chipHex: '#FFFFFF / #000000',
      topInk: BLACK,
    },
  },
  {
    kind: 'letter',
    id: 'e',
    index: 2,
    letter: 'E',
    en: 'Explore',
    ko: '탐구',
    desc: [
      '방향 없이 걸어보는 것.',
      '낯선 길에서 발견하는',
      '예상치 못한 영감들.',
    ],
    palette: {
      screen: '#F4E1A4',
      letter: '#2A2209',
      bottomBg: BLACK,
      bottomFg: '#F4E1A4',
      chipHex: '#F4E1A4 — Butter Lemon',
      topInk: BLACK,
    },
  },
  {
    kind: 'letter',
    id: 't',
    index: 3,
    letter: 'T',
    en: 'Thinking',
    ko: '생각',
    desc: [
      '조용히 앉아 천천히 생각하는 시간.',
      '머릿속의 안개가 걷히고',
      '하나의 형태가 떠오릅니다.',
    ],
    palette: {
      screen: '#5DA0C0',
      letter: '#1A2A38',
      bottomBg: BLACK,
      bottomFg: '#5DA0C0',
      chipHex: '#5DA0C0 — Dusty Blue',
      topInk: BLACK,
    },
  },
  {
    kind: 'letter',
    id: 'o',
    index: 4,
    letter: 'O',
    en: 'Output',
    ko: '작품',
    desc: [
      '생각이 손끝을 통해 세상으로 나옵니다.',
      '이것이 당신만의 예술,',
      '오직 당신이 만들 수 있는 것.',
    ],
    palette: {
      screen: '#C87649',
      letter: '#1C0D04',
      bottomBg: BLACK,
      bottomFg: '#C87649',
      chipHex: '#C87649 — Burnt Orange',
      topInk: WHITE,
    },
  },
  { kind: 'closing' },
]
```

- [ ] **Step 4: Run tests, verify passing**

```bash
npm run test
```

Expected: PASS — 4 tests passed in `cards.test.ts`.

- [ ] **Step 5: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/lib/cards.ts zeto/tests/cards.test.ts
git commit -m "feat(zeto): add card data model with v3 copy preserved + tests"
```

---

### Task 5: Define shared Framer Motion variants

**Files:**
- Create: `zeto/components/motion.ts`

- [ ] **Step 1: Write the variants file**

```ts
import type { Variants } from 'framer-motion'

export const staggerContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.08, delayChildren: 0.1 } },
}

export const fadeUp: Variants = {
  hidden: { opacity: 0, y: 16 },
  show: {
    opacity: 1,
    y: 0,
    transition: { duration: 0.55, ease: [0.22, 1, 0.36, 1] },
  },
}

export const letterReveal: Variants = {
  hidden: { opacity: 0, y: 40, scale: 0.96 },
  show: {
    opacity: 1,
    y: 0,
    scale: 1,
    transition: { duration: 0.7, ease: [0.22, 1, 0.36, 1] },
  },
}
```

No tests for variants — they're configuration values. Will be exercised via component rendering tests.

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd c:\Users\dhchd\work\zeto
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/components/motion.ts
git commit -m "feat(zeto): define shared Framer Motion variants for staggered reveal"
```

---

### Task 6: Configure fonts in layout

**Files:**
- Modify: `zeto/app/layout.tsx`

- [ ] **Step 1: Read the current layout to preserve existing structure**

```bash
cat zeto/app/layout.tsx
```

Note the metadata/html structure that create-next-app generated.

- [ ] **Step 2: Replace `zeto/app/layout.tsx` with font-aware version**

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
  weight: ['300', '500', '700'],
  subsets: ['latin'],
  variable: '--font-noto',
  display: 'swap',
  preload: false,
})

export const metadata: Metadata = {
  title: 'ZETO — Zero to Art',
  description: '생각이 예술이 되는 시간, 제토아트',
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

Note: `preload: false` on `Noto_Sans_KR` avoids Next.js warnings about Korean subset not being a standard Google Fonts subset.

- [ ] **Step 3: Verify dev server starts without font errors**

```bash
cd c:\Users\dhchd\work\zeto
npm run dev
```

Expected: server starts. http://localhost:3000 loads (still default page). No console errors about fonts. Stop server.

- [ ] **Step 4: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/app/layout.tsx
git commit -m "feat(zeto): load Bebas Neue, Barlow Condensed, Noto Sans KR via next/font"
```

---

### Task 7: Configure Tailwind theme (colors + font families)

**Files:**
- Modify: `zeto/app/globals.css` (Tailwind v4) **or** `zeto/tailwind.config.ts` (v3)

- [ ] **Step 1: Detect Tailwind version**

```bash
cd c:\Users\dhchd\work\zeto
cat zeto/app/globals.css 2>nul | head -5
```

If you see `@import "tailwindcss";` → **Tailwind v4** (use Step 2a).
If you see `@tailwind base; @tailwind components; @tailwind utilities;` → **Tailwind v3** (use Step 2b).

- [ ] **Step 2a (Tailwind v4): Replace `app/globals.css`**

```css
@import "tailwindcss";

@theme {
  --color-zeto-black: #000000;
  --color-zeto-white: #FFFFFF;
  --color-zeto-amber: #F4E1A4;
  --color-zeto-amber-deep: #2A2209;
  --color-zeto-blue: #5DA0C0;
  --color-zeto-blue-deep: #1A2A38;
  --color-zeto-orange: #C87649;
  --color-zeto-orange-deep: #1C0D04;

  --font-bebas: var(--font-bebas);
  --font-barlow: var(--font-barlow);
  --font-noto: var(--font-noto);
}

html, body {
  height: 100%;
  margin: 0;
  overflow: hidden;
}

body {
  overscroll-behavior: none;
}
```

- [ ] **Step 2b (Tailwind v3): Replace `tailwind.config.ts`**

```ts
import type { Config } from 'tailwindcss'

const config: Config = {
  content: [
    './app/**/*.{ts,tsx}',
    './components/**/*.{ts,tsx}',
  ],
  theme: {
    extend: {
      fontFamily: {
        bebas: ['var(--font-bebas)', 'sans-serif'],
        barlow: ['var(--font-barlow)', 'sans-serif'],
        noto: ['var(--font-noto)', 'sans-serif'],
      },
      colors: {
        zeto: {
          black: '#000000',
          white: '#FFFFFF',
          amber: '#F4E1A4',
          'amber-deep': '#2A2209',
          blue: '#5DA0C0',
          'blue-deep': '#1A2A38',
          orange: '#C87649',
          'orange-deep': '#1C0D04',
        },
      },
    },
  },
  plugins: [],
}
export default config
```

Then update `app/globals.css`:

```css
@tailwind base;
@tailwind components;
@tailwind utilities;

html, body {
  height: 100%;
  margin: 0;
  overflow: hidden;
}

body {
  overscroll-behavior: none;
}
```

- [ ] **Step 3: Verify dev server**

```bash
npm run dev
```

Expected: server starts. Default page renders with no console errors.

- [ ] **Step 4: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/app/globals.css
# If Tailwind v3:
# git add zeto/tailwind.config.ts
git commit -m "feat(zeto): configure Tailwind theme with brand colors and font families"
```

---

### Task 8: Build CardFrame (shared snap-section wrapper)

**Files:**
- Create: `zeto/components/CardFrame.tsx`

- [ ] **Step 1: Write the component**

```tsx
'use client'

import { motion } from 'framer-motion'
import type { CSSProperties, ReactNode } from 'react'
import { staggerContainer } from './motion'

type Props = {
  index: number
  ariaLabel: string
  bgClassName?: string
  bgStyle?: CSSProperties
  children: ReactNode
}

export function CardFrame({ index, ariaLabel, bgClassName = '', bgStyle, children }: Props) {
  return (
    <section
      data-index={index}
      aria-label={ariaLabel}
      className={`relative h-screen h-[100dvh] w-full snap-start snap-always overflow-hidden ${bgClassName}`}
      style={bgStyle}
    >
      <motion.div
        className="relative h-full w-full"
        variants={staggerContainer}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
      >
        {children}
      </motion.div>
    </section>
  )
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd c:\Users\dhchd\work\zeto
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/components/CardFrame.tsx
git commit -m "feat(zeto): add CardFrame shared snap-section wrapper with stagger container"
```

---

### Task 9: Build CardLetter component with test

**Files:**
- Create: `zeto/tests/CardLetter.test.tsx`
- Create: `zeto/components/CardLetter.tsx`

- [ ] **Step 1: Write the failing test**

```tsx
import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { CardLetter } from '@/components/CardLetter'
import { cards, type LetterCard } from '@/lib/cards'

function getLetterCard(id: 'z' | 'e' | 't' | 'o'): LetterCard {
  const found = cards.find((c) => c.kind === 'letter' && c.id === id)
  if (!found || found.kind !== 'letter') {
    throw new Error(`fixture missing letter card ${id}`)
  }
  return found
}

const eCard = getLetterCard('e')

describe('CardLetter', () => {
  it('renders aria-label with English and Korean concept', () => {
    render(<CardLetter card={eCard} indexInDeck={2} />)
    expect(screen.getByRole('region', { name: /Explore.*탐구/i })).toBeInTheDocument()
  })

  it('renders the English word, Korean concept, and all 3 desc lines', () => {
    render(<CardLetter card={eCard} indexInDeck={2} />)
    expect(screen.getByText('Explore')).toBeInTheDocument()
    expect(screen.getByText('탐 구')).toBeInTheDocument() // spaced version
    for (const line of eCard.desc) {
      expect(screen.getByText(line)).toBeInTheDocument()
    }
  })

  it('renders the chip hex label', () => {
    render(<CardLetter card={eCard} indexInDeck={2} />)
    expect(screen.getByText(eCard.palette.chipHex)).toBeInTheDocument()
  })

  it('renders index marker "02 / 04"', () => {
    render(<CardLetter card={eCard} indexInDeck={2} />)
    expect(screen.getByText('02 / 04')).toBeInTheDocument()
  })
})
```

- [ ] **Step 2: Run, verify failure**

```bash
cd c:\Users\dhchd\work\zeto
npm run test
```

Expected: FAIL — `Cannot find module '@/components/CardLetter'`.

- [ ] **Step 3: Implement `components/CardLetter.tsx`**

```tsx
'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp, letterReveal } from './motion'
import type { LetterCard } from '@/lib/cards'
import { LETTER_TOTAL } from '@/lib/cards'

type Props = {
  card: LetterCard
  indexInDeck: number
}

function spaced(text: string) {
  return text.split('').join(' ')
}

function pad2(n: number) {
  return n.toString().padStart(2, '0')
}

export function CardLetter({ card, indexInDeck }: Props) {
  const { palette } = card
  const ariaLabel = `${card.en} - ${card.ko}`

  return (
    <CardFrame
      index={indexInDeck}
      ariaLabel={ariaLabel}
      bgStyle={{ backgroundColor: palette.screen }}
    >
      {/* Top-left: number */}
      <motion.div
        variants={fadeUp}
        className="absolute left-5 top-6 font-barlow text-[10px] font-semibold tracking-[3px]"
        style={{ color: palette.topInk, opacity: 0.55 }}
      >
        {pad2(card.index)} / {pad2(LETTER_TOTAL)}
      </motion.div>

      {/* Top-right: progress bars */}
      <motion.div variants={fadeUp} className="absolute right-5 top-6 flex gap-1">
        {Array.from({ length: LETTER_TOTAL }).map((_, i) => {
          const active = i === card.index - 1
          return (
            <span
              key={i}
              className="block h-[2px]"
              style={{
                width: active ? 14 : 8,
                backgroundColor: palette.topInk,
                opacity: active ? 0.9 : 0.25,
                transition: 'width 0.4s ease, opacity 0.4s ease',
              }}
            />
          )
        })}
      </motion.div>

      {/* Big letter, centered ~42% */}
      <motion.div
        aria-hidden
        variants={letterReveal}
        className="absolute left-1/2 top-[42%] -translate-x-1/2 -translate-y-1/2 font-bebas leading-[0.78]"
        style={{
          fontSize: 'clamp(180px, 50vw, 280px)',
          letterSpacing: '-8px',
          color: palette.letter,
        }}
      >
        {card.letter}
      </motion.div>

      {/* Bottom dark block */}
      <div
        className="absolute bottom-0 left-0 right-0 px-5 pb-7 pt-5"
        style={{ backgroundColor: palette.bottomBg }}
      >
        <motion.div
          variants={fadeUp}
          className="font-bebas text-[30px] leading-none"
          style={{ color: palette.bottomFg, letterSpacing: '-0.5px' }}
        >
          {card.en}
        </motion.div>
        <motion.div
          variants={fadeUp}
          className="mt-1.5 font-noto text-[11px] font-medium"
          style={{ color: palette.bottomFg, opacity: 0.75, letterSpacing: '4px' }}
        >
          {spaced(card.ko)}
        </motion.div>
        <motion.div
          variants={fadeUp}
          className="mt-3 font-noto text-[11px] font-light leading-[1.85]"
          style={{ color: palette.bottomFg, opacity: 0.55 }}
        >
          {card.desc.map((line, i) => (
            <div key={i}>{line}</div>
          ))}
        </motion.div>
        <motion.div variants={fadeUp} className="mt-3 flex items-center gap-2">
          <span
            className="block h-[5px] w-[5px] rounded-full"
            style={{ backgroundColor: palette.bottomFg, opacity: 0.85 }}
          />
          <span
            className="font-barlow text-[10px] font-semibold"
            style={{ color: palette.bottomFg, opacity: 0.5, letterSpacing: '2px' }}
          >
            {palette.chipHex}
          </span>
        </motion.div>
      </div>
    </CardFrame>
  )
}
```

- [ ] **Step 4: Run tests, verify passing**

```bash
npm run test
```

Expected: PASS — 4 tests in CardLetter.test.tsx + 4 in cards.test.ts (8 total).

- [ ] **Step 5: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/components/CardLetter.tsx zeto/tests/CardLetter.test.tsx
git commit -m "feat(zeto): add CardLetter component with staggered reveal + tests"
```

---

### Task 10: Build CardCover (typographic cover)

**Files:**
- Create: `zeto/components/CardCover.tsx`

- [ ] **Step 1: Write the component**

```tsx
'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp } from './motion'

const LETTERS = ['Z', 'E', 'T', 'O'] as const

export function CardCover() {
  return (
    <CardFrame index={0} ariaLabel="ZETO — Zero to Art" bgClassName="bg-black">
      {/* Big ZETO centered */}
      <div
        className="absolute left-1/2 top-1/2 flex -translate-x-1/2 -translate-y-1/2 font-bebas leading-[0.78] text-white"
        style={{ fontSize: 'clamp(110px, 28vw, 200px)', letterSpacing: '-4px' }}
      >
        {LETTERS.map((l, i) => (
          <motion.span
            key={l}
            initial={{ opacity: 0, y: -50 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: false, amount: 0.5 }}
            transition={{
              delay: 0.15 + i * 0.07,
              duration: 0.7,
              ease: [0.22, 1, 0.36, 1],
            }}
          >
            {l}
          </motion.span>
        ))}
      </div>

      {/* Tagline */}
      <motion.div
        variants={fadeUp}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
        className="absolute bottom-24 left-0 right-0 text-center font-barlow text-[11px] font-semibold text-white/55"
        style={{ letterSpacing: '6px' }}
      >
        ZERO TO ART
      </motion.div>

      <motion.div
        variants={fadeUp}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
        className="absolute bottom-14 left-0 right-0 text-center font-noto text-[13px] font-medium text-white/45"
        style={{ letterSpacing: '3px' }}
      >
        제토아트
      </motion.div>

      {/* Swipe down hint */}
      <motion.div
        className="absolute bottom-5 left-1/2 -translate-x-1/2 font-bebas text-base text-white/35"
        animate={{ y: [0, 6, 0] }}
        transition={{ duration: 1.5, repeat: Infinity, ease: 'easeInOut' }}
      >
        ↓
      </motion.div>
    </CardFrame>
  )
}
```

- [ ] **Step 2: Verify TypeScript**

```bash
cd c:\Users\dhchd\work\zeto
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/components/CardCover.tsx
git commit -m "feat(zeto): add typographic cover with per-letter drop reveal"
```

---

### Task 11: Build CardClosing (4-tile + copy + colorbar)

**Files:**
- Create: `zeto/components/CardClosing.tsx`

- [ ] **Step 1: Write the component**

```tsx
'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp } from './motion'

// 4색 타일 매핑은 원본 v3 그대로 유지 (의도된 swap).
// 알파벳 카드의 primary와 다름. 수정 금지.
const TILES = [
  { letter: 'Z', bg: '#C87649' }, // Burnt Orange
  { letter: 'E', bg: '#F4E1A4' }, // Butter Lemon
  { letter: 'T', bg: '#5DA0C0' }, // Dusty Blue
  { letter: 'O', bg: '#5DA0C0' }, // Dusty Blue (intentional v3 mapping)
]

const COLOR_BAR = [
  { label: 'Zero', bg: '#FFFFFF' },
  { label: 'Explore', bg: '#F4E1A4' },
  { label: 'Thinking', bg: '#5DA0C0' },
  { label: 'Output', bg: '#C87649' },
]

export function CardClosing() {
  return (
    <CardFrame index={5} ariaLabel="제토아트 — 생각이 예술이 되는 시간" bgClassName="bg-black">
      <div className="grid h-full grid-rows-[1fr_2fr_0.6fr]">
        {/* Top: 4 color tiles */}
        <div className="grid grid-cols-4">
          {TILES.map((t, i) => (
            <motion.div
              key={t.letter + i}
              initial={{ opacity: 0, scale: 0.92 }}
              whileInView={{ opacity: 1, scale: 1 }}
              viewport={{ once: false, amount: 0.5 }}
              transition={{ delay: 0.05 * i, duration: 0.55, ease: [0.22, 1, 0.36, 1] }}
              className="flex items-center justify-center"
              style={{ backgroundColor: t.bg }}
            >
              <span
                className="font-bebas text-black"
                style={{ fontSize: 'clamp(70px, 22vw, 110px)', letterSpacing: '-2px', lineHeight: 0.85 }}
              >
                {t.letter}
              </span>
            </motion.div>
          ))}
        </div>

        {/* Middle: copy */}
        <div className="flex flex-col items-center justify-center gap-3 px-10 text-center">
          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="font-bebas text-white"
            style={{ fontSize: 'clamp(48px, 13vw, 80px)', lineHeight: 0.95, letterSpacing: '-1px' }}
          >
            생각이 <span style={{ color: '#5DA0C0' }}>예술</span>이
            <br />
            되는 시간
          </motion.div>

          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="font-bebas"
            style={{ fontSize: 'clamp(38px, 11vw, 64px)', color: '#C87649', lineHeight: 0.95, letterSpacing: '-1px' }}
          >
            제토아트
          </motion.div>

          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="font-barlow text-[11px] font-semibold text-white/40"
            style={{ letterSpacing: '4px' }}
          >
            THINKING TO ART — ZETO
          </motion.div>

          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="mt-2 border-t border-white/10 pt-3 font-noto text-[11px] font-light leading-[1.9] text-white/35"
          >
            백지 위에 서는 것을 두려워하지 마세요.
            <br />
            제토와 함께라면, 그 빈 공간이 곧 아이들의 캔버스입니다.
          </motion.div>
        </div>

        {/* Bottom: color bar */}
        <div className="grid grid-cols-4">
          {COLOR_BAR.map((c, i) => (
            <motion.div
              key={c.label}
              initial={{ opacity: 0 }}
              whileInView={{ opacity: 1 }}
              viewport={{ once: false, amount: 0.5 }}
              transition={{ delay: 0.4 + 0.05 * i, duration: 0.5 }}
              className="flex items-center justify-center"
              style={{ backgroundColor: c.bg }}
            >
              <span
                className="font-barlow font-bold text-black/55"
                style={{ fontSize: 'clamp(9px, 2.8vw, 12px)', letterSpacing: '3px', textTransform: 'uppercase' }}
              >
                {c.label}
              </span>
            </motion.div>
          ))}
        </div>
      </div>
    </CardFrame>
  )
}
```

- [ ] **Step 2: Verify TypeScript**

```bash
cd c:\Users\dhchd\work\zeto
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/components/CardClosing.tsx
git commit -m "feat(zeto): add closing card with 4-tile grid + brand copy (v3 mapping preserved)"
```

---

### Task 12: Build Deck (snap container + keyboard nav)

**Files:**
- Create: `zeto/components/Deck.tsx`

- [ ] **Step 1: Write the component**

```tsx
'use client'

import { useEffect, useRef, useState } from 'react'
import { cards } from '@/lib/cards'
import { CardCover } from './CardCover'
import { CardLetter } from './CardLetter'
import { CardClosing } from './CardClosing'

export function Deck() {
  const containerRef = useRef<HTMLDivElement>(null)
  const [activeIndex, setActiveIndex] = useState(0)

  // Track active card via IntersectionObserver (used for keyboard nav)
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

  // Keyboard navigation
  useEffect(() => {
    const root = containerRef.current
    if (!root) return

    const handler = (e: KeyboardEvent) => {
      const navKeysNext = ['ArrowDown', 'PageDown']
      const navKeysPrev = ['ArrowUp', 'PageUp']
      const isNext = navKeysNext.includes(e.key)
      const isPrev = navKeysPrev.includes(e.key)
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
    <div
      ref={containerRef}
      tabIndex={0}
      className="h-screen h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory outline-none focus:outline-none"
    >
      {cards.map((card, i) => {
        if (card.kind === 'cover') return <CardCover key="cover" />
        if (card.kind === 'closing') return <CardClosing key="closing" />
        return <CardLetter key={card.id} card={card} indexInDeck={i} />
      })}
    </div>
  )
}
```

- [ ] **Step 2: Verify TypeScript**

```bash
cd c:\Users\dhchd\work\zeto
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/components/Deck.tsx
git commit -m "feat(zeto): add Deck snap container with intersection tracking + keyboard nav"
```

---

### Task 13: Wire Deck into page

**Files:**
- Modify: `zeto/app/page.tsx`

- [ ] **Step 1: Replace `app/page.tsx`**

```tsx
import { Deck } from '@/components/Deck'

export default function Page() {
  return <Deck />
}
```

- [ ] **Step 2: Start dev server and verify full flow**

```bash
cd c:\Users\dhchd\work\zeto
npm run dev
```

Open http://localhost:3000 in a desktop browser at mobile viewport (DevTools → iPhone 14 Pro).

**Verify:**
- Cover renders black with "ZETO" centered + tagline + bouncing ↓
- Scroll: snaps to Z card (white bg, big black Z, black bottom block with "Zero / 시 작" + 3 lines)
- Continue snapping: E (amber), T (blue), O (orange), Closing (4 tiles + copy)
- Stagger reveal: number → bars → letter → word → ko → desc → chip visible on each entry
- Scrolling back up: re-reveals (once: false)
- Browser console: no errors, no warnings

If anything is off, fix and re-verify before committing.

- [ ] **Step 3: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/app/page.tsx
git commit -m "feat(zeto): mount Deck on root page — full 6-card flow live"
```

---

### Task 14: Run manual verification checklist

**Files:** none modified. Verification only.

- [ ] **Step 1: Run all automated tests**

```bash
cd c:\Users\dhchd\work\zeto
npm run test
```

Expected: all tests PASS (8 total — 4 in cards.test.ts, 4 in CardLetter.test.tsx).

- [ ] **Step 2: TypeScript check**

```bash
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 3: Production build**

```bash
npm run build
```

Expected: successful build, no errors. Note any warnings (Korean font preload warning is expected and OK).

- [ ] **Step 4: Manual browser checklist**

Start `npm run dev`, open http://localhost:3000 with DevTools mobile emulation (iPhone 14 Pro / Pixel 7).

Go through this checklist:

- [ ] Snap scroll works on mouse wheel
- [ ] Snap scroll works on touch drag (use DevTools touch emulation)
- [ ] All 6 cards reachable, snap deterministically
- [ ] Cover: ZETO 글자 drop animation 작동, ↓ 바운스 작동
- [ ] Z card: 흰 배경, 검정 Z, 검정 하단 + 흰 텍스트, "01 / 04"
- [ ] E card: 버터 옐로 배경, 다크 앰버 E, "02 / 04"
- [ ] T card: 더스티 블루, "03 / 04"
- [ ] O card: 번트 오렌지, "04 / 04", 상단 흰 잉크
- [ ] Closing: 4 타일 + 카피 + 컬러바
- [ ] Scroll up returns to previous cards, reveal triggers again
- [ ] ↓/↑/PgDn/PgUp 키보드 nav 작동 (페이지 클릭 후 포커스)
- [ ] DevTools → Rendering → "Emulate prefers-reduced-motion: reduce" 활성화 → 모션이 즉시 표시되거나 짧은 fade로 축소

- [ ] **Step 5: Lighthouse audit (mobile)**

DevTools → Lighthouse → Mode: Navigation, Device: Mobile, Categories: Performance + Accessibility.

Target: Performance ≥ 90, Accessibility ≥ 95. Record results.

If Performance < 90: check bundle size with `npm run build` output. If Accessibility < 95: read warnings and fix (likely `aria-label` missing somewhere or color contrast issue).

- [ ] **Step 6: Add prefers-reduced-motion handling if not already**

If reduced-motion check failed in Step 4: open `components/CardFrame.tsx` and `components/CardCover.tsx`, add at the top of each:

```tsx
import { useReducedMotion } from 'framer-motion'
```

And conditionally disable variants when `useReducedMotion()` returns true. Then re-run Step 4 check.

Note: Framer Motion respects `prefers-reduced-motion` automatically for transforms/opacity by default. Manual handling is only needed if the check fails.

- [ ] **Step 7: No commit (verification only)** — proceed to next task if all checks pass.

---

### Task 15: README

**Files:**
- Create: `zeto/README.md`

- [ ] **Step 1: Write README**

```markdown
# ZETO — Zero to Art

ZETO 학원 브랜드 모바일 페이지. 풀스크린 스냅 스크롤 카드뉴스 6장.

## Stack
- Next.js 15+ (App Router)
- TypeScript
- Tailwind CSS
- Framer Motion
- Vitest + Testing Library

## Development

```bash
npm install
npm run dev          # http://localhost:3000
npm run test         # Vitest
npm run build        # production build
```

## Structure
- `lib/cards.ts` — 단일 카드 데이터 소스 (카피·색상)
- `components/Deck.tsx` — 스냅 컨테이너 + 키보드 nav
- `components/CardFrame.tsx` — 공통 외피 (snap + reveal)
- `components/CardCover.tsx` / `CardLetter.tsx` / `CardClosing.tsx` — 카드별 콘텐츠
- `components/motion.ts` — 공유 Framer Motion variants

## Deploy

Vercel에 import. Root Directory = `zeto/`. Framework = Next.js 자동 감지.

## Brand colors
- Z: `#000` / `#FFF`
- E: `#F4E1A4` Butter Lemon
- T: `#5DA0C0` Dusty Blue
- O: `#C87649` Burnt Orange
```

- [ ] **Step 2: Commit**

```bash
cd c:\Users\dhchd\work
git add zeto/README.md
git commit -m "docs(zeto): add README with stack, dev commands, and deploy notes"
```

---

### Task 16: Final push + Vercel deploy

**Files:** none modified.

- [ ] **Step 1: Verify everything still passes**

```bash
cd c:\Users\dhchd\work\zeto
npm run test
npx tsc --noEmit
npm run build
```

All three must succeed before pushing.

- [ ] **Step 2: Confirm with user before push**

This step is hard-to-reverse (publishes commits + triggers Vercel deploy if connected). Ask the user:

> "All checks pass. Ready to push to origin/main and trigger Vercel deploy?"

Wait for explicit OK.

- [ ] **Step 3: Push**

```bash
cd c:\Users\dhchd\work
git push origin main
```

- [ ] **Step 4: Vercel import (user-driven)**

User logs into Vercel dashboard → New Project → import from this repo → set Root Directory = `zeto/` → Deploy.

(Or, if Vercel CLI is already authenticated locally, the user can run `cd zeto && vercel --prod`. CLI install: `npm i -g vercel`.)

- [ ] **Step 5: Smoke test the production URL**

Open the Vercel-assigned URL on a real mobile device or DevTools mobile emulation. Re-run the manual checklist from Task 14 Step 4 on production.

---

## Spec coverage check

| Spec section | Implementing task |
|--------------|-------------------|
| §2 Tech stack (Next.js, TS, Tailwind, Framer) | Tasks 1, 2, 3, 6, 7 |
| §2 Cover typographic | Task 10 |
| §2 Closing tile mapping preserved | Task 11 (with comment) |
| §3 File structure | All tasks |
| §4 Data model (cards.ts) | Task 4 |
| §5 CardFrame | Task 8 |
| §5 CardLetter | Task 9 |
| §5 CardCover | Task 10 |
| §5 CardClosing | Task 11 |
| §5 Deck + keyboard nav | Task 12 |
| §6 Motion variants | Task 5 |
| §7 Accessibility (aria, reduced-motion) | Tasks 9, 10, 11, 14 |
| §8 Responsive (clamp, 100dvh) | Tasks 7, 8, 9 |
| §9 Unit + component tests | Tasks 4, 9 |
| §9 Manual checklist + Lighthouse | Task 14 |
| §10 Deploy | Task 16 |

All spec requirements covered.
