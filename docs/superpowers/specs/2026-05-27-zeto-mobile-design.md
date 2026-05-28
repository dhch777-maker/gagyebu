# ZETO Mobile Brand Page — Design Spec

**Date:** 2026-05-27
**Status:** Approved
**Project location:** `c:\Users\dhchd\work\zeto\`

## 1. Product summary

ZETO 학원의 브랜드 카드뉴스를 Next.js 모바일 페이지로 리메이크한다. 기존 정적 HTML(`zeto_card/zeto cardnews v3.html`)의 콘텐츠·색감·6카드 구성을 그대로 유지하되, 풀스크린 스냅 스크롤 + 프리미엄 미니멀 톤 + 스태거드 reveal 모션으로 격을 한 단계 끌어올린다. Vercel에 배포해 운영한다.

### Goals
- 모바일에서 한 카드가 화면을 꽉 채우고 스냅으로 넘겨지는 카드뉴스 경험
- "더 세련된" 인상: 대형 타이포 + 다크 하단 블록 + 진행 인디케이터 + 요소별 시차 등장
- Vercel 즉시 배포 가능한 단일 정적 Next.js 사이트

### Non-goals (YAGNI)
- 다국어, CMS 연결, 다크모드 토글
- 분석/Analytics (필요 시 향후 Vercel Analytics 한 줄 추가)
- 학원 정보 섹션(커리큘럼/위치/CTA) — 별도 작업으로 분리
- 풀 E2E 테스트 (단위 + 수동 체크리스트로 충분)

## 2. Locked design decisions

| Item | Decision |
|------|----------|
| Scope | 6 cards: Cover · Z · E · T · O · Closing |
| Flow | 풀스크린 스냅 스크롤 (`scroll-snap-type: y mandatory`) |
| Style | Premium Minimal: 대형 알파벳 + 하단 다크 블록 + 진행 인디케이터 |
| Motion | Staggered reveal (번호 → 알파벳 → 영문 word → 한글 → desc) |
| Tech | Next.js (App Router) · TypeScript · Tailwind CSS · Framer Motion |
| Cover | 타이포그래픽 (대형 "ZETO" + "Zero to Art" 태그라인) |
| Closing | 4색 타일 + 카피 구조 유지, 타이포만 업그레이드 |
| Location | `c:\Users\dhchd\work\zeto\` |

### Color palette (원본 v3와 동일)
- White / Black — Z (시작)
- `#F4E1A4` Butter Lemon — E (탐구), 배경 `#2a2209`
- `#5DA0C0` Dusty Blue — T (생각), 배경 `#1a2a38`
- `#C87649` Burnt Orange — O (작품), 배경 `#1c0d04`

### Typography
- `Bebas Neue` — 영문 대형 (알파벳, 영문 word, cover, closing 헤드)
- `Barlow Condensed` 600/700 — 라벨, 컬러바
- `Noto Sans KR` 300/500/700 — 한글 본문

`next/font/google` 로 로딩, `display: 'swap'`, `preload: true`.

## 3. File structure

```
zeto/
├── app/
│   ├── layout.tsx              # 폰트 로딩, metadata, viewport
│   ├── page.tsx                # <Deck/> 마운트
│   └── globals.css             # Tailwind base + scroll-snap container
├── components/
│   ├── Deck.tsx                # 스냅 컨테이너 (6 카드 wrap)
│   ├── CardFrame.tsx           # 공통 외피 (snap-section + reveal variants)
│   ├── CardCover.tsx           # #1 타이포그래픽 커버
│   ├── CardLetter.tsx          # #2–5 Z/E/T/O 공통 컴포넌트 (내부에 progress bars 포함)
│   ├── CardClosing.tsx         # #6 4색 타일 + 카피
│   └── motion.ts               # 공유 Framer variants
├── lib/
│   └── cards.ts                # 카드 데이터 단일 소스
├── tailwind.config.ts
├── package.json
├── tsconfig.json
└── README.md
```

## 4. Data model (`lib/cards.ts`)

```ts
export type LetterCard = {
  kind: 'letter'
  id: 'z' | 'e' | 't' | 'o'
  index: number              // 1..4 (진행 인디케이터용)
  letter: 'Z' | 'E' | 'T' | 'O'
  en: 'Zero' | 'Explore' | 'Thinking' | 'Output'
  ko: '시작' | '탐구' | '생각' | '작품'
  desc: readonly string[]    // 3 lines
  color: {
    background: string       // light: text-block 배경
    letterBg: string         // dark: letter-block 배경
    text: string             // 본문 텍스트 컬러
    chipHex: string          // "#F4E1A4 — Butter Lemon"
  }
  theme: 'dark' | 'amber' | 'blue' | 'orange'
}

export type CoverCard   = { kind: 'cover' }
export type ClosingCard = { kind: 'closing' }
export type Card = CoverCard | LetterCard | ClosingCard

export const cards: readonly Card[] = [/* 6 entries */]
```

카피·색상 변경은 이 파일만 수정하면 끝나도록 단일 소스로 유지.

## 5. Components

### CardFrame
- `<section>` with `h-[100dvh] w-full snap-start snap-always overflow-hidden relative`
- `<motion.div variants={staggerContainer} whileInView="show" viewport={{ once: false, amount: 0.5 }}>` 로 자식을 감쌈
- 60% 이상 보이면 reveal 트리거. `once: false` 로 재진입 시 다시 실행.

### CardLetter (Z/E/T/O)
화면 분할:
- 상단 (절대 위치, top 22px):
  - 좌: 번호 `02 / 04` (Inter/Barlow Condensed, letter-spacing 3px)
  - 우: 가로 진행바 4칸 (현재 인덱스만 진하게). 각 칸 `width: 8px height: 2px`, 비활성 `opacity 0.25`
- 중앙 (vertical 45% 정렬): 대형 알파벳 (`clamp(180px, 50vw, 280px)`, Bebas Neue)
- 하단 다크 블록 (절대 위치, 화면 하단, padding 18px 18px 24px):
  - 영문 word (Bebas, 28–32px)
  - 한글 (Noto Sans KR 500, 13px, letter-spacing 4px)
  - desc 3줄 (Noto Sans KR 300, 11px, line-height 1.9)
  - color chip (5px dot + hex 라벨)

### CardCover (#1)
- 풀블랙 배경
- 중앙: "ZETO" (Bebas, ~80vw, letter-spacing -4px, `#fff`). 글자별 30ms 시차 drop 모션
- 하단: "ZERO TO ART" (Barlow Condensed 600, letter-spacing 6px) + "제토아트" (Noto Sans KR 500, 작게)
- 우하단: 스와이프 다운 힌트 (↓ 바운스 애니메이션, infinite)

### CardClosing (#6)
- Grid `1fr 2fr 0.6fr`
- 상단: 4색 타일 — 원본 v3 그대로 유지 (각 알파벳 카드의 primary color와 일부 swap되어 있음. 그래픽 밸런스를 위한 의도로 보고 보존)
  - Z 타일 = `#C87649` (Burnt Orange)
  - E 타일 = `#F4E1A4` (Butter Lemon)
  - T 타일 = `#5DA0C0` (Dusty Blue)
  - O 타일 = `#5DA0C0` (Dusty Blue)
- 중앙: "생각이 *예술*이 / 되는 시간" + "제토아트" (Burnt Orange 강조) + "THINKING TO ART — ZETO" 서브 + 인용문 (인용문은 원본 v3 그대로: "백지 위에 서는 것을 두려워하지 마세요. 제토와 함께라면, 그 빈 공간이 곧 아이들의 캔버스입니다.")
- 하단: Zero/Explore/Thinking/Output 컬러바 라벨

### Deck
- 6개 카드를 map으로 렌더 (cover · z · e · t · o · closing)
- 키보드 ↓/↑/PgDn/PgUp 핸들러 (현재 카드 인덱스를 ref로 추적해 다음/이전 카드로 `scrollIntoView({ behavior: 'smooth' })`)
- 현재 인덱스 추적은 단순화: section에 `data-index` 부여하고 IntersectionObserver로 갱신. Letter 카드 내부 progress bars는 props로 전달된 정적 index만 사용하므로 cross-card 상태 동기화는 키보드 nav용으로만 필요

## 6. Motion variants (`components/motion.ts`)

```ts
import type { Variants } from 'framer-motion'

export const staggerContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.08, delayChildren: 0.1 } },
}

export const fadeUp: Variants = {
  hidden: { opacity: 0, y: 16 },
  show:   { opacity: 1, y: 0, transition: { duration: 0.55, ease: [0.22, 1, 0.36, 1] } },
}

export const letterReveal: Variants = {
  hidden: { opacity: 0, y: 40, scale: 0.96 },
  show:   { opacity: 1, y: 0,  scale: 1,    transition: { duration: 0.7, ease: [0.22, 1, 0.36, 1] } },
}
```

Reveal 순서 (CardLetter): 번호 → 진행바 → 알파벳(letterReveal) → 영문 word → 한글 → desc → chip. 전체 0.6–0.9s 안에 완결.

## 7. Accessibility

- 키보드 nav: ↓/↑/PgDn/PgUp 지원
- `useReducedMotion()` 으로 모션 축소 (0.1s fade, scale 제거)
- `<section aria-label="...">`, 장식용 큰 알파벳은 `aria-hidden="true"`
- 본문 컨트라스트 WCAG AA (원본 v3 통과)

## 8. Responsive

- Primary: 모바일 375–430px
- Tablet/Desktop ≥ 600px: 카드 `max-width: 480px` 가운데 클램프, 배경은 카드 컬러 디밍한 풀스크린
- 카드 높이는 `100dvh` 사용. 모바일 주소창 변동(접힘/펼침) 시 카드가 jump하지 않음. `dvh` 미지원 환경(구형 브라우저) 자동 폴백: Tailwind `h-screen` 클래스를 함께 두어 `100vh` 로 떨어짐
- 폰트 사이즈 모두 `clamp()`

## 9. Testing strategy

### Unit (Vitest)
- `lib/cards.ts` 데이터 무결성: 모든 letter 카드의 필수 필드 보유, 4개 정확히 존재

### Component (React Testing Library)
- `CardLetter` 가 prop을 옳게 렌더 + `aria-label` 검증

### Skipped
- 스냅 스크롤 / 모션 / IntersectionObserver E2E — 브라우저 동작 의존이라 단위 테스트 비용 ROI 낮음

### 수동 체크리스트 (배포 전)
- [ ] iPhone Safari: 스냅 동작, 100dvh, 폰트 로드
- [ ] Android Chrome: 스냅 동작
- [ ] Desktop Chrome: clamp 레이아웃, 키보드 nav
- [ ] `prefers-reduced-motion: reduce` 활성화 시 모션 축소
- [ ] Lighthouse Mobile: Performance ≥ 90, Accessibility ≥ 95

## 10. Deploy

1. `zeto/` 안에서 `create-next-app` 으로 스캐폴드 → 의존성 설치
2. 컴포넌트·데이터·스타일 구현
3. 로컬 dev 서버에서 수동 체크리스트 1차 통과
4. main 브랜치로 commit + push
5. Vercel 대시보드에서 `zeto/` 디렉토리 import (Root Directory 설정) → Production 도메인 자동 발급
