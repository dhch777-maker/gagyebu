# ZETO — Zero to Art

ZETO 학원 브랜드 모바일 페이지. 풀스크린 스냅 스크롤 카드뉴스 6장.

## Stack

- Next.js 16 (App Router) · TypeScript
- Tailwind CSS v4 (`@theme` in `globals.css`)
- Framer Motion 12 (staggered reveal)
- Vitest + Testing Library

## Development

```bash
npm install
npm run dev          # http://localhost:3000
npm run test         # Vitest
npm run build        # production build
```

## Structure

- `lib/cards.ts` — 단일 카드 데이터 소스 (카피·색상 팔레트)
- `components/Deck.tsx` — 스냅 컨테이너 + 키보드 nav (↓/↑/PgDn/PgUp)
- `components/CardFrame.tsx` — 공통 외피 (snap section + Framer stagger)
- `components/CardCover.tsx` / `CardLetter.tsx` / `CardClosing.tsx` — 카드별 콘텐츠
- `components/motion.ts` — 공유 Framer Motion variants

## Brand colors

| | Color | Hex |
|---|---|---|
| Z (시작) | Black / White | `#000` / `#FFF` |
| E (탐구) | Butter Lemon | `#F4E1A4` |
| T (생각) | Dusty Blue | `#5DA0C0` |
| O (작품) | Burnt Orange | `#C87649` |

## Deploy

Vercel에 import. **Root Directory = `zeto/`**. Framework은 Next.js로 자동 감지.

또는 Vercel CLI:

```bash
npm i -g vercel
cd zeto
vercel --prod
```
