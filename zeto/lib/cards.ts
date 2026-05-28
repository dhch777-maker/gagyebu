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
export type ManifestoCard = { kind: 'manifesto' }
export type ClosingCard = { kind: 'closing' }
export type Card = CoverCard | ManifestoCard | LetterCard | ClosingCard

const BLACK = '#000000'
const WHITE = '#FFFFFF'

export const cards: readonly Card[] = [
  { kind: 'cover' },
  { kind: 'manifesto' },
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
