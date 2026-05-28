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
