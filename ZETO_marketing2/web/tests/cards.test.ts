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
