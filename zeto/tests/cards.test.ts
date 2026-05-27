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
