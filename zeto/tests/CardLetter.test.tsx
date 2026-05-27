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
    expect(screen.getByText('탐 구')).toBeInTheDocument() // spaced
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
