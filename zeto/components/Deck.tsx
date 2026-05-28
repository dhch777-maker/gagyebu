'use client'

import { useEffect, useRef, useState } from 'react'
import { cards } from '@/lib/cards'
import { CardCover } from './CardCover'
import { CardManifesto } from './CardManifesto'
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
        if (card.kind === 'cover') return <CardCover key="cover" indexInDeck={i} />
        if (card.kind === 'manifesto') return <CardManifesto key="manifesto" indexInDeck={i} />
        if (card.kind === 'closing') return <CardClosing key="closing" indexInDeck={i} />
        return <CardLetter key={card.id} card={card} indexInDeck={i} />
      })}
    </div>
  )
}
