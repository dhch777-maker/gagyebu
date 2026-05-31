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
  // The big letter sits visibly INSIDE the corner with only a small bleed.
  // line-height 0.85 ≈ Bebas cap height, so the element box hugs the glyph
  // and the offsets behave intuitively.
  const base: React.CSSProperties = {
    position: 'absolute',
    fontSize: 'clamp(280px, 80vw, 480px)',
    letterSpacing: '-8px',
    lineHeight: 0.85,
    fontWeight: 700,
  }
  switch (anchor) {
    case 'top-left':     return { ...base, top: '-3vh', left: '-3vw' }
    case 'top-right':    return { ...base, top: '-3vh', right: '-3vw' }
    case 'bottom-left':  return { ...base, bottom: '-3vh', left: '-3vw' }
    case 'bottom-right': return { ...base, bottom: '-3vh', right: '-3vw' }
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
