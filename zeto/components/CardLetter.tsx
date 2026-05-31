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
        className="absolute left-5 top-6 font-druk text-[10px] font-semibold tracking-[3px]"
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

      {/* Big letter, visually centered in the color area (above the dark bottom block) */}
      <motion.div
        aria-hidden
        variants={letterReveal}
        className="absolute left-1/2 top-[32%] -translate-x-1/2 -translate-y-1/2 font-druk leading-[0.78]"
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
        className="absolute bottom-0 left-0 right-0 px-6 pt-6"
        style={{
          backgroundColor: palette.bottomBg,
          paddingBottom: 'calc(2rem + env(safe-area-inset-bottom))',
        }}
      >
        <motion.div
          variants={fadeUp}
          className="font-druk leading-none"
          style={{ color: palette.bottomFg, fontSize: 'clamp(54px, 14vw, 72px)', letterSpacing: '-0.5px' }}
        >
          {card.en}
        </motion.div>
        <motion.div
          variants={fadeUp}
          className="mt-3 font-paperlogy font-medium"
          style={{ color: palette.bottomFg, opacity: 0.85, fontSize: 'clamp(18px, 4.8vw, 22px)', letterSpacing: '4px' }}
        >
          {spaced(card.ko)}
        </motion.div>
        <motion.div
          variants={fadeUp}
          className="mt-5 font-paperlogy font-light leading-[1.85]"
          style={{ color: palette.bottomFg, opacity: 0.72, fontSize: 'clamp(15px, 4vw, 18px)' }}
        >
          {card.desc.map((line, i) => (
            <div key={i}>{line}</div>
          ))}
        </motion.div>
        <motion.div variants={fadeUp} className="mt-5 flex items-center gap-2">
          <span
            className="block h-[6px] w-[6px] rounded-full"
            style={{ backgroundColor: palette.bottomFg, opacity: 0.85 }}
          />
          <span
            className="font-druk font-semibold"
            style={{ color: palette.bottomFg, opacity: 0.6, fontSize: 'clamp(11px, 3vw, 13px)', letterSpacing: '2px' }}
          >
            {palette.chipHex}
          </span>
        </motion.div>
      </div>
    </CardFrame>
  )
}
