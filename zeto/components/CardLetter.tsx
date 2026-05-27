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
        className="absolute left-5 top-6 font-barlow text-[10px] font-semibold tracking-[3px]"
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

      {/* Big letter, centered ~42% */}
      <motion.div
        aria-hidden
        variants={letterReveal}
        className="absolute left-1/2 top-[42%] -translate-x-1/2 -translate-y-1/2 font-bebas leading-[0.78]"
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
        className="absolute bottom-0 left-0 right-0 px-5 pb-7 pt-5"
        style={{ backgroundColor: palette.bottomBg }}
      >
        <motion.div
          variants={fadeUp}
          className="font-bebas text-[30px] leading-none"
          style={{ color: palette.bottomFg, letterSpacing: '-0.5px' }}
        >
          {card.en}
        </motion.div>
        <motion.div
          variants={fadeUp}
          className="mt-1.5 font-noto text-[11px] font-medium"
          style={{ color: palette.bottomFg, opacity: 0.75, letterSpacing: '4px' }}
        >
          {spaced(card.ko)}
        </motion.div>
        <motion.div
          variants={fadeUp}
          className="mt-3 font-noto text-[11px] font-light leading-[1.85]"
          style={{ color: palette.bottomFg, opacity: 0.55 }}
        >
          {card.desc.map((line, i) => (
            <div key={i}>{line}</div>
          ))}
        </motion.div>
        <motion.div variants={fadeUp} className="mt-3 flex items-center gap-2">
          <span
            className="block h-[5px] w-[5px] rounded-full"
            style={{ backgroundColor: palette.bottomFg, opacity: 0.85 }}
          />
          <span
            className="font-barlow text-[10px] font-semibold"
            style={{ color: palette.bottomFg, opacity: 0.5, letterSpacing: '2px' }}
          >
            {palette.chipHex}
          </span>
        </motion.div>
      </div>
    </CardFrame>
  )
}
