'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp } from './motion'

type Letter = 'Z' | 'E' | 'T' | 'O'

const LETTERS: readonly Letter[] = ['Z', 'E', 'T', 'O']

// Final color each letter eases into after the drop completes.
const LETTER_COLOR: Record<Letter, string> = {
  Z: '#FFFFFF', // stays white
  E: '#F4E1A4', // Butter Lemon
  T: '#5DA0C0', // Dusty Blue
  O: '#C87649', // Burnt Orange
}

// Color-tint kicks in one letter at a time AFTER all four have dropped.
const LETTER_COLOR_DELAY: Record<Letter, number> = {
  Z: 0, // no tint
  E: 2.2,
  T: 2.9,
  O: 3.6,
}

export function CardCover({ indexInDeck }: { indexInDeck: number }) {
  return (
    <CardFrame index={indexInDeck} ariaLabel="ZETO — Zero to Art" bgClassName="bg-black">
      {/* Big ZETO centered */}
      <div
        className="absolute left-1/2 top-1/2 flex -translate-x-1/2 -translate-y-1/2 font-druk leading-[0.78]"
        style={{ fontSize: 'clamp(110px, 28vw, 200px)', letterSpacing: '-4px' }}
      >
        {LETTERS.map((l, i) => (
          <motion.span
            key={l}
            initial={{ opacity: 0, y: -50, color: '#FFFFFF' }}
            whileInView={{ opacity: 1, y: 0, color: LETTER_COLOR[l] }}
            viewport={{ once: false, amount: 0.5 }}
            transition={{
              opacity: { delay: 0.25 + i * 0.13, duration: 1.2, ease: [0.22, 1, 0.36, 1] },
              y: { delay: 0.25 + i * 0.13, duration: 1.2, ease: [0.22, 1, 0.36, 1] },
              color: { delay: LETTER_COLOR_DELAY[l], duration: 1.1, ease: 'easeInOut' },
            }}
          >
            {l}
          </motion.span>
        ))}
      </div>

      <motion.div
        variants={fadeUp}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
        className="absolute bottom-16 left-0 right-0 text-center font-paperlogy text-[14px] font-medium text-white/55"
        style={{ letterSpacing: '4px' }}
      >
        제토아트
      </motion.div>

      {/* Swipe down hint — text + arrow, both larger and brighter */}
      <motion.div
        className="absolute left-1/2 -translate-x-1/2 flex items-center gap-3 whitespace-nowrap text-white/75"
        style={{ bottom: 'calc(1.25rem + env(safe-area-inset-bottom))' }}
        animate={{ y: [0, 8, 0] }}
        transition={{ duration: 1.6, repeat: Infinity, ease: 'easeInOut' }}
      >
        <span
          className="font-paperlogy font-light"
          style={{ fontSize: 'clamp(14px, 3.6vw, 16px)', letterSpacing: '2px' }}
        >
          화면을 넘겨주세요
        </span>
        <span
          className="font-druk leading-none"
          style={{ fontSize: 'clamp(22px, 5vw, 26px)' }}
        >
          ↓
        </span>
      </motion.div>
    </CardFrame>
  )
}
