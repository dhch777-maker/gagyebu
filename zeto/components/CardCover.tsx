'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'

type Letter = 'Z' | 'E' | 'T' | 'O'

const LETTERS: readonly Letter[] = ['Z', 'E', 'T', 'O']

// Final color each letter eases into right after the drop.
const LETTER_COLOR: Record<Letter, string> = {
  Z: '#FFFFFF', // stays white
  E: '#F4E1A4', // Butter Lemon
  T: '#5DA0C0', // Dusty Blue
  O: '#C87649', // Burnt Orange
}

// Intro pause before any letter drops, so the page doesn't feel like
// it slams in the second it loads.
const INTRO = 0.9

// Letter drop start = INTRO + index*step.
const DROP_STEP = 0.13
const DROP_DURATION = 1.2

// Color tint right after each drop completes (some overlap with the
// easeOut tail so it feels like one motion).
const LETTER_COLOR_DELAY: Record<Letter, number> = {
  Z: 0, // no tint
  E: INTRO + 0.8,
  T: INTRO + 1.0,
  O: INTRO + 1.2,
}

// Swipe hint enters last — after O finishes coloring.
const HINT_DELAY = INTRO + 2.1

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
              opacity: { delay: INTRO + i * DROP_STEP, duration: DROP_DURATION, ease: [0.22, 1, 0.36, 1] },
              y: { delay: INTRO + i * DROP_STEP, duration: DROP_DURATION, ease: [0.22, 1, 0.36, 1] },
              color: { delay: LETTER_COLOR_DELAY[l], duration: 0.7, ease: 'easeInOut' },
            }}
          >
            {l}
          </motion.span>
        ))}
      </div>

      {/* Swipe hint — fades in last; bounces forever. */}
      <motion.div
        initial={{ opacity: 0 }}
        whileInView={{ opacity: 0.85 }}
        viewport={{ once: false, amount: 0.5 }}
        transition={{ delay: HINT_DELAY, duration: 0.7, ease: 'easeOut' }}
        className="absolute left-1/2 -translate-x-1/2"
        style={{ bottom: 'calc(4rem + env(safe-area-inset-bottom))' }}
      >
        <motion.div
          className="flex items-center gap-3 whitespace-nowrap text-white"
          animate={{ y: [0, 8, 0] }}
          transition={{ duration: 1.6, repeat: Infinity, ease: 'easeInOut' }}
        >
          <span
            className="font-paperlogy font-light"
            style={{ fontSize: 'clamp(14px, 3.8vw, 17px)', letterSpacing: '2px' }}
          >
            화면을 넘겨주세요
          </span>
          <span
            className="font-druk leading-none"
            style={{ fontSize: 'clamp(24px, 5.6vw, 30px)' }}
          >
            ↓
          </span>
        </motion.div>
      </motion.div>
    </CardFrame>
  )
}
