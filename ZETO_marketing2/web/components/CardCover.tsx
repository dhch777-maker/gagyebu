'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { coverLetter, coverLetterContainer, fadeUp } from './motion'
import { palette } from '@/lib/cards'

type Props = { indexInDeck: number }

const LETTERS = [
  { ch: 'Z', color: palette.z },
  { ch: 'E', color: palette.e },
  { ch: 'T', color: palette.t },
  { ch: 'O', color: palette.o },
] as const

const NAV = [
  { en: 'Zero', color: palette.ink },
  { en: 'Explore', color: palette.e },
  { en: 'Thinking', color: palette.t },
  { en: 'Output', color: palette.o },
] as const

export function CardCover({ indexInDeck }: Props) {
  return (
    <CardFrame index={indexInDeck} ariaLabel="ZETO — Think to Art">
      {/* Top-right navigation list */}
      <motion.ul
        variants={coverLetterContainer}
        className="absolute right-5 top-8 flex flex-col items-end gap-1 font-noto text-[13px] font-medium tracking-wide"
      >
        {NAV.map((item) => (
          <motion.li
            key={item.en}
            variants={fadeUp}
            style={{ color: item.color }}
          >
            {item.en} ·
          </motion.li>
        ))}
      </motion.ul>

      {/* Big ZETO letters, top-left cropped */}
      <motion.div
        aria-hidden
        variants={coverLetterContainer}
        className="absolute left-[-12vw] top-[16vh] flex font-bebas leading-[0.78]"
        style={{ fontSize: 'clamp(180px, 32vw, 300px)', letterSpacing: '-6px' }}
      >
        {LETTERS.map(({ ch, color }) => (
          <motion.span key={ch} variants={coverLetter} style={{ color }}>
            {ch}
          </motion.span>
        ))}
      </motion.div>

      {/* Bottom-right THINK / TO / ART */}
      <motion.div
        variants={coverLetterContainer}
        className="absolute bottom-10 right-6 flex flex-col items-end font-bebas font-bold leading-[0.95] text-white"
        style={{ fontSize: 'clamp(56px, 14vw, 86px)', letterSpacing: '-1px' }}
      >
        {['THINK', 'TO', 'ART'].map((w) => (
          <motion.span key={w} variants={coverLetter}>
            {w}
          </motion.span>
        ))}
      </motion.div>

      <span className="sr-only">ZETO — Think to Art</span>
    </CardFrame>
  )
}
