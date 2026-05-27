'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp } from './motion'

const LETTERS = ['Z', 'E', 'T', 'O'] as const

export function CardCover() {
  return (
    <CardFrame index={0} ariaLabel="ZETO — Zero to Art" bgClassName="bg-black">
      {/* Big ZETO centered */}
      <div
        className="absolute left-1/2 top-1/2 flex -translate-x-1/2 -translate-y-1/2 font-bebas leading-[0.78] text-white"
        style={{ fontSize: 'clamp(110px, 28vw, 200px)', letterSpacing: '-4px' }}
      >
        {LETTERS.map((l, i) => (
          <motion.span
            key={l}
            initial={{ opacity: 0, y: -50 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: false, amount: 0.5 }}
            transition={{
              delay: 0.15 + i * 0.07,
              duration: 0.7,
              ease: [0.22, 1, 0.36, 1],
            }}
          >
            {l}
          </motion.span>
        ))}
      </div>

      {/* Tagline */}
      <motion.div
        variants={fadeUp}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
        className="absolute bottom-24 left-0 right-0 text-center font-barlow text-[11px] font-semibold text-white/55"
        style={{ letterSpacing: '6px' }}
      >
        ZERO TO ART
      </motion.div>

      <motion.div
        variants={fadeUp}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
        className="absolute bottom-14 left-0 right-0 text-center font-noto text-[13px] font-medium text-white/45"
        style={{ letterSpacing: '3px' }}
      >
        제토아트
      </motion.div>

      {/* Swipe down hint */}
      <motion.div
        className="absolute bottom-5 left-1/2 -translate-x-1/2 font-bebas text-base text-white/35"
        animate={{ y: [0, 6, 0] }}
        transition={{ duration: 1.5, repeat: Infinity, ease: 'easeInOut' }}
      >
        ↓
      </motion.div>
    </CardFrame>
  )
}
