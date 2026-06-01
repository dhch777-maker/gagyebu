'use client'

import { motion } from 'framer-motion'

type Line = {
  text: 'ZERO' | 'TO' | 'ART'
  from: 'left' | 'right'
  align: 'left' | 'right'
  delay: number
}

const LINES: readonly Line[] = [
  { text: 'ZERO', from: 'left', align: 'right', delay: 0.3 },
  { text: 'TO', from: 'right', align: 'left', delay: 1.0 },
  { text: 'ART', from: 'left', align: 'right', delay: 1.7 },
]

const SLIDE_DISTANCE = 200

export function CardManifesto({ indexInDeck }: { indexInDeck: number }) {
  return (
    <section
      data-index={indexInDeck}
      aria-label="ZERO TO ART — 무에서 미로"
      className="snap-section bg-black"
    >
      {/* Zigzag stacked words */}
      <div className="absolute inset-0 flex flex-col justify-center px-8">
        {LINES.map((line) => (
          <motion.div
            key={line.text}
            initial={{
              opacity: 0,
              x: line.from === 'left' ? -SLIDE_DISTANCE : SLIDE_DISTANCE,
            }}
            whileInView={{ opacity: 1, x: 0 }}
            viewport={{ once: false, amount: 0.4 }}
            transition={{
              delay: line.delay,
              duration: 1.5,
              ease: [0.22, 1, 0.36, 1],
            }}
            className={`font-druk text-white ${
              line.align === 'right' ? 'self-end' : 'self-start'
            }`}
            style={{
              fontSize: 'clamp(96px, 28vw, 200px)',
              letterSpacing: '-5px',
              lineHeight: 0.8,
              marginTop: line.text === 'ZERO' ? 0 : '-0.04em',
            }}
          >
            {line.text}
          </motion.div>
        ))}
      </div>

      {/* Subtle tagline — tight zigzag near center */}
      <motion.div
        initial={{ opacity: 0, y: 6 }}
        whileInView={{ opacity: 1, y: 0 }}
        viewport={{ once: false, amount: 0.4 }}
        transition={{ delay: 2.5, duration: 0.9, ease: 'easeOut' }}
        className="absolute bottom-36 left-0 right-0 flex flex-col items-center gap-0.5 font-paperlogy font-light text-white/75"
        style={{ fontSize: 'clamp(22px, 6vw, 28px)', letterSpacing: '4px' }}
      >
        <span style={{ transform: 'translateX(-1.2em)' }}>
          <strong style={{ fontWeight: 900 }}>무[無]</strong>에서
        </span>
        <span style={{ transform: 'translateX(1.2em)' }}>
          <strong style={{ fontWeight: 900 }}>미[美]</strong>로
        </span>
      </motion.div>

      {/* Swipe hint — fades in last; label + arrow, bounces forever */}
      <motion.div
        initial={{ opacity: 0 }}
        whileInView={{ opacity: 0.85 }}
        viewport={{ once: false, amount: 0.5 }}
        transition={{ delay: 3.8, duration: 0.7, ease: 'easeOut' }}
        className="absolute left-1/2 -translate-x-1/2"
        style={{ bottom: 'calc(1.5rem + env(safe-area-inset-bottom))' }}
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
    </section>
  )
}
