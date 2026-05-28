'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp } from './motion'

// 4색 타일 매핑: 사용자 지정 순서 (Z=오렌지, E=블루, T=블루, O=옐로)
const TILES = [
  { letter: 'Z', bg: '#C87649' }, // Burnt Orange
  { letter: 'E', bg: '#5DA0C0' }, // Dusty Blue
  { letter: 'T', bg: '#5DA0C0' }, // Dusty Blue
  { letter: 'O', bg: '#F4E1A4' }, // Butter Lemon
]

const COLOR_BAR = [
  { label: 'Zero', bg: '#FFFFFF' },
  { label: 'Explore', bg: '#F4E1A4' },
  { label: 'Thinking', bg: '#5DA0C0' },
  { label: 'Output', bg: '#C87649' },
]

export function CardClosing({ indexInDeck }: { indexInDeck: number }) {
  return (
    <CardFrame index={indexInDeck} ariaLabel="제토미술 — 생각이 예술이 되는 시간" bgClassName="bg-black">
      <div className="grid h-full grid-rows-[1fr_2fr_0.6fr]">
        {/* Top: 4 color tiles */}
        <div className="grid grid-cols-4">
          {TILES.map((t, i) => (
            <motion.div
              key={t.letter + i}
              initial={{ opacity: 0, scale: 0.92 }}
              whileInView={{ opacity: 1, scale: 1 }}
              viewport={{ once: false, amount: 0.5 }}
              transition={{ delay: 0.25 + 0.28 * i, duration: 1.5, ease: [0.22, 1, 0.36, 1] }}
              className="flex items-center justify-center"
              style={{ backgroundColor: t.bg }}
            >
              <span
                className="font-bebas text-black"
                style={{ fontSize: 'clamp(70px, 22vw, 110px)', letterSpacing: '-2px', lineHeight: 0.85 }}
              >
                {t.letter}
              </span>
            </motion.div>
          ))}
        </div>

        {/* Middle: copy */}
        <div className="flex flex-col items-center justify-center px-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 24 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: false, amount: 0.4 }}
            transition={{ delay: 0.4, duration: 1.3, ease: [0.22, 1, 0.36, 1] }}
            className="font-bebas text-white"
            style={{ fontSize: 'clamp(48px, 13vw, 80px)', lineHeight: 0.95, letterSpacing: '-1px' }}
          >
            <span style={{ color: '#F4E1A4' }}>생각</span>이 <span style={{ color: '#5DA0C0' }}>예술</span>이
            <br />
            되는 시간
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 24 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: false, amount: 0.4 }}
            transition={{ delay: 2.2, duration: 1.4, ease: [0.22, 1, 0.36, 1] }}
            className="mt-12 font-noto"
            style={{
              fontSize: 'clamp(68px, 19vw, 116px)',
              color: '#C87649',
              lineHeight: 0.95,
              letterSpacing: '-3px',
              fontWeight: 900,
            }}
          >
            제토미술
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 8 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: false, amount: 0.4 }}
            transition={{ delay: 3.8, duration: 0.9, ease: 'easeOut' }}
            className="mt-4 font-barlow text-[11px] font-semibold text-white/45"
            style={{ letterSpacing: '4px' }}
          >
            THINKING TO ART — ZETO
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 8 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: false, amount: 0.4 }}
            transition={{ delay: 4.4, duration: 1.0, ease: 'easeOut' }}
            className="mt-4 border-t border-white/10 pt-3 font-noto text-[11px] font-light leading-[1.9] text-white/40"
          >
            백지 위에 서는 것을 두려워하지 마세요.
            <br />
            제토와 함께라면, 그 빈 공간이 곧 아이들의 캔버스입니다.
          </motion.div>
        </div>

        {/* Bottom: color bar */}
        <div className="grid grid-cols-4">
          {COLOR_BAR.map((c, i) => (
            <motion.div
              key={c.label}
              initial={{ opacity: 0 }}
              whileInView={{ opacity: 1 }}
              viewport={{ once: false, amount: 0.5 }}
              transition={{ delay: 1.6 + 0.18 * i, duration: 1.2, ease: [0.22, 1, 0.36, 1] }}
              className="flex items-center justify-center"
              style={{ backgroundColor: c.bg }}
            >
              <span
                className="font-barlow font-bold text-black/70"
                style={{ fontSize: 'clamp(15px, 4.8vw, 20px)', letterSpacing: '3px', textTransform: 'uppercase' }}
              >
                {c.label}
              </span>
            </motion.div>
          ))}
        </div>
      </div>
    </CardFrame>
  )
}
