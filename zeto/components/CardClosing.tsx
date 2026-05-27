'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { fadeUp } from './motion'

// 4색 타일 매핑은 원본 v3 그대로 유지 (의도된 swap).
// 알파벳 카드의 primary와 다름. 수정 금지.
const TILES = [
  { letter: 'Z', bg: '#C87649' }, // Burnt Orange
  { letter: 'E', bg: '#F4E1A4' }, // Butter Lemon
  { letter: 'T', bg: '#5DA0C0' }, // Dusty Blue
  { letter: 'O', bg: '#5DA0C0' }, // Dusty Blue (intentional v3 mapping)
]

const COLOR_BAR = [
  { label: 'Zero', bg: '#FFFFFF' },
  { label: 'Explore', bg: '#F4E1A4' },
  { label: 'Thinking', bg: '#5DA0C0' },
  { label: 'Output', bg: '#C87649' },
]

export function CardClosing() {
  return (
    <CardFrame index={5} ariaLabel="제토아트 — 생각이 예술이 되는 시간" bgClassName="bg-black">
      <div className="grid h-full grid-rows-[1fr_2fr_0.6fr]">
        {/* Top: 4 color tiles */}
        <div className="grid grid-cols-4">
          {TILES.map((t, i) => (
            <motion.div
              key={t.letter + i}
              initial={{ opacity: 0, scale: 0.92 }}
              whileInView={{ opacity: 1, scale: 1 }}
              viewport={{ once: false, amount: 0.5 }}
              transition={{ delay: 0.05 * i, duration: 0.55, ease: [0.22, 1, 0.36, 1] }}
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
        <div className="flex flex-col items-center justify-center gap-3 px-10 text-center">
          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="font-bebas text-white"
            style={{ fontSize: 'clamp(48px, 13vw, 80px)', lineHeight: 0.95, letterSpacing: '-1px' }}
          >
            생각이 <span style={{ color: '#5DA0C0' }}>예술</span>이
            <br />
            되는 시간
          </motion.div>

          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="font-bebas"
            style={{ fontSize: 'clamp(38px, 11vw, 64px)', color: '#C87649', lineHeight: 0.95, letterSpacing: '-1px' }}
          >
            제토아트
          </motion.div>

          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="font-barlow text-[11px] font-semibold text-white/40"
            style={{ letterSpacing: '4px' }}
          >
            THINKING TO ART — ZETO
          </motion.div>

          <motion.div
            variants={fadeUp}
            initial="hidden"
            whileInView="show"
            viewport={{ once: false, amount: 0.5 }}
            className="mt-2 border-t border-white/10 pt-3 font-noto text-[11px] font-light leading-[1.9] text-white/35"
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
              transition={{ delay: 0.4 + 0.05 * i, duration: 0.5 }}
              className="flex items-center justify-center"
              style={{ backgroundColor: c.bg }}
            >
              <span
                className="font-barlow font-bold text-black/55"
                style={{ fontSize: 'clamp(9px, 2.8vw, 12px)', letterSpacing: '3px', textTransform: 'uppercase' }}
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
