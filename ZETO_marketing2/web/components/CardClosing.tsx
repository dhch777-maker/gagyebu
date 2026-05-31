'use client'

import { motion } from 'framer-motion'
import { CardFrame } from './CardFrame'
import { bodyContainer, bodyLine, fadeUp } from './motion'

type Props = { indexInDeck: number }

const LINES = [
  'ZERO 에서 시작해',
  '깊이 보고,',
  '골똘히 생각하고,',
  '자기만의 방식으로 표현하는 우리 아이.',
  '',
  '제토는 그 과정을',
  '가장 가까이에서 함께합니다.',
  '',
  'Think to Art ,',
  '생각이 예술이 되는 시간',
] as const

export function CardClosing({ indexInDeck }: Props) {
  return (
    <CardFrame index={indexInDeck} ariaLabel="ZETO ART — closing">
      <motion.div
        variants={bodyContainer}
        className="absolute left-0 right-0 top-[28%] flex flex-col items-center text-center font-noto font-normal leading-[2] text-white"
        style={{ fontSize: 'clamp(17px, 4.4vw, 19px)' }}
      >
        {LINES.map((line, i) => (
          <motion.div key={i} variants={bodyLine} className="min-h-[1.4em]">
            {line || ' '}
          </motion.div>
        ))}
      </motion.div>

      <motion.div
        variants={fadeUp}
        transition={{ delay: 1.5, duration: 0.6, ease: [0.22, 1, 0.36, 1] }}
        className="absolute bottom-[18%] left-0 right-0 text-center font-bebas text-white"
        style={{ fontSize: 'clamp(40px, 10vw, 56px)', letterSpacing: '0px' }}
      >
        ZETO ART.
      </motion.div>
    </CardFrame>
  )
}
