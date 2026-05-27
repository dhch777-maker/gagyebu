'use client'

import { motion } from 'framer-motion'
import type { CSSProperties, ReactNode } from 'react'
import { staggerContainer } from './motion'

type Props = {
  index: number
  ariaLabel: string
  bgClassName?: string
  bgStyle?: CSSProperties
  children: ReactNode
}

export function CardFrame({ index, ariaLabel, bgClassName = '', bgStyle, children }: Props) {
  return (
    <section
      data-index={index}
      aria-label={ariaLabel}
      className={`relative h-screen h-[100dvh] w-full snap-start snap-always overflow-hidden ${bgClassName}`}
      style={bgStyle}
    >
      <motion.div
        className="relative h-full w-full"
        variants={staggerContainer}
        initial="hidden"
        whileInView="show"
        viewport={{ once: false, amount: 0.5 }}
      >
        {children}
      </motion.div>
    </section>
  )
}
