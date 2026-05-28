'use client'

import { motion } from 'framer-motion'
import type { ReactNode } from 'react'
import { staggerContainer } from './motion'

type Props = {
  index: number
  ariaLabel: string
  children: ReactNode
}

export function CardFrame({ index, ariaLabel, children }: Props) {
  return (
    <section
      data-index={index}
      aria-label={ariaLabel}
      className="relative h-[100dvh] w-full snap-start snap-always overflow-hidden bg-black"
    >
      <motion.div
        className="relative h-full w-full"
        variants={staggerContainer}
        initial="hidden"
        whileInView="show"
        viewport={{ once: true, amount: 0.5 }}
      >
        {children}
      </motion.div>
    </section>
  )
}
