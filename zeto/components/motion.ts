import type { Variants } from 'framer-motion'

export const staggerContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.16, delayChildren: 0.2 } },
}

export const fadeUp: Variants = {
  hidden: { opacity: 0, y: 20 },
  show: {
    opacity: 1,
    y: 0,
    transition: { duration: 1.1, ease: [0.22, 1, 0.36, 1] },
  },
}

export const letterReveal: Variants = {
  hidden: { opacity: 0, y: 48, scale: 0.95 },
  show: {
    opacity: 1,
    y: 0,
    scale: 1,
    transition: { duration: 1.4, ease: [0.22, 1, 0.36, 1] },
  },
}
