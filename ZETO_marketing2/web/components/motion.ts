import type { Variants } from 'framer-motion'

const EASE = [0.22, 1, 0.36, 1] as const

// Top-level container — controls when child variants fire.
export const staggerContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0, delayChildren: 0 } },
}

// Curtain reveal for the large cropped letter (Z/E/T/O letter cards).
// Direction is applied at call-site via `custom` prop ('down' or 'up').
export const curtainReveal: Variants = {
  hidden: (dir: 'down' | 'up' = 'down') => ({
    clipPath: dir === 'down' ? 'inset(100% 0 0 0)' : 'inset(0 0 100% 0)',
    opacity: 0.4,
    y: 8,
  }),
  show: {
    clipPath: 'inset(0 0 0 0)',
    opacity: 1,
    y: 0,
    transition: { duration: 0.75, ease: EASE },
  },
}

// Per-letter cascade for headings like "01 / ZERO".
// Used as a child variant — parent provides the stagger.
export const cascadeChar: Variants = {
  hidden: { opacity: 0, y: 12 },
  show: { opacity: 1, y: 0, transition: { duration: 0.38, ease: EASE } },
}

export const cascadeContainer: Variants = {
  hidden: {},
  show: {
    transition: { staggerChildren: 0.035, delayChildren: 0.25 },
  },
}

// Per-line fade-up for 3-line body blocks.
export const bodyLine: Variants = {
  hidden: { opacity: 0, y: 8 },
  show: { opacity: 0.85, y: 0, transition: { duration: 0.45, ease: EASE } },
}

export const bodyContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.09, delayChildren: 0.55 } },
}

// Hint bar — scales from anchor edge. transform-origin is applied via style at call-site.
export const hintBar: Variants = {
  hidden: { scaleY: 0, opacity: 0 },
  show: {
    scaleY: 1,
    opacity: 1,
    transition: { duration: 0.6, ease: EASE, delay: 0.9 },
  },
}

// Cover-specific: vertical drop-in for each of Z/E/T/O in the cover row.
export const coverLetter: Variants = {
  hidden: { opacity: 0, y: -24 },
  show: { opacity: 1, y: 0, transition: { duration: 0.4, ease: EASE } },
}

export const coverLetterContainer: Variants = {
  hidden: {},
  show: { transition: { staggerChildren: 0.12, delayChildren: 0.1 } },
}

// Simple fade-in (subhead, nav labels, etc.)
export const fadeUp: Variants = {
  hidden: { opacity: 0, y: 6 },
  show: { opacity: 1, y: 0, transition: { duration: 0.5, ease: EASE } },
}
