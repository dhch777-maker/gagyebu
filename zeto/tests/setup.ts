import '@testing-library/jest-dom/vitest'

// jsdom lacks IntersectionObserver, which Framer Motion's `whileInView` requires.
// Provide a minimal no-op stub so motion components can mount during tests.
class IntersectionObserverStub {
  readonly root: Element | null = null
  readonly rootMargin: string = ''
  readonly thresholds: ReadonlyArray<number> = []
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
  takeRecords(): IntersectionObserverEntry[] {
    return []
  }
}

if (typeof globalThis.IntersectionObserver === 'undefined') {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(globalThis as any).IntersectionObserver = IntersectionObserverStub
}
