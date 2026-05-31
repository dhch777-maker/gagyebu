import type { Metadata, Viewport } from 'next'
import { Anton } from 'next/font/google'
import './globals.css'

// English display font — Anton (free Google Fonts) as a Druk substitute.
// Druk itself is a commercial typeface; swap in licensed woff2 files here
// if a license is obtained.
const druk = Anton({
  weight: '400',
  subsets: ['latin'],
  variable: '--font-druk',
  display: 'swap',
})

// Korean font Paperlogy is loaded via @font-face in globals.css from the
// noonnu CDN (SIL OFL license).

export const metadata: Metadata = {
  title: 'ZETO — Zero to Art',
  description: '생각이 예술이 되는 시간, 제토미술',
}

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  viewportFit: 'cover',
  themeColor: '#000000',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko" className={druk.variable}>
      <body className="antialiased bg-black text-white">{children}</body>
    </html>
  )
}
