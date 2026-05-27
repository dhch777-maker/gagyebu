import type { Metadata, Viewport } from 'next'
import { Bebas_Neue, Barlow_Condensed, Noto_Sans_KR } from 'next/font/google'
import './globals.css'

const bebas = Bebas_Neue({
  weight: '400',
  subsets: ['latin'],
  variable: '--font-bebas',
  display: 'swap',
})

const barlow = Barlow_Condensed({
  weight: ['600', '700'],
  subsets: ['latin'],
  variable: '--font-barlow',
  display: 'swap',
})

const notoKR = Noto_Sans_KR({
  weight: ['300', '500', '700'],
  subsets: ['latin'],
  variable: '--font-noto',
  display: 'swap',
  preload: false,
})

export const metadata: Metadata = {
  title: 'ZETO — Zero to Art',
  description: '생각이 예술이 되는 시간, 제토아트',
}

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  themeColor: '#000000',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko" className={`${bebas.variable} ${barlow.variable} ${notoKR.variable}`}>
      <body className="antialiased bg-black text-white">{children}</body>
    </html>
  )
}
