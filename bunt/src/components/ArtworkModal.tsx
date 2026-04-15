'use client'

import { useState, useEffect } from 'react'
import { Artwork } from '@/types'

interface Props {
  artworks: Artwork[]
  initialIndex: number
  onClose: () => void
}

export default function ArtworkModal({ artworks, initialIndex, onClose }: Props) {
  const [index, setIndex] = useState(initialIndex)
  const artwork = artworks[index]

  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
      if (e.key === 'ArrowLeft') setIndex(i => Math.max(0, i - 1))
      if (e.key === 'ArrowRight') setIndex(i => Math.min(artworks.length - 1, i + 1))
    }
    window.addEventListener('keydown', handleKey)
    return () => window.removeEventListener('keydown', handleKey)
  }, [artworks.length, onClose])

  const dateStr = new Date(artwork.created_at).toLocaleDateString('ko-KR', {
    year: 'numeric', month: 'long',
  })

  return (
    <div
      onClick={onClose}
      style={{
        position: 'fixed', inset: 0,
        background: 'rgba(0,0,0,0.92)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        zIndex: 1000, padding: 24,
      }}
    >
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: '#111',
          border: '1px solid #2a2a2a',
          maxWidth: 640, width: '100%',
          fontFamily: 'Georgia, serif',
          position: 'relative',
        }}
      >
        <button onClick={onClose} style={{
          position: 'absolute', top: 12, right: 14,
          background: 'none', border: 'none',
          fontSize: 10, letterSpacing: 2, color: '#444',
          cursor: 'pointer', fontFamily: 'Georgia, serif',
        }}>✕ CLOSE</button>

        <div style={{
          width: '100%', aspectRatio: '4/3',
          background: '#0a0a0a',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          position: 'relative', overflow: 'hidden',
        }}>
          {artwork.image_path ? (
            <img src={artwork.image_path} alt={artwork.title}
              style={{ width: '100%', height: '100%', objectFit: 'contain' }} />
          ) : (
            <span style={{ fontSize: 80, opacity: 0.15 }}>🎨</span>
          )}
          {index > 0 && (
            <button onClick={() => setIndex(i => i - 1)} style={{
              position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)',
              background: 'rgba(0,0,0,0.7)', border: '1px solid #333',
              color: '#d4a853', fontSize: 10, letterSpacing: 1,
              padding: '8px 12px', cursor: 'pointer', fontFamily: 'Georgia, serif',
            }}>← PREV</button>
          )}
          {index < artworks.length - 1 && (
            <button onClick={() => setIndex(i => i + 1)} style={{
              position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)',
              background: 'rgba(0,0,0,0.7)', border: '1px solid #333',
              color: '#d4a853', fontSize: 10, letterSpacing: 1,
              padding: '8px 12px', cursor: 'pointer', fontFamily: 'Georgia, serif',
            }}>NEXT →</button>
          )}
        </div>

        <div style={{ padding: '20px 24px' }}>
          <h2 style={{ fontSize: 18, color: '#d4a853', letterSpacing: 3, marginBottom: 4 }}>
            {artwork.title}
          </h2>
          <p style={{ fontSize: 10, color: '#888', letterSpacing: 1 }}>{dateStr}</p>
          <div style={{ width: 30, height: 1, background: '#d4a85340', margin: '12px 0' }} />
          <p style={{ fontSize: 10, color: '#888', letterSpacing: 1 }}>
            {index + 1} / {artworks.length}
          </p>
        </div>
      </div>
    </div>
  )
}
