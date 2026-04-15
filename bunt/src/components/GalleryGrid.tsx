'use client'

import { useState } from 'react'
import { Artwork } from '@/types'
import ArtworkModal from './ArtworkModal'

const PLACEHOLDER_COLORS = [
  '#1a1410', '#12100e', '#1c1814', '#141210', '#1e1a16',
  '#100e0c', '#181410', '#1a1612', '#161210',
]

interface Props { artworks: Artwork[] }

export default function GalleryGrid({ artworks }: Props) {
  const [selectedIndex, setSelectedIndex] = useState<number | null>(null)
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null)

  return (
    <>
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(3, 1fr)',
        gap: 8,
      }}>
        {artworks.map((artwork, index) => (
          <div
            key={artwork.id}
            onClick={() => setSelectedIndex(index)}
            onMouseEnter={() => setHoveredIndex(index)}
            onMouseLeave={() => setHoveredIndex(null)}
            style={{
              aspectRatio: '1',
              background: artwork.image_path ? 'transparent' : PLACEHOLDER_COLORS[index % PLACEHOLDER_COLORS.length],
              border: '1px solid #1e1e1e',
              cursor: 'pointer',
              position: 'relative',
              overflow: 'hidden',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
          >
            {artwork.image_path ? (
              <img
                src={artwork.image_path}
                alt={artwork.title}
                style={{ width: '100%', height: '100%', objectFit: 'cover' }}
              />
            ) : (
              <span style={{ fontSize: 36, opacity: 0.2 }}>🎨</span>
            )}

            {/* 신규 작품 황동 점 (7일 이내) */}
            {new Date(artwork.created_at) > new Date(Date.now() - 7 * 24 * 60 * 60 * 1000) && (
              <div style={{
                position: 'absolute', top: 8, right: 8,
                width: 8, height: 8,
                background: '#d4a853', borderRadius: '50%',
              }} />
            )}

            {/* 호버 오버레이 */}
            <div style={{
              position: 'absolute', inset: 0,
              background: 'rgba(0,0,0,0.65)',
              opacity: hoveredIndex === index ? 1 : 0,
              transition: 'opacity 0.2s',
              display: 'flex',
              alignItems: 'flex-end',
              padding: '10px',
            }}>
              <span style={{ fontSize: 11, color: '#d4a853', letterSpacing: 1 }}>
                {artwork.title}
              </span>
            </div>
          </div>
        ))}
      </div>

      {selectedIndex !== null && (
        <ArtworkModal
          artworks={artworks}
          initialIndex={selectedIndex}
          onClose={() => setSelectedIndex(null)}
        />
      )}
    </>
  )
}
